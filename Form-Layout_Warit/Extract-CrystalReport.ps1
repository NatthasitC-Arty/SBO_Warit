# ============================================================
# Extract-CrystalReport.ps1
# Extracts SQL Command, Parameters, Formulas, Connection Info,
# Table Links, Record Selection from all .rpt files
# using Crystal Reports SDK (GAC v13.0.4000.0)
# ============================================================

param(
    [string]$RptFolder = "C:\SDA\Form_SDA_MARK01",
    [string]$OutputRoot = "C:\SDA\ExtractedSQL",
    [switch]$FlatOutput
)

# --- Load Crystal Reports Assemblies from GAC ---
try {
    [System.Reflection.Assembly]::Load("CrystalDecisions.CrystalReports.Engine, Version=13.0.4000.0, Culture=neutral, PublicKeyToken=692fbea5521e1304") | Out-Null
    [System.Reflection.Assembly]::Load("CrystalDecisions.Shared, Version=13.0.4000.0, Culture=neutral, PublicKeyToken=692fbea5521e1304") | Out-Null
    [System.Reflection.Assembly]::Load("CrystalDecisions.ReportAppServer.ClientDoc, Version=13.0.4000.0, Culture=neutral, PublicKeyToken=692fbea5521e1304") | Out-Null
    [System.Reflection.Assembly]::Load("CrystalDecisions.ReportAppServer.DataDefModel, Version=13.0.4000.0, Culture=neutral, PublicKeyToken=692fbea5521e1304") | Out-Null
    Write-Host "[OK] Crystal Reports assemblies loaded." -ForegroundColor Green
}
catch {
    Write-Host "[ERROR] Failed to load Crystal Reports assemblies: $_" -ForegroundColor Red
    exit 1
}

# --- Helper: Extract SQL Command Text from DatabaseController ---
function Get-SqlCommandText {
    param($ReportClientDocument, [string]$ReportName)

    $sqlTexts = @()
    try {
        $dbController = $ReportClientDocument.DatabaseController
        $database = $dbController.Database
        $tables = $database.Tables

        for ($i = 0; $i -lt $tables.Count; $i++) {
            $table = $tables[$i]
            # Prefer Alias (the name shown in Selected Tables / Database Expert)
            # and fall back to Name when no alias is defined.
            $tableName = $null
            try { $tableName = $table.Alias } catch {}
            if ([string]::IsNullOrWhiteSpace($tableName)) { $tableName = $table.Name }
            $commandText = ""

            # Try to get CommandText via reflection (COM interop)
            try {
                $commandText = $table.GetType().InvokeMember(
                    "CommandText",
                    [System.Reflection.BindingFlags]::GetProperty,
                    $null, $table, $null
                )
            }
            catch {
                # Not a command table, try QualifiedName
                try {
                    $commandText = "-- Table: $tableName (QualifiedName: $($table.QualifiedName))"
                }
                catch {
                    $commandText = "-- Table: $tableName (no CommandText)"
                }
            }

            if ($commandText -and $commandText.Trim().Length -gt 0) {
                $sqlTexts += @{
                    Report = $ReportName
                    Table  = $tableName
                    SQL    = $commandText
                }
            }
        }
    }
    catch {
        $sqlTexts += @{
            Report = $ReportName
            Table  = "(error)"
            SQL    = "-- Error reading tables: $_"
        }
    }

    return $sqlTexts
}

# --- Helper: Extract Parameters ---
function Get-Parameters {
    param($ReportDocument)

    $params = @()
    try {
        $paramFields = $ReportDocument.DataDefinition.ParameterFields
        foreach ($p in $paramFields) {
            $params += [PSCustomObject]@{
                Name          = $p.ParameterFieldName
                Type          = $p.ParameterValueType.ToString()
                ReportName    = $p.ReportName
                DefaultValues = try { ($p.DefaultValues | ForEach-Object { $_.Description }) -join ", " } catch { "" }
            }
        }
    }
    catch {
        $params += [PSCustomObject]@{ Name = "(error)"; Type = "$_"; ReportName = ""; DefaultValues = "" }
    }
    return $params
}

# --- Helper: Extract Formula Fields ---
function Get-FormulaFields {
    param($ReportDocument)

    $formulas = @()
    try {
        foreach ($f in $ReportDocument.DataDefinition.FormulaFields) {
            $formulas += [PSCustomObject]@{
                Name = $f.FormulaName
                Text = $f.Text
            }
        }
    }
    catch {
        $formulas += [PSCustomObject]@{ Name = "(error)"; Text = "$_" }
    }
    return $formulas
}

# --- Helper: Extract Connection Info ---
function Get-ConnectionInfo {
    param($ReportDocument)

    $connections = @()
    try {
        foreach ($tbl in $ReportDocument.Database.Tables) {
            $logon = $tbl.LogOnInfo.ConnectionInfo
            $connections += [PSCustomObject]@{
                TableName    = $tbl.Name
                ServerName   = $logon.ServerName
                DatabaseName = $logon.DatabaseName
                UserID       = $logon.UserID
            }
        }
    }
    catch {
        $connections += [PSCustomObject]@{ TableName = "(error)"; ServerName = "$_"; DatabaseName = ""; UserID = "" }
    }
    return $connections
}

# --- Helper: Extract Table Links ---
function Get-TableLinks {
    param($ReportDocument)

    $links = @()
    try {
        foreach ($link in $ReportDocument.Database.Links) {
            $links += [PSCustomObject]@{
                SourceTable = $link.SourceTable.Name
                SourceField = $link.SourceFields[0].Name
                DestTable   = $link.DestinationTable.Name
                DestField   = $link.DestinationFields[0].Name
                JoinType    = $link.JoinType.ToString()
            }
        }
    }
    catch {
        $links += [PSCustomObject]@{ SourceTable = "(error)"; SourceField = "$_"; DestTable = ""; DestField = ""; JoinType = "" }
    }
    return $links
}

# --- Helper: Extract Record Selection Formula ---
function Get-RecordSelection {
    param($ReportDocument)

    try {
        return $ReportDocument.RecordSelectionFormula
    }
    catch {
        return "-- Error: $_"
    }
}

# --- Main Processing ---
$rptFiles = Get-ChildItem -Path $RptFolder -Recurse -Filter "*.rpt"
$totalFiles = $rptFiles.Count
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  Crystal Report SQL Extractor" -ForegroundColor Cyan
Write-Host "  Found $totalFiles .rpt files" -ForegroundColor Cyan
Write-Host "  Source: $RptFolder" -ForegroundColor Cyan
Write-Host "  Output: $OutputRoot" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

# Create output root
if (!(Test-Path $OutputRoot)) {
    New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
}

$counter = 0
$errors = @()

foreach ($rptFile in $rptFiles) {
    $counter++
    $relativePath = $rptFile.FullName.Substring($RptFolder.Length).TrimStart('\', '/')
    $relativeDir = Split-Path $relativePath -Parent
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($rptFile.Name)

    Write-Host "[$counter/$totalFiles] Processing: $relativePath" -ForegroundColor Yellow

    # Create matching output subfolder
    $outputDir = Join-Path $OutputRoot $relativeDir
    if (!(Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }

    if ($FlatOutput) {
        # Write all files directly into $outputDir, prefixed with the report's basename
        $reportDir = $outputDir
        $filePrefix = "${baseName}_"
    }
    else {
        # Create a subfolder per report for individual .sql files
        $reportDir = Join-Path $outputDir $baseName
        if (!(Test-Path $reportDir)) {
            New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
        }
        $filePrefix = ""
    }

    $txtFile = Join-Path $reportDir "$baseName.txt"
    $sbTxt = [System.Text.StringBuilder]::new()
    $sqlFileCount = 0
    $usedSqlNames = @{}

    $rd = $null
    try {
        # Load report
        $rd = New-Object CrystalDecisions.CrystalReports.Engine.ReportDocument
        $rd.Load($rptFile.FullName)

        $header = "Report: $($rptFile.Name)`r`nPath:   $relativePath`r`nExtracted: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

        # ===================== Individual .sql files =====================

        # --- Main Report SQL Commands ---
        try {
            $rcd = $rd.ReportClientDocument
            $sqlResults = Get-SqlCommandText -ReportClientDocument $rcd -ReportName "Main"

            $mainIdx = 0
            foreach ($sql in $sqlResults) {
                # Skip non-command entries (lines starting with "-- Table:")
                if ($sql.SQL -match "^-- Table:") { continue }

                $mainIdx++
                $safeTblName = ($sql.Table -replace '[\\/:*?"<>|]', '_').Trim()
                $rawName = "${filePrefix}${safeTblName}"
                $sqlFileName = "$rawName.sql"
                $dupIdx = 2
                while ($usedSqlNames.ContainsKey($sqlFileName)) {
                    $sqlFileName = "${rawName}_${dupIdx}.sql"
                    $dupIdx++
                }
                $usedSqlNames[$sqlFileName] = $true
                $sqlFilePath = Join-Path $reportDir $sqlFileName

                $content = "-- ============================================================`r`n"
                $content += "-- $header`r`n"
                $content += "-- Source: Main Report`r`n"
                $content += "-- Table:  $($sql.Table)`r`n"
                $content += "-- ============================================================`r`n`r`n"
                $content += $sql.SQL

                $content | Out-File -FilePath $sqlFilePath -Encoding UTF8 -Force
                $sqlFileCount++
                Write-Host "  -> SQL [$sqlFileCount]: $sqlFileName" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "  [WARN] Error reading main report SQL: $_" -ForegroundColor DarkYellow
        }

        # --- Subreport SQL Commands ---
        try {
            $subreportNames = $rd.Subreports | ForEach-Object { $_.Name }
            if ($subreportNames -and $subreportNames.Count -gt 0) {
                foreach ($subName in $subreportNames) {
                    try {
                        $rcd = $rd.ReportClientDocument
                        $subRcd = $rcd.SubreportController.GetSubreport($subName)
                        $subSqlResults = Get-SqlCommandText -ReportClientDocument $subRcd -ReportName $subName

                        $subIdx = 0
                        foreach ($sql in $subSqlResults) {
                            if ($sql.SQL -match "^-- Table:") { continue }

                            $subIdx++
                            $safeSubName = ($subName -replace '[\\/:*?"<>|]', '_').Trim()
                            $safeTblName = ($sql.Table -replace '[\\/:*?"<>|]', '_').Trim()
                            $rawName = "${filePrefix}${safeSubName}_${safeTblName}"
                            $sqlFileName = "$rawName.sql"
                            $dupIdx = 2
                            while ($usedSqlNames.ContainsKey($sqlFileName)) {
                                $sqlFileName = "${rawName}_${dupIdx}.sql"
                                $dupIdx++
                            }
                            $usedSqlNames[$sqlFileName] = $true
                            $sqlFilePath = Join-Path $reportDir $sqlFileName

                            $content = "-- ============================================================`r`n"
                            $content += "-- $header`r`n"
                            $content += "-- Source: Subreport [$subName]`r`n"
                            $content += "-- Table:  $($sql.Table)`r`n"
                            $content += "-- ============================================================`r`n`r`n"
                            $content += $sql.SQL

                            $content | Out-File -FilePath $sqlFilePath -Encoding UTF8 -Force
                            $sqlFileCount++
                            Write-Host "  -> SQL [$sqlFileCount]: $sqlFileName" -ForegroundColor Green
                        }
                    }
                    catch {
                        Write-Host "  [WARN] Error reading subreport '$subName': $_" -ForegroundColor DarkYellow
                    }
                }
            }
        }
        catch {
            Write-Host "  [WARN] Error enumerating subreports: $_" -ForegroundColor DarkYellow
        }

        if ($sqlFileCount -eq 0) {
            Write-Host "  -> (No SQL commands found)" -ForegroundColor DarkGray
        }

        # ===================== .txt file =====================
        [void]$sbTxt.AppendLine("============================================================")
        [void]$sbTxt.AppendLine($header)
        [void]$sbTxt.AppendLine("============================================================")
        [void]$sbTxt.AppendLine("")

        # --- 2. Parameters ---
        [void]$sbTxt.AppendLine("************************************************************")
        [void]$sbTxt.AppendLine("  PARAMETERS")
        [void]$sbTxt.AppendLine("************************************************************")

        $params = Get-Parameters -ReportDocument $rd
        if ($params.Count -eq 0) {
            [void]$sbTxt.AppendLine("(No parameters)")
        }
        else {
            foreach ($p in $params) {
                [void]$sbTxt.AppendLine("Parameter: $($p.Name)")
                [void]$sbTxt.AppendLine("  Type:     $($p.Type)")
                [void]$sbTxt.AppendLine("  Report:   $($p.ReportName)")
                [void]$sbTxt.AppendLine("  Default:  $($p.DefaultValues)")
                [void]$sbTxt.AppendLine("")
            }
        }
        [void]$sbTxt.AppendLine("")

        # --- 3. Formula Fields ---
        [void]$sbTxt.AppendLine("************************************************************")
        [void]$sbTxt.AppendLine("  FORMULA FIELDS")
        [void]$sbTxt.AppendLine("************************************************************")

        $formulas = Get-FormulaFields -ReportDocument $rd
        if ($formulas.Count -eq 0) {
            [void]$sbTxt.AppendLine("(No formula fields)")
        }
        else {
            foreach ($f in $formulas) {
                [void]$sbTxt.AppendLine("Formula: $($f.Name)")
                [void]$sbTxt.AppendLine("--- BEGIN ---")
                [void]$sbTxt.AppendLine($f.Text)
                [void]$sbTxt.AppendLine("--- END ---")
                [void]$sbTxt.AppendLine("")
            }
        }
        [void]$sbTxt.AppendLine("")

        # --- 4. Connection Info ---
        [void]$sbTxt.AppendLine("************************************************************")
        [void]$sbTxt.AppendLine("  CONNECTION INFO")
        [void]$sbTxt.AppendLine("************************************************************")

        $connections = Get-ConnectionInfo -ReportDocument $rd
        if ($connections.Count -eq 0) {
            [void]$sbTxt.AppendLine("(No connection info)")
        }
        else {
            foreach ($c in $connections) {
                [void]$sbTxt.AppendLine("Table:    $($c.TableName)")
                [void]$sbTxt.AppendLine("  Server:   $($c.ServerName)")
                [void]$sbTxt.AppendLine("  Database: $($c.DatabaseName)")
                [void]$sbTxt.AppendLine("  UserID:   $($c.UserID)")
                [void]$sbTxt.AppendLine("")
            }
        }
        [void]$sbTxt.AppendLine("")

        # --- 5. Table Links ---
        [void]$sbTxt.AppendLine("************************************************************")
        [void]$sbTxt.AppendLine("  TABLE LINKS")
        [void]$sbTxt.AppendLine("************************************************************")

        $links = Get-TableLinks -ReportDocument $rd
        if ($links.Count -eq 0) {
            [void]$sbTxt.AppendLine("(No table links)")
        }
        else {
            foreach ($l in $links) {
                [void]$sbTxt.AppendLine("$($l.SourceTable).$($l.SourceField) --[$($l.JoinType)]--> $($l.DestTable).$($l.DestField)")
            }
        }
        [void]$sbTxt.AppendLine("")

        # --- 6. Record Selection Formula ---
        [void]$sbTxt.AppendLine("************************************************************")
        [void]$sbTxt.AppendLine("  RECORD SELECTION FORMULA")
        [void]$sbTxt.AppendLine("************************************************************")

        $recSel = Get-RecordSelection -ReportDocument $rd
        if ([string]::IsNullOrWhiteSpace($recSel)) {
            [void]$sbTxt.AppendLine("(No record selection formula)")
        }
        else {
            [void]$sbTxt.AppendLine($recSel)
        }
        [void]$sbTxt.AppendLine("")

        # Write .txt file
        $sbTxt.ToString() | Out-File -FilePath $txtFile -Encoding UTF8 -Force
        Write-Host "  -> TXT: $txtFile" -ForegroundColor Green

    }
    catch {
        $errMsg = "Error processing '$relativePath': $_"
        Write-Host "  [ERROR] $errMsg" -ForegroundColor Red
        $errors += $errMsg

        # Write error to output file
        $errFile = Join-Path $reportDir "${filePrefix}_ERROR.txt"
        "ERROR processing this report`n$errMsg" | Out-File -FilePath $errFile -Encoding UTF8 -Force
    }
    finally {
        if ($rd) {
            try { $rd.Close() } catch {}
            try { $rd.Dispose() } catch {}
        }
    }
}

# --- Summary ---
Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  DONE! Processed $counter / $totalFiles files" -ForegroundColor Cyan
if ($errors.Count -gt 0) {
    Write-Host "  Errors: $($errors.Count)" -ForegroundColor Red
    foreach ($e in $errors) {
        Write-Host "    - $e" -ForegroundColor Red
    }
}
else {
    Write-Host "  No errors!" -ForegroundColor Green
}
Write-Host "  Output folder: $OutputRoot" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
