# ============================================================
# Export current FMS (UDV) to CSV in UDV_Map.csv format.
# Uses .NET SqlClient direct SQL (no DI API dependency).
# Reads CSHS + OUQR + SHS1 and writes a multi-row CSV
# (1 row per trigger field, grouped by FMS key).
# Output is round-trip safe: re-import with Action=UPSERT.
# ============================================================
param(
    [string]$Server      = "SLD-C072",
    [string]$CompanyDB   = "SBO_SDA",
    [string]$DBUser      = "sa",
    [string]$DBPassword  = "1q2w3e4r",
    [string]$OutFile     = "",
    [string]$LogFile     = "",
    [string]$ExportAction = "UPSERT"
)

if (-not $OutFile) {
    $OutFile = Join-Path $PSScriptRoot ("..\Config\UDV_Export_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}
if (-not $LogFile) {
    $LogFile = Join-Path $PSScriptRoot "..\Export_UDV_Log.txt"
}

# ------------------------------------------------------------
# Logging
# ------------------------------------------------------------
function Write-Log {
    param([string]$Msg, [string]$Level = "INFO")
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Msg
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

function To-YN {
    param($Val)
    $s = ([string]$Val).Trim().ToUpper()
    if ($s -in @("Y","YES","TRUE","1")) { "Y" } else { "N" }
}

# ------------------------------------------------------------
# Helper: read column from SqlDataReader by trying multiple names.
# Returns null if no match. Treats DBNull as null.
# ------------------------------------------------------------
function Read-ReaderByNames {
    param($Reader, $ColMap, [string[]]$Names)
    foreach ($n in $Names) {
        if ($ColMap.Contains($n)) {
            $v = $Reader[$n]
            if ($null -ne $v -and -not ($v -is [DBNull])) { return $v }
        }
    }
    return $null
}

# ------------------------------------------------------------
# Parse QueryBody for $[$...] tokens to derive trigger fields
# (fallback when SHS1 is empty / unavailable).
# ------------------------------------------------------------
function Get-QueryTriggers {
    param([string]$QueryBody)
    $list = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($QueryBody)) { return $list }

    $rx1 = [regex]'\$\[\$\d+\.([^.\]]+)'
    foreach ($m in $rx1.Matches($QueryBody)) {
        $f = $m.Groups[1].Value.Trim()
        if ($f -and (-not $list.Contains($f))) { [void]$list.Add($f) }
    }
    $rx2 = [regex]'\$\[([A-Za-z][A-Za-z0-9_]*)\.([^\]]+)\]'
    foreach ($m in $rx2.Matches($QueryBody)) {
        $f = $m.Groups[2].Value.Trim()
        if ($f -and (-not $list.Contains($f))) { [void]$list.Add($f) }
    }
    return $list
}

Write-Log "=== Start UDV/FMS Export (SQL Direct) ==="
Write-Log "Server   : $Server"
Write-Log "CompanyDB: $CompanyDB"
Write-Log "OutFile  : $OutFile"

# ------------------------------------------------------------
# Open SQL connection (no DI API needed)
# ------------------------------------------------------------
$connStr = "Server=$Server;Database=$CompanyDB;User ID=$DBUser;Password=$DBPassword;Connection Timeout=10"
$conn = New-Object System.Data.SqlClient.SqlConnection $connStr
try {
    $conn.Open()
    Write-Log "Connected to SQL: $Server / $CompanyDB"

    # ------------------------------------------------------------
    # 1) Discover CSHS column names (some columns are version-specific)
    # ------------------------------------------------------------
    $cshsCols = New-Object System.Collections.Generic.HashSet[string]
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='CSHS' ORDER BY ORDINAL_POSITION"
    $rdr = $cmd.ExecuteReader()
    while ($rdr.Read()) { [void]$cshsCols.Add([string]$rdr["COLUMN_NAME"]) }
    $rdr.Close()
    $cshsColList = ($cshsCols | Sort-Object) -join ", "
    Write-Log "CSHS columns ($($cshsCols.Count)): $cshsColList"

    # ------------------------------------------------------------
    # 2) Load OUQR queries into hashtable[IntrnalKey]
    # ------------------------------------------------------------
    $queries = @{}
    $cmd.CommandText = "SELECT IntrnalKey, QCategory, QName, QString FROM OUQR"
    $rdr = $cmd.ExecuteReader()
    while ($rdr.Read()) {
        $key = [int]$rdr["IntrnalKey"]
        $queries[$key] = @{
            Category = [int]$rdr["QCategory"]
            Name     = [string]$rdr["QName"]
            Body     = if ($rdr["QString"] -is [DBNull]) { "" } else { [string]$rdr["QString"] }
        }
    }
    $rdr.Close()
    Write-Log "Loaded $($queries.Count) saved queries from OUQR"

    # ------------------------------------------------------------
    # 2b) Load SHS1 trigger field list per CSHS IndexID
    # SHS1 (IndexID int, FieldID nvarchar) — canonical trigger storage
    # ------------------------------------------------------------
    $shs1Map = @{}
    try {
        $cmd.CommandText = "SELECT IndexID, FieldID FROM SHS1 ORDER BY IndexID"
        $rdr = $cmd.ExecuteReader()
        while ($rdr.Read()) {
            $iid = [int]$rdr["IndexID"]
            $fld = if ($rdr["FieldID"] -is [DBNull]) { "" } else { [string]$rdr["FieldID"] }
            if (-not $shs1Map.ContainsKey($iid)) { $shs1Map[$iid] = New-Object System.Collections.Generic.List[string] }
            if ($fld) { [void]$shs1Map[$iid].Add($fld) }
        }
        $rdr.Close()
        Write-Log "Loaded SHS1 triggers for $($shs1Map.Count) CSHS row(s)"
    } catch {
        Write-Log "SHS1 not available: $($_.Exception.Message)" "WARN"
    }

    # ------------------------------------------------------------
    # 3) Read all CSHS rows and emit multi-row CSV
    # ------------------------------------------------------------
    $cmd.CommandText = "SELECT * FROM CSHS ORDER BY FormID, ItemID, ColID"
    $rdr = $cmd.ExecuteReader()
    $output = New-Object System.Collections.ArrayList
    $rowNo = 0
    while ($rdr.Read()) {
        $rowNo++

        $indexId  = [int](Read-ReaderByNames $rdr $cshsCols @("IndexID"))
        $formId   = [string](Read-ReaderByNames $rdr $cshsCols @("FormID"))
        $itemId   = [string](Read-ReaderByNames $rdr $cshsCols @("ItemID"))
        $colId    = [string](Read-ReaderByNames $rdr $cshsCols @("ColID","ColumnID"))
        # v10 CSHS: ActionT (numeric 2=Q, 0=F), QueryId, FrceRfrsh, ByField, FieldID
        $actionT  = ([string](Read-ReaderByNames $rdr $cshsCols @("ActionT","Action","ActionType"))).Trim()
        $queryId  = Read-ReaderByNames $rdr $cshsCols @("QueryId","QueryID")
        $fixedVal = [string](Read-ReaderByNames $rdr $cshsCols @("StringVal","StringValue","FixedValue","DefaultValue","Value"))
        $refresh  = To-YN (Read-ReaderByNames $rdr $cshsCols @("Refresh","AutoRefresh"))
        $byField  = [string](Read-ReaderByNames $rdr $cshsCols @("ByField"))
        # CSHS.FieldID = optional single explicit trigger (SHS1 is the real list).
        $trigId   = [string](Read-ReaderByNames $rdr $cshsCols @("FieldID","TriggerID","TrigID","TrigerID"))
        $trigCol  = [string](Read-ReaderByNames $rdr $cshsCols @("TriggerCol","TrigCol","TriggerColumn","TrigerCol"))
        $forceRf  = To-YN (Read-ReaderByNames $rdr $cshsCols @("FrceRfrsh","ForceRfsh","ForceRefresh","ForceRefr","DisplaySaved"))

        # Decide FMSAction
        $fmsAction = if ($queryId -and [int]$queryId -gt 0) { "Q" }
                     elseif ($actionT -in @("Q")) { "Q" }
                     else { "F" }

        $qInfo = $null
        if ($fmsAction -eq "Q" -and $queryId) {
            $qi = [int]$queryId
            if ($queries.ContainsKey($qi)) { $qInfo = $queries[$qi] }
        }

        # Build trigger list — SHS1 first (canonical), then CSHS.FieldID,
        # finally fall back to QueryBody parsing if both are empty.
        $allTriggers = New-Object System.Collections.Generic.List[string]
        if ($shs1Map.ContainsKey($indexId)) {
            foreach ($t in $shs1Map[$indexId]) {
                if ($t -and (-not $allTriggers.Contains($t))) { [void]$allTriggers.Add($t) }
            }
        }
        if ((-not [string]::IsNullOrWhiteSpace($trigId)) -and (-not $allTriggers.Contains($trigId))) {
            [void]$allTriggers.Add($trigId)
        }
        if ($allTriggers.Count -eq 0) {
            $qBodyForParse = if ($qInfo) { [string]$qInfo.Body } else { "" }
            foreach ($t in (Get-QueryTriggers $qBodyForParse)) {
                if (-not $allTriggers.Contains($t)) { [void]$allTriggers.Add($t) }
            }
        }

        # Emit 1+ rows (one per trigger, or 1 empty-trigger row)
        $emitArr = if ($allTriggers.Count -eq 0) { @("") } else { @($allTriggers) }
        foreach ($trig in $emitArr) {
            [void]$output.Add([pscustomobject]@{
                Action        = $ExportAction
                FormID        = $formId
                ItemID        = $itemId
                ColumnID      = $colId
                FMSAction     = $fmsAction
                QueryName     = if ($qInfo) { $qInfo.Name }     else { "" }
                QueryCategory = if ($qInfo) { $qInfo.Category } else { "" }
                QueryBody     = if ($qInfo) { $qInfo.Body }     else { "" }
                FixedValue    = if ($fmsAction -eq "F") { $fixedVal } else { "" }
                Refresh       = $refresh
                ByField       = $byField
                TriggerID     = $trig
                TriggerColumn = $trigCol
                ForceRefresh  = $forceRf
            })
        }
    }
    $rdr.Close()
    Write-Log "Read $rowNo CSHS rows; expanded to $($output.Count) CSV rows"

    # ------------------------------------------------------------
    # 4) Write CSV (UTF-8 with BOM)
    # ------------------------------------------------------------
    $dir = Split-Path $OutFile -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    if ($output.Count -eq 0) {
        $csvLines = @('"Action","FormID","ItemID","ColumnID","FMSAction","QueryName","QueryCategory","QueryBody","FixedValue","Refresh","ByField","TriggerID","TriggerColumn","ForceRefresh"')
    } else {
        $csvLines = @($output | ConvertTo-Csv -NoTypeInformation)
    }
    [System.IO.File]::WriteAllLines($OutFile, [string[]]$csvLines, (New-Object System.Text.UTF8Encoding $true))

    Write-Log "=== Done. Exported $($output.Count) records ==="
    Write-Log "Output: $OutFile"
} finally {
    if ($conn.State -eq 'Open') { $conn.Close() }
}
exit 0
