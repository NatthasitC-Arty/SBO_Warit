# ============================================================
# Export SAP B1 Form Settings (CPRF table) to CSV.
# CPRF = Column ProFile: column visibility/order/width per
# (FormID, ItemID, ColID) on every form's matrix.
#
# Schema is dynamic (17 cols, version-specific) — exporter
# discovers columns from INFORMATION_SCHEMA then dumps them all.
# Output is round-trip safe: re-import with Action=UPSERT.
# ============================================================
param(
    [string]$Server      = "10.10.10.115",
    [string]$CompanyDB   = "SBO_SDA_Training",
    [string]$DBUser      = "sa",
    [Parameter(Mandatory=$true)][string]$DBPassword,
    [string]$OutFile     = "",
    [string]$LogFile     = "",
    [string]$ExportAction = "UPSERT",
    # Optional filters — blank = export everything.
    [string]$FormID      = "",
    [string]$ItemID      = "",
    [string]$ColID       = "",
    [int]$CommandTimeout = 180
)

$ErrorActionPreference = 'Stop'
trap {
    $msg = "TRAP: $($_.Exception.GetType().FullName): $($_.Exception.Message)`r`n$($_.InvocationInfo.PositionMessage)"
    Add-Content -Path $LogFile -Value "[ERROR] $msg" -Encoding UTF8
    Write-Host $msg -ForegroundColor Red
    break
}

if (-not $OutFile) {
    $OutFile = Join-Path $PSScriptRoot ("..\Config\CPRF_Export_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}
if (-not $LogFile) {
    $LogFile = Join-Path $PSScriptRoot "..\Export_CPRF_Log.txt"
}

function Write-Log {
    param([string]$Msg, [string]$Level = "INFO")
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Msg
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

Write-Log "=== Start CPRF Export ==="
Write-Log "Server   : $Server"
Write-Log "CompanyDB: $CompanyDB"
Write-Log "OutFile  : $OutFile"

$connStr = "Server=$Server;Database=$CompanyDB;User ID=$DBUser;Password=$DBPassword;Connection Timeout=10"
$conn = New-Object System.Data.SqlClient.SqlConnection $connStr
try {
  try {
    $conn.Open()
    Write-Log "Connected to SQL: $Server / $CompanyDB"

    # 1) Discover CPRF columns dynamically (preserve ordinal order)
    $cprfCols = New-Object System.Collections.Generic.List[string]
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='CPRF' ORDER BY ORDINAL_POSITION"
    $rdr = $cmd.ExecuteReader()
    while ($rdr.Read()) { [void]$cprfCols.Add([string]$rdr["COLUMN_NAME"]) }
    $rdr.Close()
    if ($cprfCols.Count -eq 0) {
        Write-Log "CPRF table not found in $CompanyDB" "ERROR"
        exit 2
    }
    Write-Log "CPRF columns ($($cprfCols.Count)): $($cprfCols -join ', ')"

    # 2) Build SELECT with optional WHERE
    $whereParts = @()
    if ($FormID) { $whereParts += "FormID = @fid" }
    if ($ItemID) { $whereParts += "ItemID = @iid" }
    if ($ColID)  { $whereParts += "ColID = @cid" }
    $whereClause = if ($whereParts.Count -gt 0) { " WHERE " + ($whereParts -join " AND ") } else { "" }

    $colList = ($cprfCols | ForEach-Object { "[$_]" }) -join ", "
    $cmd.CommandText = "SELECT $colList FROM CPRF$whereClause ORDER BY FormID, ItemID, ColID"
    $cmd.Parameters.Clear()
    if ($FormID) { [void]$cmd.Parameters.AddWithValue("@fid", $FormID) }
    if ($ItemID) { [void]$cmd.Parameters.AddWithValue("@iid", $ItemID) }
    if ($ColID)  { [void]$cmd.Parameters.AddWithValue("@cid", $ColID) }
    Write-Log "Filter: $(if($whereClause){$whereClause}else{'(none)'})"
    Write-Log "SQL: $($cmd.CommandText)"

    # 3) Read rows -> ordered hashtable per row -> pscustomobject (preserves col order)
    $cmd.CommandTimeout = $CommandTimeout
    Write-Log "Executing reader (timeout=$CommandTimeout s)..."
    $rdr = $cmd.ExecuteReader()
    Write-Log "Reader opened. Iterating rows..."
    $output = New-Object System.Collections.ArrayList
    $rowNo = 0
    while ($rdr.Read()) {
        $rowNo++
        $ord = [ordered]@{ Action = $ExportAction }
        foreach ($c in $cprfCols) {
            $v = $rdr[$c]
            $ord[$c] = if ($v -is [DBNull]) { "" } else { [string]$v }
        }
        [void]$output.Add([pscustomobject]$ord)
        if ($rowNo % 5000 -eq 0) { Write-Log "  ...read $rowNo rows so far" }
    }
    $rdr.Close()
    Write-Log "Read $rowNo CPRF rows"

    # 4) Write CSV (UTF-8 with BOM)
    $dir = Split-Path $OutFile -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    if ($output.Count -eq 0) {
        $header = '"Action",' + (($cprfCols | ForEach-Object { "`"$_`"" }) -join ',')
        $csvLines = @($header)
    } else {
        $csvLines = @($output | ConvertTo-Csv -NoTypeInformation)
    }
    [System.IO.File]::WriteAllLines($OutFile, [string[]]$csvLines, (New-Object System.Text.UTF8Encoding $true))

    Write-Log "=== Done. Exported $($output.Count) record(s) ==="
    Write-Log "Output: $OutFile"
  } catch {
    Write-Log "EXCEPTION: $($_.Exception.GetType().FullName): $($_.Exception.Message)" "ERROR"
    Write-Log "At: $($_.InvocationInfo.PositionMessage)" "ERROR"
    if ($_.Exception.InnerException) {
      Write-Log "Inner: $($_.Exception.InnerException.Message)" "ERROR"
    }
    throw
  }
} finally {
    if ($conn.State -eq 'Open') { $conn.Close() }
}
exit 0
