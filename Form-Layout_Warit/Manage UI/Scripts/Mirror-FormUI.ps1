# ============================================================
# Mirror CPRF rows from (SourceUserSign, SourceTPLId) to
# (TargetUserSign, TargetTPLId) for a given FormID.
#
# Use case: manager edits Form Settings under a local template
# (e.g. UserSign=1, TPLId=34 = 1_local_template_3 of UI ALL),
# then push that profile to the master template (UserSign=26,
# TPLId=3 = "UI ALL") so new users get it.
#
# Logic in transaction:
#   1. Read source rows (SELECT COUNT-only in DryRun)
#   2. DELETE target rows for (FormID, TargetUserSign, TargetTPLId)
#   3. INSERT each source row with UserSign + TPLId rewritten
#   4. Commit | Rollback on error
# ============================================================
param(
    [string]$Server      = "10.10.10.115",
    [string]$CompanyDB   = "SBO_Update_UI",
    [string]$DBUser      = "sa",
    [Parameter(Mandatory=$true)][string]$DBPassword,
    [Parameter(Mandatory=$true)][string]$FormID,
    [Parameter(Mandatory=$true)][int]$SourceUserSign,
    [Parameter(Mandatory=$true)][int]$SourceTPLId,
    [Parameter(Mandatory=$true)][int]$TargetUserSign,
    [Parameter(Mandatory=$true)][int]$TargetTPLId,
    [string]$LogFile     = "",
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
if (-not $LogFile) { $LogFile = Join-Path $PSScriptRoot "..\Mirror_FormUI_Log.txt" }

function Write-Log {
    param([string]$Msg, [string]$Level = "INFO")
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Msg
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

trap {
    Write-Log "TRAP: $($_.Exception.GetType().FullName): $($_.Exception.Message)" "ERROR"
    Write-Log "At: $($_.InvocationInfo.PositionMessage)" "ERROR"
    break
}

Write-Log "=== Mirror Form UI ==="
Write-Log "FormID         : $FormID"
Write-Log "Source         : UserSign=$SourceUserSign TPLId=$SourceTPLId"
Write-Log "Target         : UserSign=$TargetUserSign TPLId=$TargetTPLId"
Write-Log "DryRun         : $DryRun"

if ($SourceUserSign -eq $TargetUserSign -and $SourceTPLId -eq $TargetTPLId) {
    Write-Log "Source == Target. Nothing to do." "WARN"
    exit 1
}

$connStr = "Server=$Server;Database=$CompanyDB;User ID=$DBUser;Password=$DBPassword;Connection Timeout=15"
$conn = New-Object System.Data.SqlClient.SqlConnection $connStr
try {
    $conn.Open()
    Write-Log "Connected: $Server / $CompanyDB"

    # 1) Discover CPRF columns (preserve ordinal order)
    $cprfCols = New-Object System.Collections.Generic.List[string]
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME='CPRF' ORDER BY ORDINAL_POSITION"
    $rdr = $cmd.ExecuteReader()
    while ($rdr.Read()) { [void]$cprfCols.Add([string]$rdr["COLUMN_NAME"]) }
    $rdr.Close()
    Write-Log "CPRF cols ($($cprfCols.Count)): $($cprfCols -join ', ')"

    # 2) Count source / target
    $cmd.CommandText = "SELECT COUNT(*) FROM CPRF WHERE FormID=@f AND UserSign=@u AND TPLId=@t"
    $cmd.Parameters.Clear()
    [void]$cmd.Parameters.AddWithValue("@f", $FormID)
    [void]$cmd.Parameters.AddWithValue("@u", $SourceUserSign)
    [void]$cmd.Parameters.AddWithValue("@t", $SourceTPLId)
    $srcCount = [int]$cmd.ExecuteScalar()
    Write-Log "Source rows: $srcCount"

    $cmd.Parameters["@u"].Value = $TargetUserSign
    $cmd.Parameters["@t"].Value = $TargetTPLId
    $tgtCount = [int]$cmd.ExecuteScalar()
    Write-Log "Target rows (existing, will be deleted): $tgtCount"

    if ($srcCount -eq 0) {
        Write-Log "No source rows. Aborting." "ERROR"
        exit 2
    }

    if ($DryRun) {
        Write-Log "DryRun=ON. Would DELETE $tgtCount target rows and INSERT $srcCount new rows."
        Write-Log "=== DryRun done ==="
        exit 0
    }

    # 3) Real transaction
    $tx = $conn.BeginTransaction()
    try {
        # 3a) DELETE target
        $del = $conn.CreateCommand()
        $del.Transaction = $tx
        $del.CommandText = "DELETE FROM CPRF WHERE FormID=@f AND UserSign=@u AND TPLId=@t"
        [void]$del.Parameters.AddWithValue("@f", $FormID)
        [void]$del.Parameters.AddWithValue("@u", $TargetUserSign)
        [void]$del.Parameters.AddWithValue("@t", $TargetTPLId)
        $nDel = $del.ExecuteNonQuery()
        Write-Log "DELETE target rows: $nDel"

        # 3b) Read source rows
        $srcSel = $conn.CreateCommand()
        $srcSel.Transaction = $tx
        $colList = ($cprfCols | ForEach-Object { "[$_]" }) -join ", "
        $srcSel.CommandText = "SELECT $colList FROM CPRF WHERE FormID=@f AND UserSign=@u AND TPLId=@t"
        [void]$srcSel.Parameters.AddWithValue("@f", $FormID)
        [void]$srcSel.Parameters.AddWithValue("@u", $SourceUserSign)
        [void]$srcSel.Parameters.AddWithValue("@t", $SourceTPLId)
        $srcRdr = $srcSel.ExecuteReader()
        $srcRows = New-Object System.Collections.ArrayList
        while ($srcRdr.Read()) {
            $row = @{}
            foreach ($c in $cprfCols) { $row[$c] = $srcRdr[$c] }
            [void]$srcRows.Add($row)
        }
        $srcRdr.Close()
        Write-Log "Buffered $($srcRows.Count) source rows"

        # 3c) INSERT remapped
        $colNames = ($cprfCols | ForEach-Object { "[$_]" }) -join ", "
        $paramNames = ($cprfCols | ForEach-Object { "@p_$_" }) -join ", "
        $ins = $conn.CreateCommand()
        $ins.Transaction = $tx
        $ins.CommandText = "INSERT INTO CPRF ($colNames) VALUES ($paramNames)"
        foreach ($c in $cprfCols) {
            [void]$ins.Parameters.Add("@p_$c", [System.Data.SqlDbType]::NVarChar)
        }

        $nIns = 0
        foreach ($row in $srcRows) {
            foreach ($c in $cprfCols) {
                $val = $row[$c]
                # Rewrite UserSign/TPLId; copy others verbatim
                if     ($c -eq "UserSign") { $val = $TargetUserSign }
                elseif ($c -eq "TPLId")    { $val = $TargetTPLId }
                if ($null -eq $val -or $val -is [DBNull]) {
                    $ins.Parameters["@p_$c"].Value = [DBNull]::Value
                } else {
                    $ins.Parameters["@p_$c"].Value = $val
                }
            }
            [void]$ins.ExecuteNonQuery()
            $nIns++
            if ($nIns % 200 -eq 0) { Write-Log "  ...inserted $nIns rows" }
        }
        Write-Log "INSERT done: $nIns rows"

        $tx.Commit()
        Write-Log "=== Committed. DELETE=$nDel INSERT=$nIns ==="
    } catch {
        $tx.Rollback()
        Write-Log "ROLLBACK: $($_.Exception.Message)" "ERROR"
        throw
    }
} finally {
    if ($conn.State -eq 'Open') { $conn.Close() }
}
exit 0
