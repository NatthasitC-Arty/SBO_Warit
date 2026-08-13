# ============================================================
# Mirror FMS from one FormID to another via direct SQL on CSHS.
# Bypasses DI API entirely because DI API v10 silently fails to
# persist QueryId/Refresh/FieldID on oFormattedSearches.Add()/Update().
#
# Logic:
#   1. DELETE existing rows on TargetFormID (limited by ItemID if given)
#   2. INSERT SELECT from SourceFormID, only changing FormID column
#   3. All other columns (ActionT, QueryId, Refresh, ByField, FieldID,
#      FrceRfrsh) copied verbatim from source.
# ============================================================
param(
    [string]$Server      = "SLD-C072",
    [string]$CompanyDB   = "SBO_SDA",
    [string]$DBUser      = "sa",
    [string]$DBPassword  = "1q2w3e4r",
    [Parameter(Mandatory=$true)]
    [string]$SourceFormID,
    [Parameter(Mandatory=$true)]
    [string]$TargetFormID,
    [string]$ItemID      = "",                                  # blank = mirror all items on the form
    [string]$LogFile     = "",
    [switch]$DryRun
)

# Defer log path resolution to script body (param defaults can lose
# $PSScriptRoot in some invocation contexts -> bad C:\..\ resolution)
if (-not $LogFile) {
    $LogFile = Join-Path $PSScriptRoot "..\Mirror_FMS_Log.txt"
}

function Write-Log {
    param([string]$Msg, [string]$Level = "INFO")
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Msg
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

Write-Log "=== Mirror FMS: source=$SourceFormID -> target=$TargetFormID (Item=$ItemID) DryRun=$DryRun ==="

$connStr = "Server=$Server;Database=$CompanyDB;User ID=$DBUser;Password=$DBPassword;Connection Timeout=10"
$conn = New-Object System.Data.SqlClient.SqlConnection $connStr
try {
    $conn.Open()
    Write-Log "Connected to SQL: $Server / $CompanyDB"

    # 1a) Discover CSHS-related child tables.
    # B1 v10 uses SHS1 (not CSHS1) for multi-trigger field storage,
    # linked via IndexID. Also include any CSHS* / SHS* siblings.
    $childTables = @()
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "SELECT name FROM sys.tables WHERE (name LIKE 'CSHS%' OR name LIKE 'SHS%') AND name <> 'CSHS' ORDER BY name"
    $rdr = $cmd.ExecuteReader()
    while ($rdr.Read()) { $childTables += [string]$rdr["name"] }
    $rdr.Close()
    Write-Log "Child tables found: $(if ($childTables.Count -eq 0) { '(none)' } else { $childTables -join ', ' })"

    # 1b) Discover columns of each child table (for dynamic INSERT)
    $childCols = @{}
    foreach ($t in $childTables) {
        $cmd.CommandText = "SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '$t' ORDER BY ORDINAL_POSITION"
        $rdr = $cmd.ExecuteReader()
        $cols = @()
        while ($rdr.Read()) { $cols += [string]$rdr["COLUMN_NAME"] }
        $rdr.Close()
        $childCols[$t] = $cols
        Write-Log "  $t : $($cols -join ', ')"
    }

    # 2) Read source CSHS rows (capture IndexID so we can fetch child rows)
    $itemFilter = ""
    if ($ItemID) { $itemFilter = " AND ItemID = @item" }
    $sel = $conn.CreateCommand()
    $sel.CommandText = "SELECT IndexID, FormID, ItemID, ColID, ActionT, QueryId, Refresh, ByField, FieldID, FrceRfrsh FROM CSHS WHERE FormID = @src$itemFilter ORDER BY ItemID, ColID"
    [void]$sel.Parameters.AddWithValue("@src", $SourceFormID)
    if ($ItemID) { [void]$sel.Parameters.AddWithValue("@item", $ItemID) }
    $rdr = $sel.ExecuteReader()
    $srcRows = New-Object System.Collections.ArrayList
    while ($rdr.Read()) {
        [void]$srcRows.Add([pscustomobject]@{
            SrcIndexID = [int]$rdr["IndexID"]
            ItemID   = [string]$rdr["ItemID"]
            ColID    = [string]$rdr["ColID"]
            ActionT  = [string]$rdr["ActionT"]
            QueryId  = if ($rdr["QueryId"] -is [DBNull]) { $null } else { [int]$rdr["QueryId"] }
            Refresh  = [string]$rdr["Refresh"]
            ByField  = [string]$rdr["ByField"]
            FieldID  = [string]$rdr["FieldID"]
            FrceRfrsh= [string]$rdr["FrceRfrsh"]
        })
    }
    $rdr.Close()
    Write-Log "Source $SourceFormID has $($srcRows.Count) FMS row(s) to mirror"
    if ($srcRows.Count -eq 0) {
        Write-Log "Nothing to mirror — exiting." "WARN"
        exit 0
    }

    foreach ($r in $srcRows) {
        $q = if ($null -eq $r.QueryId) { "(null)" } else { $r.QueryId }
        Write-Log ("  Source  Idx={0,-5} Item={1,-12} Col={2,-22} ActionT={3}  QueryId={4,-5}  Refresh={5}  ByField={6}  FieldID={7,-18}  FrceRfrsh={8}" -f $r.SrcIndexID, $r.ItemID, $r.ColID, $r.ActionT, $q, $r.Refresh, $r.ByField, $r.FieldID, $r.FrceRfrsh)
        # Show child rows preview
        foreach ($t in $childTables) {
            $prev = $conn.CreateCommand()
            $prev.CommandText = "SELECT COUNT(*) FROM $t WHERE IndexID = @idx"
            [void]$prev.Parameters.AddWithValue("@idx", $r.SrcIndexID)
            $cnt = [int]$prev.ExecuteScalar()
            if ($cnt -gt 0) { Write-Log "             $t : $cnt row(s)" }
        }
    }

    if ($DryRun) {
        Write-Log "DryRun=ON. No DELETE / INSERT executed."
        exit 0
    }

    # 3) Execute DELETE + INSERT inside a transaction
    $tx = $conn.BeginTransaction()
    try {
        # 3a) Find current target IndexIDs (for child-table delete)
        $tgtIdxList = @()
        $qIdx = $conn.CreateCommand()
        $qIdx.Transaction = $tx
        $qIdx.CommandText = "SELECT IndexID FROM CSHS WHERE FormID = @tgt$itemFilter"
        [void]$qIdx.Parameters.AddWithValue("@tgt", $TargetFormID)
        if ($ItemID) { [void]$qIdx.Parameters.AddWithValue("@item", $ItemID) }
        $rdr = $qIdx.ExecuteReader()
        while ($rdr.Read()) { $tgtIdxList += [int]$rdr["IndexID"] }
        $rdr.Close()
        Write-Log "Current target IndexIDs: $(if ($tgtIdxList.Count -eq 0) { '(none)' } else { $tgtIdxList -join ',' })"

        # 3b) DELETE from child tables first (FK / dependency order)
        if ($tgtIdxList.Count -gt 0) {
            $idxIn = $tgtIdxList -join ","
            foreach ($t in $childTables) {
                $dc = $conn.CreateCommand()
                $dc.Transaction = $tx
                $dc.CommandText = "DELETE FROM $t WHERE IndexID IN ($idxIn)"
                $n = $dc.ExecuteNonQuery()
                Write-Log "Deleted $n row(s) from $t (target IndexIDs)"
            }
        }

        # 3c) DELETE existing rows on target CSHS
        $del = $conn.CreateCommand()
        $del.Transaction = $tx
        $del.CommandText = "DELETE FROM CSHS WHERE FormID = @tgt$itemFilter"
        [void]$del.Parameters.AddWithValue("@tgt", $TargetFormID)
        if ($ItemID) { [void]$del.Parameters.AddWithValue("@item", $ItemID) }
        $delCount = $del.ExecuteNonQuery()
        Write-Log "Deleted $delCount existing row(s) on target CSHS $TargetFormID"

        # 3d) IndexID is NOT NULL but NOT IDENTITY -> compute next manually
        $maxCmd = $conn.CreateCommand()
        $maxCmd.Transaction = $tx
        $maxCmd.CommandText = "SELECT ISNULL(MAX(IndexID), 0) FROM CSHS"
        $nextIdx = [int]$maxCmd.ExecuteScalar()
        Write-Log "Current MAX(IndexID) = $nextIdx -> new rows start at $($nextIdx + 1)"

        # 3e) INSERT each source row with FormID changed to target,
        #     then cascade child rows (CSHS1 etc) using new IndexID
        $insCount = 0
        $childCounts = @{}
        foreach ($t in $childTables) { $childCounts[$t] = 0 }
        foreach ($r in $srcRows) {
            $nextIdx++
            $ins = $conn.CreateCommand()
            $ins.Transaction = $tx
            $ins.CommandText = @"
INSERT INTO CSHS (IndexID, FormID, ItemID, ColID, ActionT, QueryId, Refresh, ByField, FieldID, FrceRfrsh)
VALUES (@idx, @fid, @iid, @cid, @act, @qid, @ref, @byf, @fld, @frc)
"@
            [void]$ins.Parameters.AddWithValue("@idx", $nextIdx)
            [void]$ins.Parameters.AddWithValue("@fid", $TargetFormID)
            [void]$ins.Parameters.AddWithValue("@iid", $r.ItemID)
            [void]$ins.Parameters.AddWithValue("@cid", $r.ColID)
            [void]$ins.Parameters.AddWithValue("@act", $r.ActionT)
            if ($null -eq $r.QueryId) {
                [void]$ins.Parameters.AddWithValue("@qid", [DBNull]::Value)
            } else {
                [void]$ins.Parameters.AddWithValue("@qid", $r.QueryId)
            }
            [void]$ins.Parameters.AddWithValue("@ref", $r.Refresh)
            [void]$ins.Parameters.AddWithValue("@byf", $r.ByField)
            [void]$ins.Parameters.AddWithValue("@fld", $r.FieldID)
            [void]$ins.Parameters.AddWithValue("@frc", $r.FrceRfrsh)
            $insCount += $ins.ExecuteNonQuery()

            # Cascade child rows (dynamic column list, IndexID -> new value)
            foreach ($t in $childTables) {
                $cols = $childCols[$t]
                $insertCols = ($cols | ForEach-Object { "[$_]" }) -join ","
                $selectCols = ($cols | ForEach-Object { if ($_ -eq "IndexID") { "@newIdx" } else { "[$_]" } }) -join ","
                $cins = $conn.CreateCommand()
                $cins.Transaction = $tx
                $cins.CommandText = "INSERT INTO $t ($insertCols) SELECT $selectCols FROM $t WHERE IndexID = @oldIdx"
                [void]$cins.Parameters.AddWithValue("@newIdx", $nextIdx)
                [void]$cins.Parameters.AddWithValue("@oldIdx", $r.SrcIndexID)
                $n = $cins.ExecuteNonQuery()
                $childCounts[$t] += $n
                if ($n -gt 0) {
                    Write-Log "  -> $t : copied $n row(s) (IndexID $($r.SrcIndexID) -> $nextIdx)"
                }
            }
        }
        $tx.Commit()
        $childSum = ($childCounts.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ", "
        Write-Log "Inserted $insCount CSHS row(s); child rows: $(if ($childSum) { $childSum } else { '(none)' }) — committed"
    } catch {
        $tx.Rollback()
        Write-Log "ROLLBACK: $($_.Exception.Message)" "ERROR"
        exit 3
    }

    # 3) Verify
    $ver = $conn.CreateCommand()
    $ver.CommandText = "SELECT ColID, ActionT, QueryId, Refresh, ByField, FieldID, FrceRfrsh FROM CSHS WHERE FormID = @tgt$itemFilter ORDER BY ColID"
    [void]$ver.Parameters.AddWithValue("@tgt", $TargetFormID)
    if ($ItemID) { [void]$ver.Parameters.AddWithValue("@item", $ItemID) }
    $rdr = $ver.ExecuteReader()
    Write-Log "Post-import state of $TargetFormID :"
    while ($rdr.Read()) {
        $q = if ($rdr["QueryId"] -is [DBNull]) { "(null)" } else { $rdr["QueryId"] }
        Write-Log ("  Result  Col={0,-22} ActionT={1}  QueryId={2,-5}  Refresh={3}  ByField={4}  FieldID={5,-18}  FrceRfrsh={6}" -f $rdr["ColID"], $rdr["ActionT"], $q, $rdr["Refresh"], $rdr["ByField"], $rdr["FieldID"], $rdr["FrceRfrsh"])
    }
    $rdr.Close()

    Write-Log "=== Done ==="
} finally {
    if ($conn.State -eq 'Open') { $conn.Close() }
}
exit 0
