# ============================================================
# Import User-Defined Values (FMS) to SAP B1 via DIRECT SQL.
# Replaces the DI API path because DI API v10 oFormattedSearches
# doesn't persist QueryId/Refresh/FieldID on Add()/Update().
#
# Reads UDV_Map.csv (new schema with ByField + TriggerID).
# Writes CSHS + OUQR via .NET SqlClient (parameterized).
#
# Multi-row CSV (same FormID|ItemID|ColumnID, multiple TriggerID):
#   - First row's data is written to CSHS (FieldID = first trigger)
#   - Additional TriggerIDs are appended to OUQR.QString as
#     `$[$<ItemID>.<Field>.0]` tokens (if not already present),
#     so B1 UI parses them out as Auto-Refresh trigger fields.
# Actions: ADD | UPDATE | DELETE | UPSERT
# ============================================================
param(
    [string]$Server      = "SLD-C072",
    [string]$CompanyDB   = "SBO_SDA",
    [string]$DBUser      = "sa",
    [string]$DBPassword  = "1q2w3e4r",
    [string]$MapFile     = "$PSScriptRoot\..\Config\UDV_Map.csv",
    [string]$LogFile     = "",
    [switch]$DryRun
)

if (-not $LogFile) {
    $LogFile = Join-Path $PSScriptRoot "..\Import_UDV_SQL_Log.txt"
}

function Write-Log {
    param([string]$Msg, [string]$Level = "INFO")
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Msg
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

function To-YN {
    param([string]$Val)
    if ($Val -and $Val.Trim().ToUpper() -in @("Y","YES","TRUE","1")) { "Y" } else { "N" }
}

# Append a `$[$<ItemID>.<Field>.NUMBER]` reference to QueryBody as a
# no-op arithmetic term (` + 0*CAST(...)`). This injects an ACTIVE SQL
# expression — B1 strips SQL comments before parsing tokens, so the
# old `-- FMS-TRIGGERS:` comment approach did not work.
#
# The injection multiplies by 0 so the numeric result is unchanged.
# Caveat: assumes the original query returns a numeric scalar (true
# for almost all SAP B1 FMS calculations). Pure-string queries
# (e.g. SELECT concat(...) FROM ...) would break — those rarely
# need multi-trigger though.
# Strip legacy trigger injections (run ONCE before adding new triggers,
# never per-iteration — that would erase newly-added CASE injections).
function Remove-LegacyTriggerInjections {
    param([string]$Body)
    # Strip `-- FMS-TRIGGERS:` comment lines
    $Body = $Body -replace '(?ms)(\r?\n)?[ \t]*--[ \t]*FMS-TRIGGERS:[^\r\n]*', ''
    # Strip `+ 0*CAST(...)` no-op injections
    $Body = $Body -replace '\s*\+\s*0\s*\*\s*CAST\(ISNULL\(NULLIF\(''\$\[\$\d+\.[^'']*''[^)]*\)[^)]*\)\s*AS\s*numeric\s*\([^)]+\)\s*\)', ''
    # Strip `* (CASE WHEN '$[$..]'='$[$..]' THEN 1 ELSE 1 END)` injections
    $Body = $Body -replace '\s*\*\s*\(CASE\s+WHEN\s+''\$\[\$\d+\.[^'']*''\s*=\s*''\$\[\$\d+\.[^'']*''\s+THEN\s+1\s+ELSE\s+1\s+END\)', ''
    return $Body.TrimEnd()
}

# Inject a quoted-token trigger reference into QueryBody as a
# multiplicative no-op factor (` * (CASE WHEN '$[$..]'='$[$..]' THEN 1 ELSE 1 END)`).
# B1 strips SQL comments and `+ 0*` no-ops before parsing tokens, but
# `* (CASE...)` stays in the main expression chain so the token is
# detected and shown as Auto-Refresh trigger field in UI.
# Caveat: assumes original query returns numeric scalar.
function Add-TriggerTokenToBody {
    param([string]$Body, [string]$ItemID, [string]$Field)
    if ([string]::IsNullOrWhiteSpace($Field)) { return $Body }
    $Body = $Body.TrimEnd()

    # Skip if body already has QUOTED reference to this field.
    $escField = [regex]::Escape($Field)
    if ($Body -match "'\$\[\$\d+\.$escField") { return $Body }
    if ($Body -match "'\$\[[A-Za-z][A-Za-z0-9_]*\.$escField") { return $Body }

    $body = $Body.TrimEnd(';').TrimEnd()
    $tok = " * (CASE WHEN '`$[`$$ItemID.$Field.NUMBER]'='`$[`$$ItemID.$Field.NUMBER]' THEN 1 ELSE 1 END)"
    return $body + $tok
}

# ------------------------------------------------------------
# Read CSV and group by (FormID|ItemID|ColumnID)
# ------------------------------------------------------------
if (-not (Test-Path $MapFile)) {
    Write-Log "MapFile not found: $MapFile" "ERROR"
    exit 2
}
$rows = @(Import-Csv -Path $MapFile)
Write-Log "=== Start UDV Import (SQL Direct) ==="
Write-Log "MapFile: $MapFile  ($($rows.Count) CSV rows)"
Write-Log "DryRun : $DryRun"

$groups = [ordered]@{}
$rowNo = 0
foreach ($r in $rows) {
    $rowNo++
    $fid = [string]$r.FormID
    $iid = [string]$r.ItemID
    $cid = [string]$r.ColumnID
    if (-not $fid -or -not $iid) {
        Write-Log "  SKIP CSV row $rowNo : missing FormID/ItemID" "WARN"
        continue
    }
    $key = "$fid|$iid|$cid"
    if (-not $groups.Contains($key)) {
        $groups[$key] = [pscustomobject]@{
            Primary  = $r
            Triggers = New-Object System.Collections.Generic.List[string]
        }
    }
    if ($r.TriggerID -and -not $groups[$key].Triggers.Contains($r.TriggerID)) {
        [void]$groups[$key].Triggers.Add($r.TriggerID)
    }
}
Write-Log "Grouped to $($groups.Count) unique CSHS key(s)"

# ------------------------------------------------------------
# Open SQL connection
# ------------------------------------------------------------
$connStr = "Server=$Server;Database=$CompanyDB;User ID=$DBUser;Password=$DBPassword;Connection Timeout=10"
$conn = New-Object System.Data.SqlClient.SqlConnection $connStr
try {
    $conn.Open()
    Write-Log "Connected to SQL: $Server / $CompanyDB"

    # Pre-fetch existing CSHS keys
    $cshsMap = @{}
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "SELECT IndexID, FormID, ItemID, ColID FROM CSHS"
    $rdr = $cmd.ExecuteReader()
    while ($rdr.Read()) {
        $k = "{0}|{1}|{2}" -f [string]$rdr["FormID"], [string]$rdr["ItemID"], [string]$rdr["ColID"]
        $cshsMap[$k] = [int]$rdr["IndexID"]
    }
    $rdr.Close()
    Write-Log "Pre-loaded $($cshsMap.Count) existing CSHS rows"

    # Pre-fetch OUQR queries by name
    $ouqrMap = @{}
    $cmd.CommandText = "SELECT IntrnalKey, QName FROM OUQR"
    $rdr = $cmd.ExecuteReader()
    while ($rdr.Read()) {
        $n = [string]$rdr["QName"]
        if ($n -and (-not $ouqrMap.ContainsKey($n))) {
            $ouqrMap[$n] = [int]$rdr["IntrnalKey"]
        }
    }
    $rdr.Close()
    Write-Log "Pre-loaded $($ouqrMap.Count) OUQR queries"

    if ($DryRun) {
        Write-Log "DryRun=ON. Preview only:"
        foreach ($entry in $groups.GetEnumerator()) {
            $g = $entry.Value
            $r = $g.Primary
            $tag = if ($cshsMap.ContainsKey($entry.Key)) { "EXISTS" } else { "NEW" }
            $tList = if ($g.Triggers.Count -gt 0) { ($g.Triggers -join ', ') } else { "(none)" }
            Write-Log "  [$tag] $($entry.Key) Action=$($r.Action) FMSAction=$($r.FMSAction) Query=$($r.QueryName) Triggers=$tList"
        }
        exit 0
    }

    # Current MAX IndexID for INSERTs
    $cmd.CommandText = "SELECT ISNULL(MAX(IndexID),0) FROM CSHS"
    $nextIdx = [int]$cmd.ExecuteScalar()
    Write-Log "Current MAX(IndexID)=$nextIdx"

    # OUQR.IntrnalKey is NOT identity in SAP B1 — manage it manually.
    $cmd.CommandText = "SELECT ISNULL(MAX(IntrnalKey),0) FROM OUQR"
    $nextQKey = [int]$cmd.ExecuteScalar()
    Write-Log "Current MAX(IntrnalKey)=$nextQKey"

    # ------------------------------------------------------------
    # Transaction
    # ------------------------------------------------------------
    $tx = $conn.BeginTransaction()
    $ok = 0; $fail = 0; $skip = 0
    try {
        foreach ($entry in $groups.GetEnumerator()) {
            $group = $entry.Value
            $r = $group.Primary
            $key = $entry.Key
            $action = ([string]$r.Action).ToUpper()
            $label = "$key Action=$action"

            try {
                if ($action -notin @("ADD","UPDATE","DELETE","UPSERT")) {
                    Write-Log "  SKIP $label : unknown action" "WARN"; $skip++; continue
                }
                $exists = $cshsMap.ContainsKey($key)

                # ---------- Handle OUQR (find/create/update for Q-type) ----------
                $queryId = $null
                if ($r.FMSAction -eq "Q" -and $action -ne "DELETE") {
                    $qname = [string]$r.QueryName
                    if (-not $qname) {
                        throw "QueryName required for FMSAction=Q"
                    }
                    # Triggers are stored in SHS1 child table, NOT injected
                    # into QueryBody. Just clean legacy injection cruft from
                    # body and save it verbatim.
                    $body = [string]$r.QueryBody
                    $body = Remove-LegacyTriggerInjections -Body $body
                    if ($ouqrMap.ContainsKey($qname)) {
                        $queryId = $ouqrMap[$qname]
                        # Update existing query body if changed
                        $uq = $conn.CreateCommand()
                        $uq.Transaction = $tx
                        $uq.CommandText = "UPDATE OUQR SET QString = @s WHERE IntrnalKey = @k"
                        [void]$uq.Parameters.AddWithValue("@s", $body)
                        [void]$uq.Parameters.AddWithValue("@k", $queryId)
                        [void]$uq.ExecuteNonQuery()
                        Write-Log "  Reuse Query '$qname' (IntrnalKey=$queryId), body updated"
                    } else {
                        # Insert new OUQR (supply IntrnalKey manually)
                        $cat = if ($r.QueryCategory) { [int]$r.QueryCategory } else { -1 }
                        $nextQKey++
                        $queryId = $nextQKey
                        $ins = $conn.CreateCommand()
                        $ins.Transaction = $tx
                        $ins.CommandText = "INSERT INTO OUQR (IntrnalKey, QName, QString, QCategory) VALUES (@k, @n, @s, @c)"
                        [void]$ins.Parameters.AddWithValue("@k", $queryId)
                        [void]$ins.Parameters.AddWithValue("@n", $qname)
                        [void]$ins.Parameters.AddWithValue("@s", $body)
                        [void]$ins.Parameters.AddWithValue("@c", $cat)
                        [void]$ins.ExecuteNonQuery()
                        $ouqrMap[$qname] = $queryId
                        Write-Log "  Created Query '$qname' IntrnalKey=$queryId"
                    }
                }

                # ---------- Handle CSHS ----------
                $actionT  = if ($r.FMSAction -eq "Q") { "2" } else { "0" }
                $refresh  = To-YN $r.Refresh
                $forceRf  = To-YN $r.ForceRefresh
                $byField  = if ($r.ByField) { ([string]$r.ByField).Trim() } else { "N" }
                # Triggers go in SHS1 (child table), not CSHS.FieldID.
                # Leave FieldID empty so UI uses the full SHS1 list.
                $fieldID  = ""

                switch ($action) {
                    "ADD" {
                        if ($exists) {
                            Write-Log "  SKIP $label : exists (use UPDATE/UPSERT)" "WARN"; $skip++; continue
                        }
                        $nextIdx++
                        $ins = $conn.CreateCommand()
                        $ins.Transaction = $tx
                        $ins.CommandText = @"
INSERT INTO CSHS (IndexID, FormID, ItemID, ColID, ActionT, QueryId, Refresh, ByField, FieldID, FrceRfrsh)
VALUES (@i, @f, @t, @c, @a, @q, @r, @b, @fl, @fr)
"@
                        [void]$ins.Parameters.AddWithValue("@i", $nextIdx)
                        [void]$ins.Parameters.AddWithValue("@f", [string]$r.FormID)
                        [void]$ins.Parameters.AddWithValue("@t", [string]$r.ItemID)
                        [void]$ins.Parameters.AddWithValue("@c", [string]$r.ColumnID)
                        [void]$ins.Parameters.AddWithValue("@a", $actionT)
                        if ($null -eq $queryId) {
                            [void]$ins.Parameters.AddWithValue("@q", [DBNull]::Value)
                        } else {
                            [void]$ins.Parameters.AddWithValue("@q", $queryId)
                        }
                        [void]$ins.Parameters.AddWithValue("@r", $refresh)
                        [void]$ins.Parameters.AddWithValue("@b", $byField)
                        [void]$ins.Parameters.AddWithValue("@fl", $fieldID)
                        [void]$ins.Parameters.AddWithValue("@fr", $forceRf)
                        [void]$ins.ExecuteNonQuery()
                        $cshsMap[$key] = $nextIdx
                        Write-Log "  OK ADD $label IndexID=$nextIdx"
                        $ok++
                    }
                    "UPDATE" {
                        if (-not $exists) {
                            Write-Log "  SKIP $label : not found (use ADD/UPSERT)" "WARN"; $skip++; continue
                        }
                        $idx = $cshsMap[$key]
                        $upd = $conn.CreateCommand()
                        $upd.Transaction = $tx
                        $upd.CommandText = "UPDATE CSHS SET ActionT=@a, QueryId=@q, Refresh=@r, ByField=@b, FieldID=@fl, FrceRfrsh=@fr WHERE IndexID=@i"
                        [void]$upd.Parameters.AddWithValue("@i", $idx)
                        [void]$upd.Parameters.AddWithValue("@a", $actionT)
                        if ($null -eq $queryId) {
                            [void]$upd.Parameters.AddWithValue("@q", [DBNull]::Value)
                        } else {
                            [void]$upd.Parameters.AddWithValue("@q", $queryId)
                        }
                        [void]$upd.Parameters.AddWithValue("@r", $refresh)
                        [void]$upd.Parameters.AddWithValue("@b", $byField)
                        [void]$upd.Parameters.AddWithValue("@fl", $fieldID)
                        [void]$upd.Parameters.AddWithValue("@fr", $forceRf)
                        [void]$upd.ExecuteNonQuery()
                        Write-Log "  OK UPDATE $label IndexID=$idx"
                        $ok++
                    }
                    "UPSERT" {
                        if ($exists) {
                            $idx = $cshsMap[$key]
                            $upd = $conn.CreateCommand()
                            $upd.Transaction = $tx
                            $upd.CommandText = "UPDATE CSHS SET ActionT=@a, QueryId=@q, Refresh=@r, ByField=@b, FieldID=@fl, FrceRfrsh=@fr WHERE IndexID=@i"
                            [void]$upd.Parameters.AddWithValue("@i", $idx)
                            [void]$upd.Parameters.AddWithValue("@a", $actionT)
                            if ($null -eq $queryId) {
                                [void]$upd.Parameters.AddWithValue("@q", [DBNull]::Value)
                            } else {
                                [void]$upd.Parameters.AddWithValue("@q", $queryId)
                            }
                            [void]$upd.Parameters.AddWithValue("@r", $refresh)
                            [void]$upd.Parameters.AddWithValue("@b", $byField)
                            [void]$upd.Parameters.AddWithValue("@fl", $fieldID)
                            [void]$upd.Parameters.AddWithValue("@fr", $forceRf)
                            [void]$upd.ExecuteNonQuery()
                            Write-Log "  OK UPSERT->UPDATE $label IndexID=$idx"
                        } else {
                            $nextIdx++
                            $ins = $conn.CreateCommand()
                            $ins.Transaction = $tx
                            $ins.CommandText = @"
INSERT INTO CSHS (IndexID, FormID, ItemID, ColID, ActionT, QueryId, Refresh, ByField, FieldID, FrceRfrsh)
VALUES (@i, @f, @t, @c, @a, @q, @r, @b, @fl, @fr)
"@
                            [void]$ins.Parameters.AddWithValue("@i", $nextIdx)
                            [void]$ins.Parameters.AddWithValue("@f", [string]$r.FormID)
                            [void]$ins.Parameters.AddWithValue("@t", [string]$r.ItemID)
                            [void]$ins.Parameters.AddWithValue("@c", [string]$r.ColumnID)
                            [void]$ins.Parameters.AddWithValue("@a", $actionT)
                            if ($null -eq $queryId) {
                                [void]$ins.Parameters.AddWithValue("@q", [DBNull]::Value)
                            } else {
                                [void]$ins.Parameters.AddWithValue("@q", $queryId)
                            }
                            [void]$ins.Parameters.AddWithValue("@r", $refresh)
                            [void]$ins.Parameters.AddWithValue("@b", $byField)
                            [void]$ins.Parameters.AddWithValue("@fl", $fieldID)
                            [void]$ins.Parameters.AddWithValue("@fr", $forceRf)
                            [void]$ins.ExecuteNonQuery()
                            $cshsMap[$key] = $nextIdx
                            Write-Log "  OK UPSERT->ADD $label IndexID=$nextIdx"
                        }
                        $ok++
                    }
                    "DELETE" {
                        if (-not $exists) {
                            Write-Log "  SKIP $label : not found" "WARN"; $skip++; continue
                        }
                        $idx = $cshsMap[$key]
                        # Remove SHS1 child rows first (FK / orphan-safety)
                        $delS = $conn.CreateCommand()
                        $delS.Transaction = $tx
                        $delS.CommandText = "DELETE FROM SHS1 WHERE IndexID=@i"
                        [void]$delS.Parameters.AddWithValue("@i", $idx)
                        $nSHS1 = $delS.ExecuteNonQuery()
                        # Now CSHS
                        $del = $conn.CreateCommand()
                        $del.Transaction = $tx
                        $del.CommandText = "DELETE FROM CSHS WHERE IndexID=@i"
                        [void]$del.Parameters.AddWithValue("@i", $idx)
                        [void]$del.ExecuteNonQuery()
                        $cshsMap.Remove($key)
                        Write-Log "  OK DELETE $label IndexID=$idx (SHS1 child rows removed: $nSHS1)"
                        $ok++
                    }
                }

                # ---------- Sync triggers to SHS1 (child table) ----------
                # SHS1 (IndexID int, FieldID nvarchar) stores the trigger
                # field list shown in B1 UI "Auto Refresh -> Fields".
                # Strategy: clear all SHS1 rows for this CSHS IndexID,
                # then INSERT one row per trigger from the CSV group.
                if ($action -ne "DELETE" -and $cshsMap.ContainsKey($key)) {
                    $idxForTrig = $cshsMap[$key]
                    $delT = $conn.CreateCommand()
                    $delT.Transaction = $tx
                    $delT.CommandText = "DELETE FROM SHS1 WHERE IndexID = @i"
                    [void]$delT.Parameters.AddWithValue("@i", $idxForTrig)
                    $deletedTrig = $delT.ExecuteNonQuery()
                    $addedTrig = 0
                    foreach ($t in $group.Triggers) {
                        if (-not $t) { continue }
                        $insT = $conn.CreateCommand()
                        $insT.Transaction = $tx
                        $insT.CommandText = "INSERT INTO SHS1 (IndexID, FieldID) VALUES (@i, @f)"
                        [void]$insT.Parameters.AddWithValue("@i", $idxForTrig)
                        [void]$insT.Parameters.AddWithValue("@f", [string]$t)
                        [void]$insT.ExecuteNonQuery()
                        $addedTrig++
                    }
                    if ($group.Triggers.Count -gt 0) {
                        $all = $group.Triggers -join ', '
                        Write-Log "  Triggers (SHS1) IndexID=$idxForTrig : -$deletedTrig +$addedTrig ($all)"
                    } elseif ($deletedTrig -gt 0) {
                        Write-Log "  Triggers (SHS1) IndexID=$idxForTrig : cleared $deletedTrig old row(s) (no triggers in CSV)"
                    }
                } elseif ($action -eq "DELETE" -and $cshsMap.ContainsKey($key) -eq $false) {
                    # CSHS row was just deleted - also clean orphan SHS1 rows
                    # (handled by the earlier DELETE since IndexID is gone)
                }
            } catch {
                Write-Log "  FAIL $label : $($_.Exception.Message)" "ERROR"
                $fail++
            }
        }
        $tx.Commit()
        Write-Log "Committed. OK=$ok FAIL=$fail SKIP=$skip"
    } catch {
        $tx.Rollback()
        Write-Log "ROLLBACK: $($_.Exception.Message)" "ERROR"
        exit 3
    }
} finally {
    if ($conn.State -eq 'Open') { $conn.Close() }
}
exit 0
