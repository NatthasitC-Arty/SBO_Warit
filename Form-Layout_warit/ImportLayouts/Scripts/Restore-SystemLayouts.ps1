# ============================================================
# Restore SAP B1 system layouts from another company DB
# Source : <SourceDB>.dbo.RDOC where Author='System' (+ RITM/RFLT/RPRM children)
# Target : <TargetDB>.dbo.RDOC (+ children)
# Skips DocCodes that already exist in target.
# ============================================================
param(
    [string]$Server       = "10.10.10.115",
    [string]$TargetDB     = "SBO_SDA_(Test)_Pre_Training",
    [string]$SourceDB     = "SBO_SDA_MARK1",
    [string]$DBUser       = "sa",
    [Parameter(Mandatory=$true)][string]$DBPassword,
    [string]$SystemAuthor = "System",
    [switch]$DryRun,
    [switch]$Force,
    [switch]$SkipBackup
)

$ErrorActionPreference = "Stop"

function New-Conn {
    param($db)
    $cs = "Server=$Server;Database=$db;User ID=$DBUser;Password=$DBPassword;Connection Timeout=15;"
    $c = New-Object System.Data.SqlClient.SqlConnection $cs
    $c.Open()
    return $c
}

function Exec-Scalar {
    param($conn, $sql, $params = @{})
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $sql
    $cmd.CommandTimeout = 600
    foreach ($k in $params.Keys) { [void]$cmd.Parameters.AddWithValue($k, $params[$k]) }
    return $cmd.ExecuteScalar()
}

function Exec-Query {
    param($conn, $sql, $params = @{})
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $sql
    $cmd.CommandTimeout = 600
    foreach ($k in $params.Keys) { [void]$cmd.Parameters.AddWithValue($k, $params[$k]) }
    $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
    $dt = New-Object System.Data.DataTable
    [void]$da.Fill($dt)
    return $dt
}

function Exec-NonQuery {
    param($conn, $tran, $sql)
    $cmd = $conn.CreateCommand()
    $cmd.Transaction = $tran
    $cmd.CommandText = $sql
    $cmd.CommandTimeout = 600
    return $cmd.ExecuteNonQuery()
}

Write-Host ""
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host " Restore SAP B1 System Layouts" -ForegroundColor Cyan
Write-Host "=========================================================" -ForegroundColor Cyan
Write-Host (" Server      : {0}" -f $Server)
Write-Host (" Source DB   : {0}" -f $SourceDB)
Write-Host (" Target DB   : {0}" -f $TargetDB)
Write-Host (" SystemAuthor: '{0}'" -f $SystemAuthor)
Write-Host (" Mode        : {0}" -f $(if ($DryRun) { "DRY RUN (no changes)" } else { "EXECUTE" }))
Write-Host ""

# ---------- 1. Connect ----------
$conn = New-Conn -db $TargetDB
Write-Host "[OK] Connected to target DB" -ForegroundColor Green

# ---------- 2. Counts ----------
$srcCnt = Exec-Scalar $conn "SELECT COUNT(*) FROM [$SourceDB].dbo.RDOC WHERE Author=@a" @{ "@a" = $SystemAuthor }
$tgtCnt = Exec-Scalar $conn "SELECT COUNT(*) FROM [$TargetDB].dbo.RDOC WHERE Author=@a" @{ "@a" = $SystemAuthor }
Write-Host ("[INFO] Source has {0} '{1}' rows" -f $srcCnt, $SystemAuthor) -ForegroundColor Gray
Write-Host ("[INFO] Target has {0} '{1}' rows currently" -f $tgtCnt, $SystemAuthor) -ForegroundColor Gray

if ($srcCnt -eq 0) {
    Write-Host "[ERROR] Source DB has no system layouts -- abort" -ForegroundColor Red
    $conn.Close(); return
}

# ---------- 3. Schema match ----------
Write-Host ""
Write-Host "[CHECK] Validating schema match..." -ForegroundColor Yellow
$schemaSql = @"
SELECT t.TABLE_NAME, t.COLUMN_NAME, t.DATA_TYPE,
       ISNULL(t.CHARACTER_MAXIMUM_LENGTH,-1) AS LenT,
       s.DATA_TYPE AS SrcType,
       ISNULL(s.CHARACTER_MAXIMUM_LENGTH,-1) AS LenS
FROM [$TargetDB].INFORMATION_SCHEMA.COLUMNS t
FULL OUTER JOIN [$SourceDB].INFORMATION_SCHEMA.COLUMNS s
    ON t.TABLE_NAME=s.TABLE_NAME AND t.COLUMN_NAME=s.COLUMN_NAME
WHERE COALESCE(t.TABLE_NAME,s.TABLE_NAME) IN ('RDOC','RITM','RDC1','RCON')
  AND (t.COLUMN_NAME IS NULL OR s.COLUMN_NAME IS NULL
       OR t.DATA_TYPE <> s.DATA_TYPE
       OR ISNULL(t.CHARACTER_MAXIMUM_LENGTH,-1) <> ISNULL(s.CHARACTER_MAXIMUM_LENGTH,-1))
"@
$diff = Exec-Query $conn $schemaSql
if ($diff.Rows.Count -gt 0) {
    Write-Host "[WARN] Schema differences detected:" -ForegroundColor Yellow
    $diff | Format-Table -AutoSize
    if (-not $Force) {
        Write-Host "Use -Force to override (risky!)" -ForegroundColor Red
        $conn.Close(); return
    }
} else {
    Write-Host "[OK] Schemas match exactly" -ForegroundColor Green
}

# ---------- 4. Preview ----------
# Children dedup: only insert children for DocCodes whose RDOC row will be inserted
# (i.e. system DocCode that doesn't exist in target yet).
$cteFilter = @"
WITH NewDocs AS (
    SELECT s.DocCode
    FROM [$SourceDB].dbo.RDOC s
    WHERE s.Author=N'$SystemAuthor'
      AND NOT EXISTS (SELECT 1 FROM [$TargetDB].dbo.RDOC t WHERE t.DocCode=s.DocCode)
)
"@
Write-Host ""
Write-Host "[PREVIEW] Rows that will be inserted:" -ForegroundColor Yellow
$prevSql = @"
$cteFilter
SELECT 'RDOC' AS Tbl, COUNT(*) AS WillInsert FROM NewDocs
UNION ALL SELECT 'RITM', COUNT(*) FROM [$SourceDB].dbo.RITM s WHERE s.DocCode IN (SELECT DocCode FROM NewDocs)
UNION ALL SELECT 'RDC1', COUNT(*) FROM [$SourceDB].dbo.RDC1 s WHERE s.DocCode IN (SELECT DocCode FROM NewDocs)
UNION ALL SELECT 'RCON', COUNT(*) FROM [$SourceDB].dbo.RCON s WHERE s.DocCode IN (SELECT DocCode FROM NewDocs)
"@
$preview = Exec-Query $conn $prevSql
$preview | Format-Table -AutoSize

if ($DryRun) {
    Write-Host "[DRYRUN] No changes made -- end" -ForegroundColor Cyan
    $conn.Close(); return
}

# ---------- 5. Backup ----------
if (-not $SkipBackup) {
    Write-Host ""
    Write-Host "[BACKUP] Creating safety backup on SQL Server host..." -ForegroundColor Yellow
    $masterConn = New-Conn -db 'master'

    # Query SQL Server default backup directory (local path on the SQL host)
    $pathCmd = $masterConn.CreateCommand()
    $pathCmd.CommandText = @"
DECLARE @p NVARCHAR(500);
EXEC master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE',
    N'Software\Microsoft\MSSQLServer\MSSQLServer', N'BackupDirectory', @p OUTPUT;
SELECT ISNULL(@p, CAST(SERVERPROPERTY('InstanceDefaultDataPath') AS NVARCHAR(500)));
"@
    $bakDir = [string]$pathCmd.ExecuteScalar()
    if ([string]::IsNullOrWhiteSpace($bakDir)) {
        Write-Host "[WARN] Cannot determine SQL Server backup directory -- skipping backup" -ForegroundColor Yellow
        Write-Host "[INFO] Transaction rollback will protect against partial inserts" -ForegroundColor Gray
    } else {
        $ts = Get-Date -Format "yyyyMMdd_HHmmss"
        $bakFile = "$bakDir\Pre_Restore_${TargetDB}_$ts.bak" -replace '\(','_' -replace '\)','_'
        $bakCmd = $masterConn.CreateCommand()
        $bakCmd.CommandTimeout = 1800
        $bakCmd.CommandText = "BACKUP DATABASE [$TargetDB] TO DISK=N'$bakFile' WITH FORMAT, COMPRESSION, INIT, NAME=N'Pre-Restore safety backup'"
        try {
            [void]$bakCmd.ExecuteNonQuery()
            Write-Host "[OK] Backup written to SQL host: $bakFile" -ForegroundColor Green
        } catch {
            Write-Host "[WARN] Backup failed: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "[INFO] Transaction rollback will still protect against partial inserts" -ForegroundColor Gray
            Write-Host "[INFO] You also have Backups\RDOC_Backup_20260421_170239.* as fallback" -ForegroundColor Gray
        }
    }
    $masterConn.Close()
}

# ---------- 6. Confirm ----------
if (-not $Force) {
    Write-Host ""
    Write-Host "Ready to restore. Type 'yes' to proceed: " -ForegroundColor Red -NoNewline
    $ans = Read-Host
    if ($ans -ne 'yes') {
        Write-Host "Cancelled." -ForegroundColor Yellow
        $conn.Close(); return
    }
}

# ---------- 7. Insert in transaction ----------
Write-Host ""
Write-Host "[RESTORE] Beginning transaction..." -ForegroundColor Cyan
$tran = $conn.BeginTransaction()

try {
    # Stage NewDocs into a temp table local to this connection so all inserts share the same set
    $stage = @"
SELECT s.DocCode INTO #NewDocs
FROM [$SourceDB].dbo.RDOC s
WHERE s.Author=N'$SystemAuthor'
  AND NOT EXISTS (SELECT 1 FROM [$TargetDB].dbo.RDOC t WHERE t.DocCode=s.DocCode);
CREATE INDEX IX_NewDocs ON #NewDocs(DocCode);
"@
    [void](Exec-NonQuery $conn $tran $stage)

    # Clean orphan children FIRST (rows whose DocCode no longer has parent in target.RDOC)
    # These are leftover from the bad delete operation that hit RDOC but not children.
    foreach ($tbl in @('RITM','RDC1','RCON')) {
        $sql = "DELETE FROM [$TargetDB].dbo.$tbl WHERE DocCode NOT IN (SELECT DocCode FROM [$TargetDB].dbo.RDOC)"
        $n = Exec-NonQuery $conn $tran $sql
        Write-Host ("  Cleaned {0} orphans: {1,5} rows" -f $tbl, $n) -ForegroundColor DarkYellow
    }

    $sqlRDOC = @"
INSERT INTO [$TargetDB].dbo.RDOC
SELECT s.* FROM [$SourceDB].dbo.RDOC s
WHERE s.DocCode IN (SELECT DocCode FROM #NewDocs);
"@
    $n = Exec-NonQuery $conn $tran $sqlRDOC
    Write-Host ("  RDOC: {0,5} rows" -f $n) -ForegroundColor Green

    foreach ($tbl in @('RITM','RDC1','RCON')) {
        $sql = @"
INSERT INTO [$TargetDB].dbo.$tbl
SELECT s.* FROM [$SourceDB].dbo.$tbl s
WHERE s.DocCode IN (SELECT DocCode FROM #NewDocs);
"@
        $n = Exec-NonQuery $conn $tran $sql
        Write-Host ("  {0}: {1,5} rows" -f $tbl, $n) -ForegroundColor Green
    }

    # Clean DFLT_PRNTING orphans (use TRY so missing table doesn't abort)
    try {
        $del = Exec-NonQuery $conn $tran "DELETE FROM [$TargetDB].dbo.DFLT_PRNTING WHERE DocCode IS NOT NULL AND DocCode<>N'' AND DocCode NOT IN (SELECT DocCode FROM [$TargetDB].dbo.RDOC)"
        Write-Host ("  DFLT_PRNTING orphans cleaned: {0,5} rows" -f $del) -ForegroundColor Green
    } catch {
        Write-Host ("  DFLT_PRNTING: skipped ({0})" -f $_.Exception.Message) -ForegroundColor DarkGray
    }

    [void](Exec-NonQuery $conn $tran "DROP TABLE #NewDocs")

    $tran.Commit()
    Write-Host ""
    Write-Host "[COMMIT] Transaction committed" -ForegroundColor Green
} catch {
    $tran.Rollback()
    Write-Host ""
    Write-Host ("[ROLLBACK] Error: {0}" -f $_.Exception.Message) -ForegroundColor Red
    throw
} finally {
    $conn.Close()
}

# ---------- 8. Verify ----------
Write-Host ""
Write-Host "[VERIFY] Final RDOC counts by Author:" -ForegroundColor Yellow
$conn2 = New-Conn -db $TargetDB
$verify = Exec-Query $conn2 "SELECT ISNULL(Author,'<NULL>') AS Author, COUNT(*) AS Cnt FROM RDOC GROUP BY Author ORDER BY Cnt DESC"
$verify | Format-Table -AutoSize
$conn2.Close()

Write-Host "Done. Logout/Login SAP B1 client and try Print Preview." -ForegroundColor Cyan
