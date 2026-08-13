# ============================================================
# Rollback by Selection: list non-system layouts in RDOC, let user pick which to delete.
# Optional keyword filter narrows the list before picking.
# Deletes RITM/RDC1/RCON children + DFLT_PRNTING orphans, all in one transaction.
# ============================================================
param(
    [Parameter(Mandatory=$true)][string]$Server,
    [Parameter(Mandatory=$true)][string]$CompanyDB,
    [string]$DBUser     = "sa",
    [Parameter(Mandatory=$true)][string]$DBPassword,
    [string]$SystemAuthor = "System",
    [string]$Filter       = "",
    [switch]$DryRun,
    [switch]$Force
)

Write-Host "=== Rollback by selection ===" -ForegroundColor Cyan
Write-Host "Server      : $Server"
Write-Host "CompanyDB   : $CompanyDB"
Write-Host "Skip Author : '$SystemAuthor' (system layouts kept)"
if ($Filter) { Write-Host "Filter      : '$Filter'" }
Write-Host ""

$cs = "Server=$Server;Database=$CompanyDB;User ID=$DBUser;Password=$DBPassword;Connection Timeout=10;"
$conn = New-Object System.Data.SqlClient.SqlConnection $cs
try {
    $conn.Open()
} catch {
    Write-Host "ERROR connecting: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Interactive keyword filter if not passed in
if (-not $Filter) {
    $Filter = Read-Host "Filter by DocName/Author keyword (empty=show all)"
}

$sql = "SELECT DocCode, DocName, TypeCode, ISNULL(Author,'') AS Author FROM RDOC WHERE (Author<>@a OR Author IS NULL)"
if ($Filter) { $sql += " AND (DocName LIKE @f OR Author LIKE @f OR TypeCode LIKE @f)" }
$sql += " ORDER BY TypeCode, DocName"

$cmd = $conn.CreateCommand()
$cmd.CommandText = $sql
[void]$cmd.Parameters.AddWithValue("@a", $SystemAuthor)
if ($Filter) { [void]$cmd.Parameters.AddWithValue("@f", "%$Filter%") }

$rdr = $cmd.ExecuteReader()
$layouts = New-Object System.Collections.ArrayList
while ($rdr.Read()) {
    [void]$layouts.Add([PSCustomObject]@{
        Idx      = $layouts.Count + 1
        DocCode  = [string]$rdr["DocCode"]
        DocName  = [string]$rdr["DocName"]
        TypeCode = [string]$rdr["TypeCode"]
        Author   = [string]$rdr["Author"]
    })
}
$rdr.Close()

if ($layouts.Count -eq 0) {
    Write-Host "No matching layouts found." -ForegroundColor Green
    $conn.Close()
    return
}

Write-Host ""
Write-Host ("=== {0} layout(s) found ===" -f $layouts.Count) -ForegroundColor Yellow
$layouts | Format-Table @{Label='#';Expression={$_.Idx};Width=5}, DocCode, TypeCode, Author, DocName -AutoSize

Write-Host "Selection examples: 1   |   1,3,5   |   1-10   |   1-5,8,12-15   |   all   |   (empty=cancel)" -ForegroundColor DarkGray
$pick = Read-Host "Enter selection"
if ([string]::IsNullOrWhiteSpace($pick)) {
    Write-Host "Cancelled." -ForegroundColor Yellow
    $conn.Close()
    return
}

# Parse selection
$selected = New-Object System.Collections.ArrayList
if ($pick.Trim().ToLower() -eq "all") {
    foreach ($l in $layouts) { [void]$selected.Add($l) }
} else {
    $byIdx = @{}
    foreach ($l in $layouts) { $byIdx[[int]$l.Idx] = $l }
    foreach ($part in ($pick -split ',')) {
        $p = $part.Trim()
        if ($p -match '^(\d+)\s*-\s*(\d+)$') {
            $a = [int]$matches[1]; $b = [int]$matches[2]
            if ($a -gt $b) { $tmp=$a; $a=$b; $b=$tmp }
            for ($i=$a; $i -le $b; $i++) {
                if ($byIdx.ContainsKey($i) -and -not $selected.Contains($byIdx[$i])) {
                    [void]$selected.Add($byIdx[$i])
                }
            }
        } elseif ($p -match '^\d+$') {
            $i = [int]$p
            if ($byIdx.ContainsKey($i) -and -not $selected.Contains($byIdx[$i])) {
                [void]$selected.Add($byIdx[$i])
            }
        } elseif ($p) {
            Write-Host ("  [skip invalid token] '{0}'" -f $p) -ForegroundColor DarkGray
        }
    }
}

if ($selected.Count -eq 0) {
    Write-Host "Nothing selected. Cancelled." -ForegroundColor Yellow
    $conn.Close()
    return
}

Write-Host ""
Write-Host ("=== {0} layout(s) selected to DELETE ===" -f $selected.Count) -ForegroundColor Red
$selected | Format-Table @{Label='#';Expression={$_.Idx};Width=5}, DocCode, TypeCode, Author, DocName -AutoSize

if ($DryRun) {
    Write-Host "DryRun mode - no changes made." -ForegroundColor Cyan
    $conn.Close()
    return
}

if (-not $Force) {
    $ans = Read-Host ("Delete these {0} layout(s)? Type 'yes' to confirm" -f $selected.Count)
    if ($ans -ne "yes") {
        Write-Host "Cancelled." -ForegroundColor Yellow
        $conn.Close()
        return
    }
}

# Build #DelDocs by inlining DocCodes (avoids sp_executesql temp-table scope issue)
$docCodes = $selected | ForEach-Object { "'" + ($_.DocCode -replace "'","''") + "'" }
$inList = $docCodes -join ","

$tran = $conn.BeginTransaction()
try {
    foreach ($tbl in @('RITM','RDC1','RCON')) {
        try {
            $c = $conn.CreateCommand()
            $c.Transaction = $tran
            $c.CommandText = "DELETE FROM dbo.$tbl WHERE DocCode IN ($inList)"
            $c.CommandTimeout = 300
            $cn = $c.ExecuteNonQuery()
            Write-Host ("  {0,-4}: deleted {1,5} child rows" -f $tbl, $cn) -ForegroundColor DarkYellow
        } catch {
            Write-Host ("  {0,-4}: skipped ({1})" -f $tbl, $_.Exception.Message) -ForegroundColor DarkGray
        }
    }

    $exec = $conn.CreateCommand()
    $exec.Transaction = $tran
    $exec.CommandText = "DELETE FROM RDOC WHERE DocCode IN ($inList)"
    $exec.CommandTimeout = 300
    $n = $exec.ExecuteNonQuery()
    Write-Host ("  RDOC: deleted {0,5} rows" -f $n) -ForegroundColor Green

    try {
        $d = $conn.CreateCommand()
        $d.Transaction = $tran
        $d.CommandText = "DELETE FROM dbo.DFLT_PRNTING WHERE DocCode IN ($inList)"
        $d.CommandTimeout = 300
        $dn = $d.ExecuteNonQuery()
        Write-Host ("  DFLT_PRNTING entries cleaned: {0,5} rows" -f $dn) -ForegroundColor DarkYellow
    } catch {
        Write-Host ("  DFLT_PRNTING: skipped ({0})" -f $_.Exception.Message) -ForegroundColor DarkGray
    }

    $tran.Commit()
    Write-Host ""
    Write-Host ("=== Summary: Deleted={0} / Selected={1} ===" -f $n, $selected.Count) -ForegroundColor Cyan
} catch {
    $tran.Rollback()
    Write-Host ""
    Write-Host ("[ROLLBACK] {0}" -f $_.Exception.Message) -ForegroundColor Red
    throw
} finally {
    $conn.Close()
}
