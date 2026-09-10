# ============================================================
# Probe UI Template definitions: find table holding TPLId -> Name
# mapping. Also dump manager's TPLId breakdown for a given FormID.
# ============================================================
param(
    [string]$Server      = "10.10.10.115",
    [string]$CompanyDB   = "SBO_SDA_Training",
    [string]$DBUser      = "sa",
    [Parameter(Mandatory=$true)][string]$DBPassword,
    [string]$FormID      = "139",
    [int]$UserSign       = 1
)

$ErrorActionPreference = 'Stop'
$conn = New-Object System.Data.SqlClient.SqlConnection "Server=$Server;Database=$CompanyDB;User ID=$DBUser;Password=$DBPassword;Connection Timeout=10"
$conn.Open()

function Run([string]$title, [string]$sql) {
    Write-Host ""; Write-Host "=== $title ===" -ForegroundColor Cyan
    $c = $conn.CreateCommand(); $c.CommandText = $sql
    $a = New-Object System.Data.SqlClient.SqlDataAdapter $c
    $t = New-Object System.Data.DataTable; [void]$a.Fill($t)
    if ($t.Rows.Count -eq 0) { Write-Host "(empty)" -ForegroundColor DarkGray }
    else { $t | Format-Table -AutoSize | Out-Host }
}

# 1) Tables likely holding UI Template definitions
Run "Tables with TPLId column" @"
SELECT t.name AS TableName,
       STUFF((SELECT ', ' + c.name FROM sys.columns c WHERE c.object_id=t.object_id ORDER BY c.column_id FOR XML PATH('')),1,2,'') AS Columns
FROM sys.tables t
WHERE EXISTS (SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name IN ('TPLId','TPLID','TplId','TemplateId','TemplateID'))
ORDER BY t.name
"@

# 2) Manager's TPLId rows for FormID
Run "TPLId breakdown for FormID=$FormID UserSign=$UserSign" @"
SELECT TPLId,
       COUNT(*) AS Rows_,
       COUNT(DISTINCT ItemID) AS Items,
       COUNT(DISTINCT ColID) AS Cols,
       MIN(CASE WHEN VisInForm='Y' THEN 'Y' ELSE NULL END) AS HasVisibleY
FROM CPRF
WHERE FormID = '$FormID' AND UserSign = $UserSign
GROUP BY TPLId
ORDER BY TPLId
"@

# 3) Distinct TPLId across ALL users for the form
Run "TPLId universe for FormID=$FormID (across all users)" @"
SELECT TPLId, COUNT(*) AS TotalRows, COUNT(DISTINCT UserSign) AS UserCount
FROM CPRF WHERE FormID = '$FormID'
GROUP BY TPLId ORDER BY TPLId
"@

$conn.Close()
Write-Host ""; Write-Host "Done." -ForegroundColor Green
