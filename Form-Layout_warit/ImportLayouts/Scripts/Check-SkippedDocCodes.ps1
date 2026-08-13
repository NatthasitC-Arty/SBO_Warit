param(
    [string]$Server="10.10.10.115",
    [string]$DBUser="sa",
    [Parameter(Mandatory=$true)][string]$DBPassword,
    [string]$SourceDB="SBO_SDA_MARK1",
    [string]$TargetDB="SBO_SDA_(Test)_Pre_Training"
)
$ErrorActionPreference="Stop"
$cs="Server=$Server;Database=master;User ID=$DBUser;Password=$DBPassword;Connection Timeout=15;"
$conn = New-Object System.Data.SqlClient.SqlConnection $cs
$conn.Open()

$cmd = $conn.CreateCommand()
$cmd.CommandText = @"
SELECT s.TypeCode, s.DocCode,
       s.DocName AS SystemName,
       t.DocName AS YourName,
       t.Author  AS YourAuthor
FROM [$SourceDB].dbo.RDOC s
INNER JOIN [$TargetDB].dbo.RDOC t ON s.DocCode = t.DocCode
WHERE s.Author = 'System'
  AND t.Author <> 'System'   -- only show conflicts (your custom layout still wins)
ORDER BY s.TypeCode, s.DocCode
"@
$da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
$dt = New-Object System.Data.DataTable
[void]$da.Fill($dt)

Write-Host ""
Write-Host ("=== {0} system DocCodes that conflicted with your custom layouts ===" -f $dt.Rows.Count) -ForegroundColor Yellow
$dt | Format-Table TypeCode, DocCode, SystemName, YourName, YourAuthor -AutoSize -Wrap
$conn.Close()
