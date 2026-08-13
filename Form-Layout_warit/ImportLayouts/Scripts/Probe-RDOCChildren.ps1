param(
    [string]$Server = "10.10.10.115",
    [string]$DBUser = "sa",
    [Parameter(Mandatory=$true)][string]$DBPassword,
    [string]$SourceDB = "SBO_SDA_MARK1",
    [string]$TargetDB = "SBO_SDA_(Test)_Pre_Training"
)

$ErrorActionPreference = "Stop"
$cs = "Server=$Server;Database=master;User ID=$DBUser;Password=$DBPassword;Connection Timeout=15;"
$conn = New-Object System.Data.SqlClient.SqlConnection $cs
$conn.Open()

foreach ($db in @($SourceDB, $TargetDB)) {
    Write-Host ""
    Write-Host "=== Tables in $db that have a DocCode column ===" -ForegroundColor Cyan
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = @"
SELECT t.name AS TableName,
       (SELECT COUNT(*) FROM [$db].sys.columns c WHERE c.object_id=t.object_id) AS Cols
FROM [$db].sys.tables t
INNER JOIN [$db].sys.columns c2 ON c2.object_id=t.object_id AND c2.name='DocCode'
WHERE t.name LIKE 'R%' OR t.name LIKE 'OR%'
ORDER BY t.name
"@
    $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
    $dt = New-Object System.Data.DataTable
    [void]$da.Fill($dt)
    $dt | Format-Table -AutoSize
}

$conn.Close()
