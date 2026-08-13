# ============================================================
# Backup RDOC table to disk (BCP-style export)
# Creates: RDOC_Backup_<timestamp>.csv + RDOC_Backup_<timestamp>.bak (binary)
# Use Restore-RDOC-Backup.ps1 to roll back
# ============================================================
param(
    [string]$Server     = "SLD-C072",
    [string]$CompanyDB  = "SBO_SDA",
    [string]$DBUser     = "sa",
    [string]$DBPassword = "1q2w3e4r",
    [string]$OutDir     = "C:\SDA\SDA\Form-Layout\ImportLayouts\Backups"
)

if (-not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$outFile = Join-Path $OutDir "RDOC_Backup_$ts.bak"

$cs = "Server=$Server;Database=$CompanyDB;User ID=$DBUser;Password=$DBPassword;Connection Timeout=10;"
$conn = New-Object System.Data.SqlClient.SqlConnection $cs
$conn.Open()
$cmd = $conn.CreateCommand()
$cmd.CommandText = "SELECT COUNT(*) FROM RDOC"
$total = $cmd.ExecuteScalar()
Write-Host "RDOC currently has $total rows" -ForegroundColor Cyan

# Save list of existing DocCodes (for rollback identification of NEW rows after insert)
$cmd.CommandText = "SELECT DocCode, TypeCode, DocName, Category, DATALENGTH(Template) AS Bytes, RptHash FROM RDOC"
$da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
$dt = New-Object System.Data.DataTable
[void]$da.Fill($dt)

$csvFile = $outFile -replace "\.bak$", ".csv"
$dt | Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8
Write-Host "Saved index of $($dt.Rows.Count) existing rows -> $csvFile" -ForegroundColor Green

# Save full binary dump (each row -> binary blob in single file with header)
$fs = [System.IO.File]::OpenWrite($outFile)
$bw = New-Object System.IO.BinaryWriter $fs
$bw.Write([byte[]]([byte]0x52,0x44,0x4F,0x43,0x42,0x4B,0x50,0x31)) # magic "RDOCBKP1"
$bw.Write([int32]$dt.Rows.Count)

$cmd2 = $conn.CreateCommand()
$cmd2.CommandText = "SELECT DocCode, Template FROM RDOC WHERE Template IS NOT NULL"
$reader = $cmd2.ExecuteReader()
$savedBinary = 0
while ($reader.Read()) {
    $code = [string]$reader["DocCode"]
    $blob = $reader["Template"]
    if ($blob -is [byte[]]) {
        $codeBytes = [System.Text.Encoding]::ASCII.GetBytes($code.PadRight(8))
        $bw.Write($codeBytes)
        $bw.Write([int32]$blob.Length)
        $bw.Write($blob)
        $savedBinary++
    }
}
$reader.Close()
$bw.Close()
$fs.Close()
$conn.Close()

Write-Host "Saved $savedBinary binary templates -> $outFile" -ForegroundColor Green
Write-Host ""
Write-Host "Backup complete:" -ForegroundColor Yellow
Write-Host "  Index: $csvFile"
Write-Host "  Binary: $outFile ($((Get-Item $outFile).Length) bytes)"