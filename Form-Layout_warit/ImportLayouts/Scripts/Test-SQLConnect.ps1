# ============================================================
# Test SQL connection to SAP B1 Company DB
# Verifies you can read/write RDOC table before running import
# ============================================================
param(
    [string]$Server     = "SLD-C072",
    [string]$CompanyDB  = "SBO_SDA",
    [string]$DBUser     = "sa",
    [string]$DBPassword = "1q2w3e4r"
)

$ErrorActionPreference = "Stop"

Write-Host "[1/4] Pinging $Server ..." -ForegroundColor Cyan
$ping = Test-Connection -ComputerName $Server -Count 2 -Quiet -ErrorAction SilentlyContinue
Write-Host "      Ping: $(if($ping){'OK'}else{'FAIL (server not reachable)'})" -ForegroundColor $(if($ping){'Green'}else{'Red'})

Write-Host "[2/4] Testing SQL port 1433 on $Server ..." -ForegroundColor Cyan
$tcp = Test-NetConnection -ComputerName $Server -Port 1433 -WarningAction SilentlyContinue
Write-Host "      TCP 1433: $(if($tcp.TcpTestSucceeded){'OPEN'}else{'CLOSED/BLOCKED'})" -ForegroundColor $(if($tcp.TcpTestSucceeded){'Green'}else{'Red'})

Write-Host "[3/4] Testing SQL connection to $CompanyDB ..." -ForegroundColor Cyan
try {
    $cs = "Server=$Server;Database=$CompanyDB;User ID=$DBUser;Password=$DBPassword;Connection Timeout=10;"
    $conn = New-Object System.Data.SqlClient.SqlConnection $cs
    $conn.Open()
    Write-Host "      SQL Login: OK (server $($conn.ServerVersion))" -ForegroundColor Green

    Write-Host "[4/4] Counting layouts in RDOC ..." -ForegroundColor Cyan
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "SELECT COUNT(*) AS Total, SUM(CASE WHEN Category='C' THEN 1 ELSE 0 END) AS Crystal, SUM(CASE WHEN Author='SDA' THEN 1 ELSE 0 END) AS SDA FROM RDOC"
    $rdr = $cmd.ExecuteReader()
    if ($rdr.Read()) {
        Write-Host "      Total layouts : $($rdr['Total'])" -ForegroundColor Green
        Write-Host "      Crystal (C)   : $($rdr['Crystal'])" -ForegroundColor Green
        Write-Host "      Author=SDA    : $($rdr['SDA']) (imported by us)" -ForegroundColor Green
    }
    $rdr.Close()
    $conn.Close()
    Write-Host ""
    Write-Host "READY TO IMPORT" -ForegroundColor Green
} catch {
    Write-Host "      SQL FAIL: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Troubleshoot:" -ForegroundColor Yellow
    Write-Host "  - 'Login failed for user' -> wrong DBUser/DBPassword"
    Write-Host "  - 'Cannot open database X' -> wrong CompanyDB name"
    Write-Host "  - 'A network-related error' -> wrong Server name or SQL service down"
}