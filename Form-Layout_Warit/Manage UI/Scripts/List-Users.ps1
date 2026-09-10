# ============================================================
# List SAP B1 users from OUSR (UserSign, UserCode, U_NAME, etc.)
# Used to find UserSign of the account that edited Form Settings
# manually, so we can source CPRF rows from that user.
# Also lists per-user CPRF row counts for a given FormID (helps
# spot which user actually has customizations).
# ============================================================
param(
    [string]$Server      = "10.10.10.115",
    [string]$CompanyDB   = "SBO_SDA_Training",
    [string]$DBUser      = "sa",
    [Parameter(Mandatory=$true)][string]$DBPassword,
    [string]$FormID      = ""  # optional: show CPRF row counts per user
)

$ErrorActionPreference = 'Stop'

$connStr = "Server=$Server;Database=$CompanyDB;User ID=$DBUser;Password=$DBPassword;Connection Timeout=10"
$conn = New-Object System.Data.SqlClient.SqlConnection $connStr
try {
    $conn.Open()
    Write-Host ""
    Write-Host "=== OUSR (users) ===" -ForegroundColor Cyan

    $sql = @"
SELECT u.USERID AS UserSign,
       u.USER_CODE AS UserCode,
       u.U_NAME AS UserName,
       u.SUPERUSER AS IsAdmin,
       u.LOCKED AS Locked
FROM OUSR u
ORDER BY u.USERID
"@
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $sql
    $da = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
    $dt = New-Object System.Data.DataTable
    [void]$da.Fill($dt)
    $dt | Format-Table -AutoSize | Out-Host

    if ($FormID) {
        Write-Host ""
        Write-Host "=== CPRF row count per UserSign for FormID=$FormID ===" -ForegroundColor Cyan
        $cmd2 = $conn.CreateCommand()
        $cmd2.CommandText = @"
SELECT c.UserSign,
       u.USER_CODE AS UserCode,
       COUNT(*) AS CprfRows,
       COUNT(DISTINCT c.TPLId) AS Templates
FROM CPRF c
LEFT JOIN OUSR u ON u.USERID = c.UserSign
WHERE c.FormID = @fid
GROUP BY c.UserSign, u.USER_CODE
ORDER BY CprfRows DESC
"@
        [void]$cmd2.Parameters.AddWithValue("@fid", $FormID)
        $da2 = New-Object System.Data.SqlClient.SqlDataAdapter $cmd2
        $dt2 = New-Object System.Data.DataTable
        [void]$da2.Fill($dt2)
        $dt2 | Format-Table -AutoSize | Out-Host
    }
} finally {
    if ($conn.State -eq 'Open') { $conn.Close() }
}
