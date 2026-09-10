param(
    [string]$Server="10.10.10.115",[string]$CompanyDB="SBO_Update_UI",
    [string]$DBUser="sa",[Parameter(Mandatory=$true)][string]$DBPassword,
    [string]$FormID="139"
)
$ErrorActionPreference='Stop'
$conn=New-Object System.Data.SqlClient.SqlConnection "Server=$Server;Database=$CompanyDB;User ID=$DBUser;Password=$DBPassword;Connection Timeout=10"
$conn.Open()
function Run([string]$t,[string]$s){ Write-Host "`n=== $t ===" -ForegroundColor Cyan; $c=$conn.CreateCommand(); $c.CommandText=$s; $a=New-Object System.Data.SqlClient.SqlDataAdapter $c; $d=New-Object System.Data.DataTable; [void]$a.Fill($d); if($d.Rows.Count -eq 0){Write-Host "(empty)" -ForegroundColor DarkGray}else{$d|Format-Table -AutoSize|Out-Host} }

Run "Each TPLId in CPRF for FormID=$FormID -> which UserSign owns it?" @"
SELECT c.TPLId,
       t.TPLName,
       t.UserID AS UICU_Owner,
       ou.USER_CODE AS UICU_OwnerCode,
       c.UserSign AS CPRF_UserSign,
       oc.USER_CODE AS CPRF_UserCode,
       COUNT(*) AS Rows_
FROM CPRF c
LEFT JOIN UICU t ON t.TPLId = c.TPLId
LEFT JOIN OUSR ou ON ou.USERID = t.UserID
LEFT JOIN OUSR oc ON oc.USERID = c.UserSign
WHERE c.FormID = '$FormID'
GROUP BY c.TPLId, t.TPLName, t.UserID, ou.USER_CODE, c.UserSign, oc.USER_CODE
ORDER BY c.TPLId, c.UserSign
"@

$conn.Close()
Write-Host "`nDone." -ForegroundColor Green
