param(
    [string]$Server="10.10.10.115",[string]$CompanyDB="SBO_Update_UI",
    [string]$DBUser="sa",[Parameter(Mandatory=$true)][string]$DBPassword
)
$ErrorActionPreference='Stop'
$conn=New-Object System.Data.SqlClient.SqlConnection "Server=$Server;Database=$CompanyDB;User ID=$DBUser;Password=$DBPassword;Connection Timeout=10"
$conn.Open()
function Run([string]$t,[string]$s){ Write-Host "`n=== $t ===" -ForegroundColor Cyan; $c=$conn.CreateCommand(); $c.CommandText=$s; $a=New-Object System.Data.SqlClient.SqlDataAdapter $c; $d=New-Object System.Data.DataTable; [void]$a.Fill($d); if($d.Rows.Count -eq 0){Write-Host "(empty)" -ForegroundColor DarkGray}else{$d|Format-Table -AutoSize|Out-Host} }

Run "All UICU rows (UI Templates master)" @"
SELECT u.TPLId, u.TPLName, u.TPLDesc, u.UserID, o.USER_CODE AS Owner, u.Parent, u.IsTemplate
FROM UICU u
LEFT JOIN OUSR o ON o.USERID = u.UserID
ORDER BY u.TPLId
"@

Run "UIC3: Template-User mapping (which user gets which template)" @"
SELECT u3.TPLId, t.TPLName, u3.UserID, o.USER_CODE, u3.IsTemplate
FROM UIC3 u3
LEFT JOIN UICU t ON t.TPLId = u3.TPLId
LEFT JOIN OUSR o ON o.USERID = u3.UserID
ORDER BY u3.TPLId, u3.UserID
"@

Run "Manager TPLIds for FormID=139 (joined with UICU)" @"
SELECT DISTINCT c.TPLId, t.TPLName, t.TPLDesc, t.UserID AS TemplateOwner, oo.USER_CODE AS OwnerCode, t.IsTemplate
FROM CPRF c
LEFT JOIN UICU t ON t.TPLId = c.TPLId
LEFT JOIN OUSR oo ON oo.USERID = t.UserID
WHERE c.FormID = '139' AND c.UserSign = 1
ORDER BY c.TPLId
"@

$conn.Close()
Write-Host "`nDone." -ForegroundColor Green
