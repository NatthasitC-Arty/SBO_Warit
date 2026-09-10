# Diff CPRF: rows present in one template but missing in another,
# or with different visibility/order/width.
param(
    [string]$Server="10.10.10.115",[string]$CompanyDB="SBO_Update_UI",
    [string]$DBUser="sa",[Parameter(Mandatory=$true)][string]$DBPassword,
    [Parameter(Mandatory=$true)][string]$FormID,
    [Parameter(Mandatory=$true)][int]$UserSign,
    [Parameter(Mandatory=$true)][int]$LeftTPL,
    [Parameter(Mandatory=$true)][int]$RightTPL
)
$ErrorActionPreference='Stop'
$conn=New-Object System.Data.SqlClient.SqlConnection "Server=$Server;Database=$CompanyDB;User ID=$DBUser;Password=$DBPassword;Connection Timeout=10"
$conn.Open()
function Run([string]$t,[string]$s){ Write-Host "`n=== $t ===" -ForegroundColor Cyan; $c=$conn.CreateCommand(); $c.CommandText=$s; $a=New-Object System.Data.SqlClient.SqlDataAdapter $c; $d=New-Object System.Data.DataTable; [void]$a.Fill($d); if($d.Rows.Count -eq 0){Write-Host "(no diff)" -ForegroundColor Green}else{$d|Format-Table -AutoSize|Out-Host; Write-Host "Total: $($d.Rows.Count) row(s)" -ForegroundColor Yellow} }

Run "Rows in TPLId=$LeftTPL but MISSING in TPLId=$RightTPL" @"
SELECT L.ItemID, L.ColID, L.VisInForm, L.VisualIndx, L.Width
FROM CPRF L
WHERE L.FormID='$FormID' AND L.UserSign=$UserSign AND L.TPLId=$LeftTPL
  AND NOT EXISTS (
    SELECT 1 FROM CPRF R
    WHERE R.FormID=L.FormID AND R.UserSign=L.UserSign AND R.TPLId=$RightTPL
      AND R.ItemID=L.ItemID AND R.ColID=L.ColID
  )
ORDER BY L.ItemID, L.ColID
"@

Run "Rows in TPLId=$RightTPL but MISSING in TPLId=$LeftTPL" @"
SELECT R.ItemID, R.ColID, R.VisInForm, R.VisualIndx, R.Width
FROM CPRF R
WHERE R.FormID='$FormID' AND R.UserSign=$UserSign AND R.TPLId=$RightTPL
  AND NOT EXISTS (
    SELECT 1 FROM CPRF L
    WHERE L.FormID=R.FormID AND L.UserSign=R.UserSign AND L.TPLId=$LeftTPL
      AND L.ItemID=R.ItemID AND L.ColID=R.ColID
  )
ORDER BY R.ItemID, R.ColID
"@

Run "Rows in BOTH but values DIFFER (VisInForm / VisualIndx / Width)" @"
SELECT L.ItemID, L.ColID,
       L.VisInForm  AS L_Vis,  R.VisInForm  AS R_Vis,
       L.VisualIndx AS L_Indx, R.VisualIndx AS R_Indx,
       L.Width      AS L_W,    R.Width      AS R_W
FROM CPRF L
INNER JOIN CPRF R
  ON R.FormID=L.FormID AND R.UserSign=L.UserSign AND R.TPLId=$RightTPL
 AND R.ItemID=L.ItemID AND R.ColID=L.ColID
WHERE L.FormID='$FormID' AND L.UserSign=$UserSign AND L.TPLId=$LeftTPL
  AND (L.VisInForm <> R.VisInForm OR L.VisualIndx <> R.VisualIndx OR L.Width <> R.Width)
ORDER BY L.ItemID, L.ColID
"@

$conn.Close()
Write-Host "`nDone." -ForegroundColor Green
