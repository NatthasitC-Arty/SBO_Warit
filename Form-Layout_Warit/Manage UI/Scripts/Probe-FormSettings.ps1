# ============================================================
# Probe SAP B1 schema to find tables that store matrix Form Settings
# (column visibility, order, width per FormID/ItemID/ColID).
#
# B1 stores: FMS bindings  -> CSHS (already known)
#            UDF metadata  -> CUFD
#            Column profile / form settings -> ??? (need to discover)
#
# Looks for tables with FormID/ItemID/ColID-like columns + a
# visible/order/active-like column.
# ============================================================
param(
    [string]$Server      = "10.10.10.115",
    [string]$CompanyDB   = "SBO_SDA_Training",
    [string]$DBUser      = "sa",
    [Parameter(Mandatory=$true)][string]$DBPassword
)

$ErrorActionPreference = "Stop"
$cs = "Server=$Server;Database=$CompanyDB;User ID=$DBUser;Password=$DBPassword;Connection Timeout=15;"
$conn = New-Object System.Data.SqlClient.SqlConnection $cs
$conn.Open()

function Run-Query {
    param([string]$Title, [string]$Sql)
    Write-Host ""
    Write-Host "=== $Title ===" -ForegroundColor Cyan
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $Sql
    $da  = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
    $dt  = New-Object System.Data.DataTable
    [void]$da.Fill($dt)
    if ($dt.Rows.Count -eq 0) { Write-Host "(no matches)" -ForegroundColor DarkGray }
    else { $dt | Format-Table -AutoSize | Out-Host }
}

# 1) Tables that have FormID-like + visibility/order columns
Run-Query "Tables with FormID + (Visible|Active|Order|Position|ColPos|Width)" @"
SELECT DISTINCT t.name AS TableName
FROM sys.tables t
INNER JOIN sys.columns c1 ON c1.object_id=t.object_id AND c1.name IN ('FormID','Form_ID','FormCode','FormType')
INNER JOIN sys.columns c2 ON c2.object_id=t.object_id
WHERE c2.name IN ('Visible','Active','Order','Position','ColPos','Width','Visibility','Editable','ColUID','ColID')
ORDER BY t.name
"@

# 2) Any table with both ItemID and ColID-like columns (matrix-level)
Run-Query "Tables with ItemID + ColID/ColUID columns" @"
SELECT t.name AS TableName,
       STUFF((SELECT ', ' + c.name FROM sys.columns c WHERE c.object_id=t.object_id ORDER BY c.column_id FOR XML PATH('')),1,2,'') AS Columns
FROM sys.tables t
WHERE EXISTS (SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name IN ('ItemID','ItemCode'))
  AND EXISTS (SELECT 1 FROM sys.columns c WHERE c.object_id=t.object_id AND c.name IN ('ColID','ColUID','ColumnID'))
ORDER BY t.name
"@

# 3) Known B1 form-related table-name patterns
Run-Query "Tables matching CPRF|OFRM|WDD|OFML|CFRM|CFG_|*Form* patterns" @"
SELECT t.name AS TableName,
       (SELECT COUNT(*) FROM sys.columns c WHERE c.object_id=t.object_id) AS Cols
FROM sys.tables t
WHERE t.name IN ('CPRF','OFRM','WDD1','OFML','CFRM','CFG_FORM','FORMS')
   OR t.name LIKE '%FORM%' OR t.name LIKE '%PROFILE%' OR t.name LIKE '%MTRX%' OR t.name LIKE '%MATRIX%'
ORDER BY t.name
"@

# 4) Sample peek: count rows in CSHS-related and any candidate found
Run-Query "Row counts of candidate tables" @"
SELECT 'CSHS' AS Tbl, (SELECT COUNT(*) FROM CSHS) AS Rows_
UNION ALL SELECT 'CUFD', (SELECT COUNT(*) FROM CUFD)
UNION ALL SELECT 'OUQR', (SELECT COUNT(*) FROM OUQR)
"@

$conn.Close()
Write-Host ""
Write-Host "Done. Send the output back so I can pick the right table." -ForegroundColor Green
