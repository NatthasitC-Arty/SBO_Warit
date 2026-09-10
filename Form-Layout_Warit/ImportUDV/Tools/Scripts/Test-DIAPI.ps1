# ============================================================
# Verify DI API installation + B1 login. Does NOT modify data.
# Auto-detects arch (B1 v10 = 64-bit, v9.x = 32-bit) — call from
# the matching PowerShell (System32 for x64, SysWOW64 for x86).
# ============================================================
param(
    [string]$Server      = "SLD-C072",
    [string]$CompanyDB   = "SBO_SDA",
    [string]$DBUser      = "sa",
    [string]$DBPassword  = "1q2w3e4r",
    [string]$SapUser     = "manager",
    [string]$SapPassword = "1q2w3e4r",
    [ValidateSet("MSSQL","HANA")]
    [string]$DBType      = "MSSQL"
)

$is64 = [Environment]::Is64BitProcess
Write-Host "Process bitness  : $(if($is64){'x64'}else{'x86'})"

# Build search roots based on bitness — a 32-bit process cannot
# load a 64-bit Interop wrapper (and vice versa).
$searchRoots = @()
if ($is64) {
    $searchRoots += "${env:ProgramFiles}\SAP"
    $searchRoots += "${env:ProgramW6432}\SAP"
} else {
    $searchRoots += "${env:ProgramFiles(x86)}\SAP"
}
$searchRoots = $searchRoots | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique
if (-not $searchRoots) {
    Write-Host "No SAP install folders found for this bitness." -ForegroundColor Red
    exit 2
}
Write-Host "Search roots     : $($searchRoots -join '; ')"

$dll = $null
foreach ($r in $searchRoots) {
    $hit = Get-ChildItem -Path $r -Recurse -Filter "Interop.SAPbobsCOM.dll" -ErrorAction SilentlyContinue |
        Sort-Object -Property @{Expression = { $_.FullName -match "DI API" }; Descending = $true} |
        Select-Object -First 1 -ExpandProperty FullName
    if ($hit) { $dll = $hit; break }
}
if (-not $dll) {
    Write-Host "Interop.SAPbobsCOM.dll not found under SAP folders" -ForegroundColor Red
    Write-Host "Try the OTHER bat (the matching arch). If still missing, install Data Transfer Workbench." -ForegroundColor Yellow
    exit 2
}
Write-Host "DI API DLL       : $dll"

try {
    Add-Type -Path $dll -ErrorAction Stop
} catch {
    Write-Host "Add-Type FAILED: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Architecture mismatch — relaunch via the other-arch PowerShell." -ForegroundColor Yellow
    exit 2
}
$company = New-Object -ComObject SAPbobsCOM.Company
$company.Server     = $Server
$company.CompanyDB  = $CompanyDB
$company.UserName   = $SapUser
$company.Password   = $SapPassword
$company.DbUserName = $DBUser
$company.DbPassword = $DBPassword
$company.UseTrusted = $false
$company.language   = [SAPbobsCOM.BoSuppLangs]::ln_English

if ($DBType -eq "HANA") {
    $company.DbServerType = [SAPbobsCOM.BoDataServerTypes]::dst_HANADB
    $rc = $company.Connect()
    if ($rc -ne 0) { Write-Host "Connect FAILED: $($company.GetLastErrorDescription())" -ForegroundColor Red; exit 3 }
} else {
    $tries = @(
        [SAPbobsCOM.BoDataServerTypes]::dst_MSSQL2019,
        [SAPbobsCOM.BoDataServerTypes]::dst_MSSQL2017,
        [SAPbobsCOM.BoDataServerTypes]::dst_MSSQL2016,
        [SAPbobsCOM.BoDataServerTypes]::dst_MSSQL2014,
        [SAPbobsCOM.BoDataServerTypes]::dst_MSSQL2012
    )
    $connected = $false
    foreach ($t in $tries) {
        $company.DbServerType = $t
        if ($company.Connect() -eq 0) {
            Write-Host "Connected as     : DbServerType=$t" -ForegroundColor Green
            $connected = $true; break
        }
    }
    if (-not $connected) { Write-Host "Connect FAILED: $($company.GetLastErrorDescription())" -ForegroundColor Red; exit 3 }
}

Write-Host "Company name     : $($company.CompanyName)"
Write-Host "Company version  : $($company.Version)"

# Read 1 row from CSHS to confirm read access
$rs = $company.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::BoRecordset)
try {
    $rs.DoQuery("SELECT COUNT(*) AS C FROM CSHS")
    Write-Host "CSHS row count   : $($rs.Fields.Item(0).Value)"
} finally {
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($rs)
}

$company.Disconnect() | Out-Null
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($company)
Write-Host "All checks passed." -ForegroundColor Green
