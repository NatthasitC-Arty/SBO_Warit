# ============================================================
# Convert annotated UDV .xlsx -> CSV (UTF-8 with BOM).
# Strips the column comments / formatting and produces a plain
# CSV that the Import script can read.
#
# Uses Excel COM (requires Excel installed).
# ============================================================
param(
    [string]$InputXlsx  = "",
    [string]$OutputCsv  = "",
    [string]$SheetName  = ""        # blank = first sheet
)

if (-not $InputXlsx) {
    Write-Host "ERROR: -InputXlsx required" -ForegroundColor Red
    exit 2
}
if (-not (Test-Path $InputXlsx)) {
    Write-Host "ERROR: Input not found: $InputXlsx" -ForegroundColor Red
    exit 2
}
$absIn = [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $InputXlsx).Path)

if (-not $OutputCsv) {
    $base = [System.IO.Path]::GetFileNameWithoutExtension($absIn)
    $OutputCsv = Join-Path (Split-Path $absIn -Parent) "$base.csv"
}
$absOut = [System.IO.Path]::GetFullPath($OutputCsv)
$dir = Split-Path $absOut -Parent
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

# ------------------------------------------------------------
# Launch Excel via COM
# ------------------------------------------------------------
try {
    $excel = New-Object -ComObject Excel.Application
} catch {
    Write-Host "ERROR: Cannot start Excel - is Microsoft Excel installed?" -ForegroundColor Red
    Write-Host $_.Exception.Message
    exit 3
}
$excel.Visible = $false
$excel.DisplayAlerts = $false

$wb = $null
$ws = $null
try {
    $wb = $excel.Workbooks.Open($absIn)

    # Pick sheet
    if ($SheetName) {
        try { $ws = $wb.Sheets.Item($SheetName) } catch {
            Write-Host "ERROR: Sheet '$SheetName' not found in $absIn" -ForegroundColor Red
            exit 4
        }
    } else {
        $ws = $wb.Sheets.Item(1)
    }
    $ws.Activate() | Out-Null
    Write-Host "Reading sheet: $($ws.Name)"

    # Save active sheet as CSV UTF-8 (with BOM)
    # xlCSVUTF8 = 62 (Excel 2016+), xlCSV = 6 (legacy ANSI)
    if (Test-Path $absOut) { Remove-Item $absOut -Force }
    try {
        $wb.SaveAs($absOut, 62)   # CSV UTF-8 with BOM
        Write-Host "Saved CSV (UTF-8 BOM): $absOut" -ForegroundColor Green
    } catch {
        Write-Host "xlCSVUTF8 (62) failed, trying xlCSV (6) + BOM injection..." -ForegroundColor Yellow
        $wb.SaveAs($absOut, 6)    # legacy CSV
        # Inject UTF-8 BOM manually
        $bytes = [System.IO.File]::ReadAllBytes($absOut)
        $bom = [byte[]](0xEF, 0xBB, 0xBF)
        if ($bytes.Length -lt 3 -or $bytes[0] -ne 0xEF -or $bytes[1] -ne 0xBB -or $bytes[2] -ne 0xBF) {
            $newBytes = New-Object byte[] ($bytes.Length + 3)
            [System.Array]::Copy($bom, 0, $newBytes, 0, 3)
            [System.Array]::Copy($bytes, 0, $newBytes, 3, $bytes.Length)
            [System.IO.File]::WriteAllBytes($absOut, $newBytes)
        }
        Write-Host "Saved CSV (ANSI -> BOM-prefixed): $absOut" -ForegroundColor Green
    }
} finally {
    if ($wb) { $wb.Close($false) | Out-Null }
    if ($excel) { $excel.Quit() }
    foreach ($x in @($ws, $wb, $excel)) {
        if ($x) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
exit 0
