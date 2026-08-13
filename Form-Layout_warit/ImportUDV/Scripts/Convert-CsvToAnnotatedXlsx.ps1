# ============================================================
# Convert a UDV CSV to an annotated .xlsx file with column comments.
# The comments are the yellow-note popups in Excel when you hover
# over a column header — same as the screenshot.
#
# Uses Excel COM (requires Excel installed). Run from any PowerShell
# arch — Excel COM works both x64 and x86.
# ============================================================
param(
    [string]$InputCsv  = "$PSScriptRoot\..\Config\UDV_SO_Match_SQ.csv",
    [string]$OutputXlsx = ""
)

if (-not $OutputXlsx) {
    $base = [System.IO.Path]::GetFileNameWithoutExtension($InputCsv)
    $OutputXlsx = Join-Path (Split-Path $InputCsv -Parent) "$base.xlsx"
}

# Column comments (Thai-friendly explanations)
$colDescriptions = [ordered]@{
    'Action'        = "คำสั่ง:`nADD = สร้างใหม่`nUPDATE = แก้ของเดิม`nDELETE = ลบ`nUPSERT = สร้างถ้าไม่มี / แก้ถ้ามี (แนะนำ)"
    'FormID'        = "ID ของ Form ใน B1`nหาจาก View > System Information (Ctrl+Shift+I)`n`nตัวอย่าง:`n139 = Sales Order`n149 = AR Credit Memo`n133 = AR Invoice`n140 = Purchase Order"
    'ItemID'        = "ItemUID ของ field/matrix บน Form`nหาจาก System Info ที่ status bar`n`nตัวอย่าง:`n38 = matrix ของ Sales/AR docs`n76 = matrix ของ Journal Entry"
    'ColumnID'      = "ColUID ของ matrix column`n`nค่า:`n- ตัวเลข (เช่น 11, 14, 15) = system column`n- UDF code (เช่น U_SLD_T_BeDis) = UDF column`n- ว่าง / 0 = header field (ไม่ใช่ matrix)"
    'FMSAction'     = "ประเภท FMS:`nQ = ใช้ Saved Query`nF = ค่าคงที่ (Fixed Value)"
    'QueryName'     = "ชื่อ query ใน OUQR`n`nLogic:`n- เจอใน OUQR.QName -> reuse IntrnalKey + UPDATE QString`n- ไม่เจอ -> INSERT row ใหม่ใน OUQR`n`nWARNING: ชื่อซ้ำ = share ทุก form ที่ใช้"
    'QueryCategory' = "Category ของ query`n`n-1 = General (ใช้บ่อยสุด)`nอื่นๆ = IntrnalKey ของ category"
    'QueryBody'     = "SQL จริงของ Saved Query`n`nToken syntax:`n`$[`$ItemID.Col.Type] = matrix col`n`$[`$ItemID.U_xxx.Type] = matrix UDF col`n`$[Table.Field] = master/header ref`n`nType: 0 = raw, NUMBER = numeric, DATE = date"
    'FixedValue'    = "ค่าคงที่ (เมื่อ FMSAction=F)`nเช่น 'DefaultRemark'`n`nNote: เก็บใน v10 ยังไม่ชัดเจน column ไหน — F-type ส่วนใหญ่เป็น placeholder"
    'Refresh'       = "Auto Refresh:`nY = refresh เมื่อ trigger field เปลี่ยน`nN = ต้องกด Shift+F2 เอง`n`nสัมพันธ์กับ TriggerID + SHS1"
    'ByField'       = "DB flag ใน CSHS:`nN = ปกติ`nY = trigger by specific field`nC = checkbox state (กรณีพิเศษ)`n`nDefault: N"
    'TriggerID'     = "ชื่อ trigger field`n`nMulti-trigger pattern:`n1 FMS หลาย triggers = หลาย rows ใน CSV`nที่มี (FormID, ItemID, ColumnID) เดียวกัน`n`nเก็บใน SHS1.FieldID (ไม่ใช่ CSHS.FieldID)`n`nค่า: ตัวเลข (col) หรือ UDF code"
    'TriggerColumn' = "Legacy column - ไม่ใช้ใน B1 v10`nปล่อยว่างเสมอ"
    'ForceRefresh'  = "Display Saved Values:`nY = แสดงค่าทันทีเมื่อเปิด form`nN = แสดงว่าง user ต้องกด refresh`n`nBest: Y สำหรับ calculation field"
}

if (-not (Test-Path $InputCsv)) {
    Write-Host "ERROR: Input CSV not found: $InputCsv" -ForegroundColor Red
    exit 2
}

Write-Host "Reading: $InputCsv"
$rows = @(Import-Csv -Path $InputCsv)
if ($rows.Count -eq 0) {
    Write-Host "ERROR: CSV has no data rows" -ForegroundColor Red
    exit 2
}
$headers = @($rows[0].PSObject.Properties.Name)
Write-Host "Rows: $($rows.Count), Columns: $($headers.Count)"

# ------------------------------------------------------------
# Launch Excel via COM
# ------------------------------------------------------------
try {
    $excel = New-Object -ComObject Excel.Application
} catch {
    Write-Host "ERROR: Cannot start Excel — is Microsoft Excel installed?" -ForegroundColor Red
    Write-Host $_.Exception.Message
    exit 3
}
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Sheets.Item(1)
    $ws.Name = "UDV"

    # Header row with comments
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $col = $i + 1
        $cell = $ws.Cells.Item(1, $col)
        $cell.Value2 = $headers[$i]
        $cell.Font.Bold = $true
        $cell.Interior.Color = 13434879   # light yellow (RGB 255,255,204)
        if ($colDescriptions.Contains($headers[$i])) {
            $note = $colDescriptions[$headers[$i]]
            $cmt = $cell.AddComment($note)
            $cmt.Shape.TextFrame.AutoSize = $true
        }
    }

    # Data rows
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $row = $rows[$r]
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $val = [string]$row.($headers[$c])
            # Prepend ' for cells starting with = + - @ to keep them literal
            if ($val -match '^[=\+\-@]') { $val = "'" + $val }
            $ws.Cells.Item($r + 2, $c + 1).Value2 = $val
        }
    }

    # Freeze top row
    $ws.Range("A2").Select() | Out-Null
    $excel.ActiveWindow.FreezePanes = $true

    # Auto-fit columns (cap width so QueryBody column doesn't get too wide)
    $usedRange = $ws.UsedRange
    [void]$usedRange.Columns.AutoFit()
    for ($i = 1; $i -le $headers.Count; $i++) {
        $w = $ws.Columns.Item($i).ColumnWidth
        if ($w -gt 50) { $ws.Columns.Item($i).ColumnWidth = 50 }
    }

    # Save as xlsx (51 = xlOpenXMLWorkbook)
    $absOut = [System.IO.Path]::GetFullPath($OutputXlsx)
    $dir = Split-Path $absOut -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    if (Test-Path $absOut) { Remove-Item $absOut -Force }
    $wb.SaveAs($absOut, 51)
    Write-Host "Created: $absOut" -ForegroundColor Green
    Write-Host "Open in Excel to see column header comments (hover over headers)"
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
