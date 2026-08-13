# ============================================================
# Batch Import Crystal Layouts to SAP B1 (SQL Direct)
# Reads RPT_Import_Map.xlsx -> INSERT into RDOC
# Maps ObjectType -> TypeCode (Items variant by default)
# ============================================================
param(
    [string]$Server     = "SLD-C072",
    [string]$CompanyDB  = "SBO_SDA",
    [string]$DBUser     = "sa",
    [string]$DBPassword = "1q2w3e4r",
    [string]$MapFile    = "$PSScriptRoot\..\Config\RPT_Import_Map.xlsx",
    [string]$RptRoot    = "$PSScriptRoot\..\..",
    [string]$LogFile    = "$PSScriptRoot\..\Import_SQL_Log.txt",
    [string]$Author     = "SDA",
    [ValidateSet("Update","Skip","Insert")]
    [string]$OnDuplicate = "Update",
    [string]$FilterFileName = "",
    [switch]$UseFileNameAsDocName,
    [switch]$DryRun
)

# Load DB plugin
. "$PSScriptRoot\DB-MSSQL.ps1"

# ObjectType -> TypeCode (RTYP.CODE) mapping
$TypeCodeMap = @{
    "30"         = "JDT2"   # Journal Entry
    "23"         = "QUT2"   # Sales Quotation (Items)
    "17"         = "RDR2"   # Sales Order (Items)
    "15"         = "DLN2"   # Delivery (Items)
    "16"         = "RDN2"   # Returns (Items)
    "203"        = "DPI2"   # AR Down Payment (Items)
    "13"         = "INV2"   # AR Invoice (Items)
    "14"         = "RIN2"   # AR Credit Note (Items)
    "1470000113" = "PRQ2"   # Purchase Request (Items)
    "540000405"  = "PQT2"   # Purchase Quotation (Items)
    "22"         = "POR2"   # Purchase Order (Items)
    "20"         = "PDN2"   # Goods Receipt PO (Items)
    "21"         = "RPD2"   # Goods Return (Items)
    "204"        = "DPO2"   # AP Down Payment (Items)
    "18"         = "PCH2"   # AP Invoice (Items)
    "19"         = "RPC2"   # AP Credit Note (Items)
    "69"         = "IPF1"   # Landed Costs
    "24"         = "RCT1"   # Incoming Payment
    "46"         = "VPM1"   # Outgoing Payment
    "59"         = "IGN1"   # Goods Receipt
    "60"         = "IGE1"   # Goods Issue
    "67"         = "WTR1"   # Inventory Transfer
    "1250000001" = "WTQ1"   # Inventory Transfer Request
    "1470000065" = "INC1"   # Inventory Counting
    "162"        = ""       # Inventory Revaluation - no clear RTYP code, skip
    "202"        = "WOR1"   # Production Order
}

function Write-Log {
    param([string]$Msg, [string]$Level = "INFO")
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Msg
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

function Read-MapExcel {
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    function ColLetter-ToIndex([string]$letters) {
        $n = 0
        foreach ($c in $letters.ToCharArray()) { $n = $n * 26 + ([int][char]$c - 64) }
        return $n
    }

    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        # Shared strings
        $shared = @()
        $ssEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/sharedStrings.xml" } | Select-Object -First 1
        if ($ssEntry) {
            $sr = New-Object System.IO.StreamReader($ssEntry.Open())
            [xml]$ssXml = $sr.ReadToEnd(); $sr.Close()
            $ns = New-Object System.Xml.XmlNamespaceManager($ssXml.NameTable)
            $ns.AddNamespace("x","http://schemas.openxmlformats.org/spreadsheetml/2006/main")
            foreach ($si in $ssXml.SelectNodes("//x:si",$ns)) {
                # concatenate all <t> nodes (handles rich text)
                $txt = ""
                foreach ($t in $si.SelectNodes(".//x:t",$ns)) { $txt += $t.InnerText }
                $shared += ,$txt
            }
        }

        # Workbook -> find sheet "RPT_MAP" rId
        $wbEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/workbook.xml" } | Select-Object -First 1
        $sr = New-Object System.IO.StreamReader($wbEntry.Open())
        [xml]$wbXml = $sr.ReadToEnd(); $sr.Close()
        $nsw = New-Object System.Xml.XmlNamespaceManager($wbXml.NameTable)
        $nsw.AddNamespace("x","http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        $nsw.AddNamespace("r","http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        $sheetNode = $wbXml.SelectSingleNode("//x:sheet[@name='RPT_MAP']",$nsw)
        if (-not $sheetNode) { throw "Sheet 'RPT_MAP' not found in $Path" }
        $rid = $sheetNode.GetAttribute("id","http://schemas.openxmlformats.org/officeDocument/2006/relationships")

        # Relationships -> sheet target
        $relsEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/_rels/workbook.xml.rels" } | Select-Object -First 1
        $sr = New-Object System.IO.StreamReader($relsEntry.Open())
        [xml]$relsXml = $sr.ReadToEnd(); $sr.Close()
        $target = ($relsXml.Relationships.Relationship | Where-Object { $_.Id -eq $rid }).Target
        if ($target -notmatch "^/") { $target = "xl/$target" } else { $target = $target.TrimStart('/') }

        # Sheet XML
        $sheetEntry = $zip.Entries | Where-Object { $_.FullName -eq $target } | Select-Object -First 1
        $sr = New-Object System.IO.StreamReader($sheetEntry.Open())
        [xml]$shXml = $sr.ReadToEnd(); $sr.Close()
        $nss = New-Object System.Xml.XmlNamespaceManager($shXml.NameTable)
        $nss.AddNamespace("x","http://schemas.openxmlformats.org/spreadsheetml/2006/main")

        $list = @()
        foreach ($row in $shXml.SelectNodes("//x:sheetData/x:row",$nss)) {
            $rowIdx = [int]$row.r
            if ($rowIdx -lt 2) { continue }
            $cells = @{}
            foreach ($c in $row.SelectNodes("x:c",$nss)) {
                $ref = $c.r
                $letters = ($ref -replace '[0-9]','')
                $colIdx = ColLetter-ToIndex $letters
                $t = $c.t
                $vNode = $c.SelectSingleNode("x:v",$nss)
                $isNode = $c.SelectSingleNode("x:is",$nss)
                $val = $null
                if ($t -eq "s" -and $vNode) {
                    $idx = [int]$vNode.InnerText
                    if ($idx -lt $shared.Count) { $val = $shared[$idx] }
                } elseif ($t -eq "inlineStr" -and $isNode) {
                    $val = ""
                    foreach ($tt in $isNode.SelectNodes(".//x:t",$nss)) { $val += $tt.InnerText }
                } elseif ($vNode) {
                    $val = $vNode.InnerText
                }
                $cells[$colIdx] = $val
            }
            $item = [PSCustomObject]@{
                No           = $cells[1]
                Module       = $cells[2]
                RPT_FileName = $cells[3]
                RPT_Folder   = $cells[4]
                ObjectType   = $cells[8]
                FormMenuUID  = $cells[9]
                LayoutName   = $cells[10]
            }
            if ($item.RPT_FileName) { $list += $item }
        }
        return $list
    } finally {
        $zip.Dispose()
    }
}

Write-Log "=== Start SQL Direct Import ==="
Write-Log "MapFile: $MapFile"
Write-Log "DryRun : $DryRun"

$mapRows = Read-MapExcel -Path $MapFile
Write-Log "Loaded $($mapRows.Count) mapping rows"

$conn = New-DBConnection -Server $Server -Database $CompanyDB -User $DBUser -Password $DBPassword

# Pre-fetch max sequence per TypeCode to avoid query per row
$maxSeqMap = @{}
$seqCmd = $conn.CreateCommand()
$seqCmd.CommandText = "SELECT TypeCode, MAX(CAST(SUBSTRING(DocCode, LEN(TypeCode)+1, 4) AS INT)) AS MaxSeq FROM RDOC WHERE LEN(DocCode)>=5 AND $DB_ISNUM(SUBSTRING(DocCode, LEN(TypeCode)+1, 4))=1 GROUP BY TypeCode"
$rdr = $seqCmd.ExecuteReader()
while ($rdr.Read()) {
    $tc = [string]$rdr["TypeCode"]
    $maxSeqMap[$tc] = [int]$rdr["MaxSeq"]
}
$rdr.Close()
Write-Log "Pre-loaded max sequences for $($maxSeqMap.Count) TypeCodes"

$ok = 0; $fail = 0; $skip = 0
foreach ($row in $mapRows) {
    if ($FilterFileName -and ($row.RPT_FileName -notlike "*$FilterFileName*")) { continue }
    $rptPath = Join-Path $RptRoot (Join-Path $row.RPT_Folder $row.RPT_FileName)
    $objType = [string]$row.ObjectType
    if ($UseFileNameAsDocName) {
        $layoutName = [System.IO.Path]::GetFileNameWithoutExtension($row.RPT_FileName)
    } else {
        $layoutName = if ([string]::IsNullOrWhiteSpace($row.LayoutName)) { $row.RPT_FileName } else { [string]$row.LayoutName }
    }

    if (-not (Test-Path $rptPath)) {
        Write-Log "SKIP MISSING file: $rptPath" "WARN"; $skip++; continue
    }
    if ([string]::IsNullOrWhiteSpace($objType) -or $objType -eq "-") {
        Write-Log "SKIP no ObjectType: $($row.RPT_FileName)" "WARN"; $skip++; continue
    }
    $typeCode = $TypeCodeMap[$objType]
    if ([string]::IsNullOrWhiteSpace($typeCode)) {
        Write-Log "SKIP unmapped ObjectType=$objType : $($row.RPT_FileName)" "WARN"; $skip++; continue
    }

    try {
        $bytes = [System.IO.File]::ReadAllBytes($rptPath)
        $md5 = [System.Security.Cryptography.MD5]::Create()
        $hash = [BitConverter]::ToString($md5.ComputeHash($bytes)).Replace("-","")
        $md5.Dispose()

        # Check if this layout already exists (match on DocName + TypeCode only --
        # Author is intentionally ignored so re-imports update across owners).
        $chk = $conn.CreateCommand()
        $chk.CommandText = Convert-DBSql "SELECT DocCode FROM RDOC WHERE DocName=@n AND TypeCode=@t"
        Add-DBParam $chk "${DB_PARAM}n" $layoutName
        Add-DBParam $chk "${DB_PARAM}t" $typeCode
        $existingCode = $chk.ExecuteScalar()

        $action = ""; $docCode = ""
        if ($existingCode) {
            switch ($OnDuplicate) {
                "Skip"   { Write-Log ("SKIP exists [{0,3}] {1} (DocCode={2})" -f $row.No, $row.RPT_FileName, $existingCode) "WARN"; $skip++; continue }
                "Insert" { $maxSeqMap[$typeCode]++; $docCode = "{0}{1:D4}" -f $typeCode, $maxSeqMap[$typeCode]; $action = "INSERT (dup allowed)" }
                "Update" { $docCode = [string]$existingCode; $action = "UPDATE" }
            }
        } else {
            if (-not $maxSeqMap.ContainsKey($typeCode)) { $maxSeqMap[$typeCode] = 0 }
            $maxSeqMap[$typeCode]++
            $docCode = "{0}{1:D4}" -f $typeCode, $maxSeqMap[$typeCode]
            $action = "INSERT"
        }

        if ($DryRun) {
            Write-Log ("DRYRUN [{0,3}] {1} -> {2} DocCode={3} TypeCode={4} Bytes={5} Hash={6}" -f $row.No, $row.RPT_FileName, $action, $docCode, $typeCode, $bytes.Length, $hash.Substring(0,8))
            $ok++; continue
        }

        if ($action -eq "UPDATE") {
            $sql = "UPDATE RDOC SET Template=@Template, RptHash=@RptHash, UpdateDate=$DB_NOW WHERE DocCode=@DocCode"
        } else {
            $sql = "INSERT INTO RDOC (DocCode,DocName,Author,Notes,Width,Height,LMargin,RMargin,TMargin,BMargin,CanChange,PaperSize,Oreint,GridSize,GridType,ShowGrid,SnapGrid,TypeCode,FrgnReport,CanSort,LeaderCode,FollowCode,SwapOnScrn,ScreenFont,ScrFOffset,SwpInEmail,EmailFont,EmFOffset,QString,QType,RobjCode,ExtName,ExtOnErr,NumRepArs,AlgnFooter,TimeFormat,DateFormat,NumCopy,GbiSupport,Use1stPrtr,Shading,Template,Category,CreateDate,Status,B1Version,CRVersion,Local,UseSysPref,ForMobile,TypeDetail,IsIMCE,CsUrl,RptHash) VALUES (@DocCode,@DocName,@Author,'',595,842,10,30,10,10,'Y','A4','P',10,'1','Y','Y',@TypeCode,'N','Y','','','N','Arial',-1,'N','Arial',-1,'','R',0,'','S',-1,'N','0','0',1,'N','N','Y',@Template,'C',$DB_NOW,'A','','','','Y','Y','','N','',@RptHash)"
        }

        $cmd = $conn.CreateCommand()
        $cmd.CommandText = Convert-DBSql $sql
        Add-DBParam $cmd "${DB_PARAM}DocCode" $docCode
        if ($action -ne "UPDATE") {
            Add-DBParam $cmd "${DB_PARAM}DocName" $layoutName
            Add-DBParam $cmd "${DB_PARAM}Author" $Author
            Add-DBParam $cmd "${DB_PARAM}TypeCode" $typeCode
        }
        Add-BlobParam $cmd "${DB_PARAM}Template" $bytes
        Add-DBParam $cmd "${DB_PARAM}RptHash" $hash

        [void]$cmd.ExecuteNonQuery()
        Write-Log ("{0} [{1,3}] {2} -> DocCode={3} ({4} bytes, {5})" -f $action, $row.No, $row.RPT_FileName, $docCode, $bytes.Length, $layoutName)
        $ok++
    } catch {
        Write-Log ("FAIL [{0,3}] {1}: {2}" -f $row.No, $row.RPT_FileName, $_.Exception.Message) "ERROR"
        $fail++
        # Roll back the seq increment if it wasn't actually used
        $maxSeqMap[$typeCode]--
    }
}
$conn.Close()
Write-Log "=== Summary: OK=$ok FAIL=$fail SKIP=$skip ==="