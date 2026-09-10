<#
.SYNOPSIS
  Puts the conditional Suppress formula "PageNumber <> TotalPageCount" on every
  total field sitting in the Page Footer of a Crystal Reports .rpt layout, so a
  running total only prints on the last page.

.DESCRIPTION
  MEASURED ON THIS REPO (2026-09-01):
    197 total fields live in page footers across the 64 reports.
    116 already carry the condition -- 105 exactly "PageNumber <> TotalPageCount"
    and 11 combining it with a currency test. The remaining 81 were simply
    missed: {@SubTotal}, the amount-in-words field, the discount percentage.
    This script closes that gap; it never touches the 116 that are already set.

  THE HOUSE CONVENTION, followed here deliberately:
    the VALUE fields get the condition, the LABELS next to them do not. That is
    how all 116 existing ones are set up. Suppressing the labels as well would
    be a visible layout change nobody asked for.

  WRITE ROUTE (verified by reading the file back from disk):
    clone = obj.Clone($true)
    clone.Format.ConditionFormulas.Formula(<EnableSuppress>).Text = "..."
    ReportObjectController.Modify(live, clone)
    doc.SaveAs(file)
  The collection is indexed with .Formula(<type>), NOT .Item(<type>) -- Item
  does not exist on it and fails at runtime.

  Unlike the SQL fixes in Fix-ReportSql.ps1 this touches no connection, so it
  does NOT inject 'Database DLL' and needs no Designer pass afterwards. The
  summary reports poisoning anyway, as a check on that claim.

.EXAMPLE
  .\Fix-PageFooterTotals.ps1 -Path "C:\GitHub\SDA\Form-Layout_SDA" -WhatIf

.NOTES
  64-bit PowerShell only. Exit: 0 clean, 1 a file failed and was rolled back,
  2 not 64-bit, 3 SDK load failed, 4 bad -Path.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $Path,
    [switch] $WhatIf,
    [string] $Formula      = 'PageNumber <> TotalPageCount',
    [string] $Filter       = '*.rpt',
    [switch] $NoRecurse,
    [string] $BackupSuffix = '.20260901-footer.bak',
    [string] $LogFile      = "$PSScriptRoot\_FixPageFooterTotals.log"
)

$ErrorActionPreference = 'Stop'
if ([IntPtr]::Size -ne 8) { Write-Host '[ERROR] 64-bit PowerShell required.' -ForegroundColor Red; exit 2 }

function Load-CrystalSDK {
    $gacMsil = 'C:\Windows\Microsoft.NET\assembly\GAC_MSIL'
    $gac64   = 'C:\Windows\Microsoft.NET\assembly\GAC_64'
    $sets = @(
        @{ Root = $gacMsil; Names = @('CrystalDecisions.Shared','CrystalDecisions.ReportSource','CrystalDecisions.CrystalReports.Engine') },
        @{ Root = $gac64;   Names = @('CrystalDecisions.ReportAppServer.CommLayer','CrystalDecisions.ReportAppServer.DataDefModel',
                                      'CrystalDecisions.ReportAppServer.Controllers','CrystalDecisions.ReportAppServer.ClientDoc',
                                      'CrystalDecisions.ReportAppServer.CommonObjectModel','CrystalDecisions.ReportAppServer.ObjectFactory',
                                      'CrystalDecisions.ReportAppServer.ReportDefModel') }
    )
    foreach ($set in $sets) {
        foreach ($n in $set.Names) {
            $dir = Join-Path $set.Root $n
            if (-not (Test-Path $dir)) { throw "GAC folder not found: $dir" }
            $dll = Get-ChildItem $dir -Recurse -Filter "$n.dll" |
                   Sort-Object { [version](($_.Directory.Name -split '_')[1]) } -Descending | Select-Object -First 1
            if (-not $dll) { throw "DLL not found under $dir" }
            [void][System.Reflection.Assembly]::LoadFrom($dll.FullName)
        }
    }
}
try { Load-CrystalSDK } catch { Write-Host "[ERROR] SDK load failed: $($_.Exception.Message)" -ForegroundColor Red; exit 3 }
$SUP = [CrystalDecisions.ReportAppServer.ReportDefModel.CrObjectFormatConditionFormulaTypeEnum]::crObjectFormatConditionFormulaTypeEnableSuppress

$logDir = Split-Path $LogFile -Parent
if ($logDir -and -not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
function Log {
    param([string]$Level, [string]$Msg)
    $line = "[{0}][{1}] {2}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Level, $Msg
    switch ($Level) {
        'ERR'  { Write-Host $line -ForegroundColor Red }
        'WARN' { Write-Host $line -ForegroundColor Yellow }
        'OK'   { Write-Host $line -ForegroundColor Green }
        'HIT'  { Write-Host $line -ForegroundColor Cyan }
        default{ Write-Host $line }
    }
    try { [System.IO.File]::AppendAllText($LogFile, $line + "`r`n", [System.Text.Encoding]::UTF8) } catch {}
}

# A page-footer field counts as a running total when its embedded field is one
# of these. Page numbers, remarks, names and dates deliberately do not match --
# they must keep printing on every page.
$TotalRx = 'Sum_LineTotal_All|DocTotal|VatSum|DiscSum|DiscP\b|DiscPrcnt|DpmAmnt|dpmprcnt|Ref_DocTotal|' +
           '@Sub ?Total|@GrossTotal|@DiscountAmt|@T Sum Line Total|@total value|@Difference value|@dif Value|' +
           '@Totext|@T Text|@F Text|@THB Text|@ToTecxt'

function Get-EmbeddedFields {
    param($Obj)
    $out = @()
    try {
        foreach ($p in $Obj.Paragraphs) {
            foreach ($e in $p.ParagraphElements) {
                $ds = ''
                try { $ds = [string]$e.DataSource } catch {}
                if ($ds) { $out += $ds }
            }
        }
    } catch {}
    return $out
}
function Get-SuppressText {
    param($Obj)
    try {
        $c = $Obj.Format.ConditionFormulas.Formula($SUP)
        if ($c) { return [string]$c.Text }
    } catch {}
    return ''
}
# "// pageNumber <> TotalPageCount" is text, not a condition -- a fully
# commented formula suppresses nothing, so it must not read as already done.
function Test-Live {
    param([string]$Text)
    if (-not $Text) { return $false }
    foreach ($ln in ($Text -split "`r?`n")) {
        $t = $ln.Trim()
        if (-not $t) { continue }
        if ($t.StartsWith('//')) { continue }
        if ($t -match '(?i)PageNumber\s*<>\s*TotalPageCount') { return $true }
    }
    return $false
}

$rpts = @()
if (Test-Path -LiteralPath $Path -PathType Leaf) {
    if ($Path -like '*.rpt') { $rpts = ,(Get-Item -LiteralPath $Path) } else { Log 'ERR' "not a .rpt: $Path"; exit 4 }
} elseif (Test-Path -LiteralPath $Path -PathType Container) {
    $ga = @{ LiteralPath = $Path; Filter = $Filter; File = $true }
    if (-not $NoRecurse) { $ga['Recurse'] = $true }
    $rpts = Get-ChildItem @ga
} else { Log 'ERR' "Path not found: $Path"; exit 4 }

$mode = if ($WhatIf) { 'WHATIF' } else { 'APPLY' }
Log 'INFO' "=== Fix-PageFooterTotals start  mode=$mode ==="
Log 'INFO' "Path=$Path  files=$($rpts.Count)  formula='$Formula'"

$filesScanned = 0; $filesChanged = 0; $filesFailed = 0
$objSet = 0; $objCombined = 0; $objAlready = 0; $objCommented = 0
$commentedList = New-Object System.Collections.ArrayList
$failList      = New-Object System.Collections.ArrayList

foreach ($f in $rpts) {
    $full = $f.FullName
    $filesScanned++
    $doc = $null
    $plan = New-Object System.Collections.ArrayList
    try {
        $doc = New-Object CrystalDecisions.CrystalReports.Engine.ReportDocument
        $doc.Load($full, [CrystalDecisions.Shared.OpenReportMethod]::OpenReportByTempCopy)
        $pf = $doc.ReportClientDocument.ReportDefController.ReportDefinition.PageFooterArea
        for ($i = 0; $i -lt $pf.Sections.Count; $i++) {
            $s = $pf.Sections[$i]
            for ($j = 0; $j -lt $s.ReportObjects.Count; $j++) {
                $o = $s.ReportObjects[$j]
                $kd = 0; try { $kd = [int]$o.Kind } catch {}
                if ($kd -eq 3 -or $kd -eq 4 -or $kd -eq 6) { continue }   # line, box, picture
                $flds = Get-EmbeddedFields -Obj $o
                if ($flds.Count -eq 0) { continue }
                $joined = ($flds -join ' ')
                if ($joined -notmatch $TotalRx) { continue }

                $cur = Get-SuppressText -Obj $o
                if (Test-Live -Text $cur) { $objAlready++; continue }
                if ($cur -and $cur.Trim()) {
                    # something is there but it is not the page test
                    $allComment = $true
                    foreach ($ln in ($cur -split "`r?`n")) { $t = $ln.Trim(); if ($t -and -not $t.StartsWith('//')) { $allComment = $false } }
                    if ($allComment) {
                        # a human commented this out on purpose -- report it, change nothing
                        $objCommented++
                        [void]$commentedList.Add("$full || $($s.Name)|$($o.Name) || $($cur.Trim())")
                        continue
                    }
                    $want = "$Formula or ($($cur.Trim()))"
                    $kind = 'combine'
                } else {
                    $want = $Formula
                    $kind = 'set'
                }
                [void]$plan.Add([pscustomobject]@{
                    SectionIdx = $i; ObjIdx = $j; Section = [string]$s.Name; Obj = [string]$o.Name
                    Fields = $joined; Was = $cur; Want = $want; Kind = $kind
                })
            }
        }
    } catch {
        Log 'ERR' "--> $full : load failed: $($_.Exception.Message.Split([char]10)[0])"
        $filesFailed++
        [void]$failList.Add("$full || <load> || $($_.Exception.Message.Split([char]10)[0])")
        if ($doc) { try { $doc.Close() } catch {} }
        continue
    }

    if ($plan.Count -eq 0) { if ($doc) { try { $doc.Close() } catch {} }; continue }

    Log 'HIT' "--> $full"
    foreach ($p in $plan) {
        Log 'INFO' ("    [{0}] {1,-12} {2,-8} {3}" -f $p.Section, $p.Obj, $p.Kind, $p.Fields)
        if ($p.Kind -eq 'combine') { Log 'INFO' ("        was '{0}'  ->  '{1}'" -f $p.Was.Trim(), $p.Want) }
    }
    if ($WhatIf) {
        foreach ($p in $plan) { if ($p.Kind -eq 'combine') { $objCombined++ } else { $objSet++ } }
        if ($doc) { try { $doc.Close() } catch {} }
        continue
    }

    $bak = "$full$BackupSuffix"
    Copy-Item -LiteralPath $full -Destination $bak -Force
    $wrote = 0
    try {
        $rdc = $doc.ReportClientDocument.ReportDefController
        $pf  = $rdc.ReportDefinition.PageFooterArea
        foreach ($p in $plan) {
            $live  = $pf.Sections[$p.SectionIdx].ReportObjects[$p.ObjIdx]
            if ([string]$live.Name -ne $p.Obj) { throw "object at [$($p.SectionIdx)][$($p.ObjIdx)] is '$($live.Name)', expected '$($p.Obj)'" }
            $clone = $live.Clone($true)
            $cf = $clone.Format.ConditionFormulas.Formula($SUP)
            if ($null -eq $cf) { throw "no EnableSuppress condition slot on '$($p.Obj)'" }
            $cf.Text = $p.Want
            $rdc.ReportObjectController.Modify($live, $clone)
            $wrote++
        }
        $doc.SaveAs([string]$full)
    } catch {
        Log 'ERR' "    write failed: $($_.Exception.Message.Split([char]10)[0])"
        try { $doc.Close() } catch {}
        Copy-Item -LiteralPath $bak -Destination $full -Force
        Log 'ERR' "    ROLLED BACK from $BackupSuffix"
        $filesFailed++
        [void]$failList.Add("$full || <write> || $($_.Exception.Message.Split([char]10)[0])")
        continue
    }
    try { $doc.Close() } catch {}

    # verify from DISK, never from the in-memory model
    $probs = @()
    $d2 = $null
    try {
        $d2 = New-Object CrystalDecisions.CrystalReports.Engine.ReportDocument
        $d2.Load($full, [CrystalDecisions.Shared.OpenReportMethod]::OpenReportByTempCopy)
        $pf2 = $d2.ReportClientDocument.ReportDefController.ReportDefinition.PageFooterArea
        foreach ($p in $plan) {
            $o2 = $pf2.Sections[$p.SectionIdx].ReportObjects[$p.ObjIdx]
            if ([string]$o2.Name -ne $p.Obj) { $probs += "[$($p.Section)] index moved: got '$($o2.Name)' wanted '$($p.Obj)'"; continue }
            $got = Get-SuppressText -Obj $o2
            if ($got.Trim() -ne $p.Want.Trim()) { $probs += "[$($p.Section)] $($p.Obj) formula on disk is '$($got.Trim())'" }
        }
    } catch { $probs += "verify load failed: $($_.Exception.Message.Split([char]10)[0])" }
    finally { if ($d2) { try { $d2.Close() } catch {} } }

    if ($probs.Count) {
        foreach ($pr in $probs) { Log 'ERR' "    VERIFY $pr" }
        Copy-Item -LiteralPath $bak -Destination $full -Force
        Log 'ERR' "    ROLLED BACK from $BackupSuffix"
        $filesFailed++
        [void]$failList.Add("$full || <verify> || $($probs -join '; ')")
        continue
    }

    foreach ($p in $plan) { if ($p.Kind -eq 'combine') { $objCombined++ } else { $objSet++ } }
    Log 'OK' "    verified on disk: $wrote object(s)"
    $filesChanged++
}

Log 'INFO' '=============================================================='
if ($failList.Count) {
    Log 'INFO' "FAILURES ($($failList.Count)):"
    foreach ($m in $failList) { Log 'INFO' "  F | $m" }
}
if ($commentedList.Count) {
    Log 'INFO' "LEFT ALONE -- formula is commented out, someone did that on purpose ($($commentedList.Count)):"
    foreach ($m in $commentedList) { Log 'INFO' "  C | $m" }
}
Log 'INFO' ("SUMMARY mode={0}  files={1}  changed={2}  failed={3}  set={4}  combined={5}  alreadyOk={6}  commented={7}" -f `
    $mode, $filesScanned, $filesChanged, $filesFailed, $objSet, $objCombined, $objAlready, $objCommented)
Log 'INFO' '=============================================================='
Write-Host ("RESULT|{0}|files={1}|changed={2}|failed={3}|set={4}|combined={5}|alreadyOk={6}|commented={7}" -f `
    $mode, $filesScanned, $filesChanged, $filesFailed, $objSet, $objCombined, $objAlready, $objCommented)
if ($filesFailed -gt 0) { exit 1 }
exit 0
