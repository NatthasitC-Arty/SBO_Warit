<#
.SYNOPSIS
  Repairs the SQL Commands inside Crystal Reports .rpt layouts for SAP B1.

.DESCRIPTION
  Three independent transforms, each behind its own switch, each idempotent
  (running twice changes nothing the second time):

    -Approver            T1  approver box shows the FIRST approval round
                             instead of the last. Adds TOP 1 + ORDER BY.
    -ProjectJoin         T2  draft documents print duplicated lines. Replaces
                             the row-multiplying "LEFT JOIN <line> pj" with a
                             correlated OUTER APPLY.
    -DownPaymentObjType  T3  AR Down Payment approver command hard-codes
                             OWDD.ObjType = '23' (Sales Quotation).
    -ProjectName         T4  reports print the project CODE only. Adds
                             OPRJ.PrjName as a 'ProjectName' column.

  MEASURED FACTS this script is built on (2026-09-01, all verified by reading
  the file back from disk, never from memory):

    $table.CommandText = $newSql            -> in-memory ONLY. The saved file
                                               keeps the old SQL. Silent loss.
    clone.CommandText  = $newSql
    Dbc.SetTableLocation($live, $clone)     -> ACTUALLY PERSISTS.

  Database.Tables[$i] hands back a COPY of the table, so mutating the live
  object changes nothing; SetTableLocation is what puts the edit back.

  SetTableLocation has two side effects that MUST be handled:
    1. It blanks the connection UserID unless 'User ID' is written into the
       nested QE_LogonProperties bag on the clone FIRST. This script always
       does that -- a report with a blank UserID is refused by SAP B1.
    2. It injects 'Database DLL' into that same nested bag at save time. That
       key is not an OLE DB keyword, so SAP B1 refuses the report, and NO
       offline edit can remove it (8 routes tried and closed, including
       -Connect + VerifyDatabase and ReplaceConnection, which crashes the
       process). Any table this script rewrites therefore needs ONE pass of
       Database > Set Datasource Location > Update > Ctrl+S in the Designer
       afterwards. The summary reports exactly which files need it.

  ORDER OF WORK: run every SQL fix FIRST, then do the Designer Update pass
  once. Doing it the other way round throws the manual work away, because
  each SQL edit re-injects the key.

.EXAMPLE
  # see what would change, touch nothing
  .\Fix-ReportSql.ps1 -Path "C:\GitHub\SDA\Form-Layout_SDA" -Approver -ProjectJoin -WhatIf

.EXAMPLE
  # apply both fixes to one folder
  .\Fix-ReportSql.ps1 -Path "C:\GitHub\SDA\Form-Layout_SDA\2. Sales - AR" -Approver -ProjectJoin

.NOTES
  64-bit PowerShell only -- the RAS assemblies exist solely in GAC_64.
  Exit codes: 0 all good, 1 at least one file failed/rolled back,
              2 not 64-bit, 3 SDK load failed, 4 bad -Path.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $Path,
    [switch] $Approver,
    [switch] $ProjectJoin,
    [switch] $DownPaymentObjType,
    [switch] $ProjectName,
    [switch] $WhatIf,
    [string] $Filter       = '*.rpt',
    [switch] $NoRecurse,
    [string] $BackupSuffix = '.20260901-sqlfix.bak',
    [string] $LogFile      = "$PSScriptRoot\_FixReportSql.log"
)

$ErrorActionPreference = 'Stop'
$script:BF = [System.Reflection.BindingFlags]

if ([IntPtr]::Size -ne 8) {
    Write-Host '[ERROR] Run this in 64-bit PowerShell -- the RAS assemblies are GAC_64 only.' -ForegroundColor Red
    exit 2
}
if (-not ($Approver -or $ProjectJoin -or $DownPaymentObjType -or $ProjectName)) {
    Write-Host '[ERROR] Pick a transform: -Approver / -ProjectJoin / -DownPaymentObjType / -ProjectName' -ForegroundColor Red
    exit 4
}

# ------------------------------------------------------------------ SDK
function Load-CrystalSDK {
    $gacMsil = 'C:\Windows\Microsoft.NET\assembly\GAC_MSIL'
    $gac64   = 'C:\Windows\Microsoft.NET\assembly\GAC_64'
    $sets = @(
        @{ Root = $gacMsil; Names = @(
            'CrystalDecisions.Shared',
            'CrystalDecisions.ReportSource',
            'CrystalDecisions.CrystalReports.Engine') },
        @{ Root = $gac64;   Names = @(
            'CrystalDecisions.ReportAppServer.CommLayer',
            'CrystalDecisions.ReportAppServer.DataDefModel',
            'CrystalDecisions.ReportAppServer.Controllers',
            'CrystalDecisions.ReportAppServer.ClientDoc',
            'CrystalDecisions.ReportAppServer.CommonObjectModel',
            'CrystalDecisions.ReportAppServer.ObjectFactory') }
    )
    foreach ($set in $sets) {
        foreach ($n in $set.Names) {
            $dir = Join-Path $set.Root $n
            if (-not (Test-Path $dir)) { throw "GAC folder not found: $dir" }
            $dll = Get-ChildItem $dir -Recurse -Filter "$n.dll" |
                   Sort-Object { [version](($_.Directory.Name -split '_')[1]) } -Descending |
                   Select-Object -First 1
            if (-not $dll) { throw "DLL not found under $dir" }
            [void][System.Reflection.Assembly]::LoadFrom($dll.FullName)
        }
    }
}
try { Load-CrystalSDK } catch { Write-Host "[ERROR] SDK load failed: $($_.Exception.Message)" -ForegroundColor Red; exit 3 }

# ------------------------------------------------------------------ logging
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

# ------------------------------------------------------------------ RAS property bag
# The bag is a System.__ComObject; keys are listed by PropertyIDs, read/written
# through Item/Add/Remove/Contains. Assigning an object THROUGH the PowerShell
# indexer poisons the COM binder for the rest of the session -- always use
# InvokeMember or the explicit methods below.
function BagItem {
    param($Bag, [string]$Key)
    if ($null -eq $Bag) { return $null }
    try { return $Bag.Item($Key) } catch { return $null }
}
function BagHas {
    param($Bag, [string]$Key)
    if ($null -eq $Bag) { return $false }
    try { return [bool]$Bag.Contains($Key) } catch { return $false }
}
function BagSetStr {
    param($Bag, [string]$Key, [string]$Value)
    if ($null -eq $Bag) { return $false }
    try { if (BagHas $Bag $Key) { [void]$Bag.Remove($Key) } } catch {}
    try { [void]$Bag.Add($Key, [string]$Value); return $true } catch { return $false }
}

function Read-TableUser {
    param($Table)
    try {
        $ci = $Table.ConnectionInfo
        return [string]$ci.GetType().InvokeMember('UserName', $script:BF::GetProperty, $null, $ci, $null)
    } catch { return '' }
}
function Test-TableDllPoison {
    param($Table)
    try { return (BagHas (BagItem $Table.ConnectionInfo.Attributes 'QE_LogonProperties') 'Database DLL') }
    catch { return $false }
}
function Get-FieldCount {
    param($Table)
    try { return [int]$Table.DataFields.Count } catch { return -1 }
}

# ------------------------------------------------------------------ scopes
# Every (scope, DatabaseController) pair: the main report plus one per
# subreport. GetSubreport() MUST be called on a held reference -- chaining
# $rcd.SubreportController.GetSubreport(..) returns an object whose
# DatabaseController is already dead (null), and a null controller reports
# zero tables without throwing, so subreports get silently skipped.
# Ctl/Sub are carried out unused on purpose: they root the COM references.
function Get-Scopes {
    param($Doc)
    $out = New-Object System.Collections.ArrayList
    $rcd = $Doc.ReportClientDocument
    [void]$out.Add([pscustomobject]@{ Scope = 'main'; Dbc = $rcd.DatabaseController; Ctl = $null; Sub = $null })

    $ctl = $null
    try { $ctl = $rcd.SubreportController }
    catch { Log 'WARN' "    SubreportController unavailable: $($_.Exception.Message.Split([char]10)[0])" }
    if ($null -eq $ctl) { return ,$out }

    $names = @()
    try { foreach ($n in $ctl.GetSubreportNames()) { if ($n) { $names += [string]$n } } }
    catch { Log 'WARN' "    GetSubreportNames failed: $($_.Exception.Message.Split([char]10)[0])" }

    foreach ($sn in $names) {
        try {
            $sub = $ctl.GetSubreport($sn)
            if ($null -eq $sub) { Log 'ERR' "    GetSubreport('$sn') returned null -- SUBREPORT NOT PROCESSED"; continue }
            $sdbc = $sub.DatabaseController
            if ($null -eq $sdbc) { Log 'ERR' "    subreport '$sn' has a null DatabaseController -- NOT PROCESSED"; continue }
            [void]$out.Add([pscustomobject]@{ Scope = "sub:$sn"; Dbc = $sdbc; Ctl = $ctl; Sub = $sub })
        } catch { Log 'ERR' "    subreport '$sn' failed: $($_.Exception.Message.Split([char]10)[0])" }
    }
    return ,$out
}

# ================================================================== T1
# The approver Command joins OWDD -> WDD1 filtered to Status='Y', which is one
# row PER COMPLETED APPROVAL ROUND. With no ORDER BY, SQL Server may return
# them in any order and the report binds whichever row comes first -- usually
# the earliest. TOP 1 + newest-first ordering pins it to the latest round.
#
# The select list is never touched. Column aliases differ between sibling
# reports on purpose ("ApproveDate" in Sale Order (Dis) vs "ApprovalDate" in
# (Bom), and a bare `As Path`); renaming one to match the other would change
# the Crystal field name and orphan the object bound to it on the layout.
function Convert-ApproverSql {
    param([string]$Sql)
    if ($Sql -notmatch '(?i)\bOWDD\b' -or $Sql -notmatch '(?i)\bWDD1\b') { return $null }
    if ($Sql -match '(?i)\bORDER\s+BY\b') { return $null }   # already ordered

    # ---- shape B: outer SELECT T0.* over a UNION ALL of a posted and a draft branch
    if ($Sql -match '(?i)SELECT\s+T0\.\*') {
        if ($Sql -match '(?i)SELECT\s+TOP\s') { return $null }
        $new = $Sql
        # ApprovalDate holds a DATE only, so two rounds on the same day cannot be
        # separated by it. Carry UpdateTime out of BOTH branches -- adding a column
        # is safe (existing field names are untouched); renaming one would not be.
        $rxDate = New-Object System.Text.RegularExpressions.Regex `
            "(?i)(WDD1\.UpdateDate\s+AS\s+'ApprovalDate')"
        if ($rxDate.Matches($new).Count -ne 2) { return $null }   # unexpected shape, leave alone
        $new = $rxDate.Replace($new, "`$1,`r`n    WDD1.UpdateTime        AS 'ApprovalTime'")
        $new = [System.Text.RegularExpressions.Regex]::Replace($new, '(?i)SELECT\s+T0\.\*', 'SELECT TOP 1 T0.*', 1)
        $new = $new.TrimEnd() + "`r`nORDER BY T0.ApprovalDate DESC, T0.ApprovalTime DESC"
        return $new
    }

    # ---- shapes A and C: one flat SELECT
    if ($Sql -match '(?i)^\s*SELECT\s+TOP\s') { return $null }
    $rx = New-Object System.Text.RegularExpressions.Regex '(?i)^(\s*)SELECT(\s)'
    if (-not $rx.IsMatch($Sql)) { return $null }
    $new = $rx.Replace($Sql, '${1}SELECT TOP 1${2}', 1)

    # Follow the quoting style already used by this command instead of forcing
    # one house style across the repo.
    if ($Sql -match '(?i)WDD1\."UpdateDate"') {
        $new = $new.TrimEnd() + "`r`nORDER BY`r`n    WDD1.`"UpdateDate`" DESC,`r`n    WDD1.`"UpdateTime`" DESC"
    } else {
        $new = $new.TrimEnd() + "`r`nORDER BY`r`n    WDD1.UpdateDate DESC,`r`n    WDD1.UpdateTime DESC"
    }
    return $new
}

# ================================================================== T2
# "LEFT JOIN <line> pj ON <hdr>.<key> = <x>.<key> AND pj.Project IS NOT NULL"
# exists only to surface the document's Project code, but it carries no
# LineNum predicate, so every detail row is multiplied by the number of lines
# carrying a Project. 23 of the 27 affected files also have a copy-paste typo
# in the ON clause -- it compares the FIRST line alias instead of pj, so
# pj.DocEntry is never constrained at all. In the draft branch that fans out
# across DRF1, which holds the drafts of every document type in the company;
# that is why the symptom shows up on drafts.
#
# The replacement is the pattern the Sale Order reports already use.
function Convert-ProjectJoin {
    param([string]$Sql)
    if ($Sql -notmatch '(?i)\bpj\.') { return $null }

    $rx = New-Object System.Text.RegularExpressions.Regex `
        "(?im)^([ \t]*)LEFT\s+JOIN\s+(\[?[A-Za-z0-9_@]+\]?)\s+pj\s+ON\s+([A-Za-z0-9_]+)\.\[?([A-Za-z0-9_]+)\]?\s*=\s*[A-Za-z0-9_]+\.\[?[A-Za-z0-9_]+\]?\s+AND\s+pj\.Project\s+IS\s+NOT\s+NULL\s+AND\s+pj\.Project\s*<>\s*''[ \t]*\r?$"
    if (-not $rx.IsMatch($Sql)) { return $null }

    $evaluator = {
        param($m)
        $indent = $m.Groups[1].Value
        $line   = $m.Groups[2].Value
        $hdr    = $m.Groups[3].Value
        $key    = $m.Groups[4].Value
        @(
            "$indent-- OUTER APPLY keeps Project a header-level lookup; the old pj join",
            "$indent-- multiplied every detail row. Do not turn it back into a JOIN.",
            "$indent" + "OUTER APPLY (",
            "$indent    SELECT TOP 1 P.Project",
            "$indent    FROM $line P",
            "$indent    WHERE P.$key = $hdr.$key",
            "$indent      AND P.Project IS NOT NULL",
            "$indent      AND P.Project <> ''",
            "$indent) QPJ"
        ) -join "`r`n"
    }
    $new = $rx.Replace($Sql, $evaluator)

    # Everything else that still says pj.<col> now has to point at the APPLY.
    $new = [System.Text.RegularExpressions.Regex]::Replace($new, '(?i)\bpj\.', 'QPJ.')

    # Nothing may be left over. A stray pj alias means the JOIN regex missed a
    # variant and the SQL would no longer compile -- refuse rather than write it.
    if ($new -match '(?i)(?<!Q)\bpj\.') { throw 'T2: a "pj." reference survived the rewrite' }
    return $new
}

# ================================================================== T3
# The AR Down Payment approver command filters OWDD.ObjType = '23'
# (Sales Quotation) on both UNION branches although the document is an
# AR Down Payment. Every other report in the repo joins OWDD on the
# document's own ObjType, so '203' is what this should be.
#
# UNVERIFIED AGAINST LIVE DATA -- this is inferred from the surrounding code,
# not measured, which is why it sits behind its own switch.
function Convert-DownPaymentObjType {
    param([string]$Sql)
    if ($Sql -notmatch '(?i)\bODPI\b') { return $null }
    $rx = New-Object System.Text.RegularExpressions.Regex "(?i)(OWDD\.ObjType\s*=\s*)'23'"
    if (-not $rx.IsMatch($Sql)) { return $null }
    return $rx.Replace($Sql, "`$1'203'")
}

# ================================================================== T4
# Every report prints the Project CODE. OPRJ is joined in most branches but
# PrjName is never selected anywhere in the repo (0 hits across 64 files), so
# the project name has never reached a layout.
#
# WHY A SUBQUERY AND NOT THE OBVIOUS "OPRJ.PrjName AS 'ProjectName'":
#   - 4 of the 8 branches of AP Down Payment select Project with NO OPRJ join
#     in that branch, so the plain column reference would not compile there.
#   - Every branch of AP Down Payment / AP Credit Memo / Purchase Request ends
#     in a GROUP BY. Adding a plain column to the SELECT list of a grouped
#     query also requires adding it to the GROUP BY -- a second edit, in a
#     second place, that has to stay in step.
#   - Sale Quotation does not select a Project column at all; it derives the
#     most-frequent project with a correlated subquery.
# A correlated scalar subquery sidesteps all three: it needs no join, it is
# legal under GROUP BY as long as it only references grouped columns (it
# references the Project column, which is always grouped), and it can wrap the
# Sale Quotation expression as-is. One code path, no row multiplication.
#
# COLUMN COUNT IS THE HARD CONSTRAINT: UNION ALL branches must stay identical
# in count and order, so either EVERY branch gains exactly one column at the
# same position or the command is left alone.
function Convert-ProjectName {
    param([string]$Sql)
    if ($Sql -match '(?i)PrjName')      { return $null }   # already done
    if ($Sql -notmatch "(?i)\.Project\b|AS\s+'Project'") { return $null }

    # HARD STOP ON COMMON TABLE EXPRESSIONS. Learned the expensive way on
    # 2026-09-01: Sale Quotation (BOM) and the six AR Down Payment reports are
    # built as "WITH ComponentSums AS ( ... UNION ALL ... ) SELECT ... OUTER
    # APPLY (...) QPJ". The UNION ALL inside the CTE looks exactly like a
    # top-level branch, so the new column landed INSIDE the CTE while the QPJ
    # alias it referenced lives in the final SELECT, two scopes away:
    #   ADO 0x80040e14 / SQL state 42S22
    #   The multi-part identifier "QPJ.Project" could not be bound
    # A naive "is the alias defined in this segment" check does NOT catch it,
    # because the CTE body and the final SELECT sit in the same segment.
    # Handling a CTE properly means finding where the WITH list closes and
    # working only on the final SELECT. Until that exists, refuse the command
    # loudly rather than write SQL that cannot run.
    if ($Sql -match '(?im)^[ \t]*(WITH|,)[ \t]*[A-Za-z0-9_]+[ \t]+AS[ \t]*\(') {
        throw 'T4: command is built on a WITH/CTE -- UNION ALL inside a CTE is not a top-level branch, ProjectName is not safe to add here'
    }

    # Parenthesis depth, computed ONCE over the whole command, is the only
    # reliable way to find a branch's own SELECT list. Line shape is not:
    # "FROM QUT1 T0" sits alone on its line inside a scalar subquery in Sale
    # Quotation, so a line-based scan stops three branches early. Quoted
    # literals are skipped so a paren inside a string cannot skew the count.
    $depth = New-Object 'int[]' $Sql.Length
    $lvl = 0; $q = $false
    for ($i = 0; $i -lt $Sql.Length; $i++) {
        $c = $Sql[$i]
        if ($c -eq "'") { $q = -not $q }
        elseif (-not $q) {
            if ($c -eq '(') { $lvl++ } elseif ($c -eq ')') { $lvl-- }
        }
        $depth[$i] = $lvl
    }

    $selAll  = [regex]::Matches($Sql, '(?i)\bSELECT\b')
    $fromAll = [regex]::Matches($Sql, '(?i)\bFROM\b')

    # segment boundaries: the UNION ALL lines that separate the branches
    $bounds = @(0)
    foreach ($m in [regex]::Matches($Sql, '(?im)^[ \t]*UNION[ \t]+ALL[ \t]*\r?$')) { $bounds += $m.Index + $m.Length }
    $bounds += $Sql.Length

    $edits = New-Object System.Collections.ArrayList
    for ($b = 0; $b -lt $bounds.Count - 1; $b++) {
        $segS = $bounds[$b]; $segE = $bounds[$b + 1]

        # walk past any wrapper: "SELECT T0.* FROM ( <the real branch> ) T0"
        $pos = $segS; $selIdx = -1; $fromIdx = -1; $fromStart = -1; $L = 0
        while ($true) {
            $selIdx = -1
            foreach ($m in $selAll) { if ($m.Index -ge $pos -and $m.Index -lt $segE) { $selIdx = $m.Index; break } }
            if ($selIdx -lt 0) { break }
            $L = $depth[$selIdx]
            $fromIdx = -1
            foreach ($m in $fromAll) {
                if ($m.Index -gt $selIdx -and $m.Index -lt $segE -and $depth[$m.Index] -eq $L) { $fromStart = $m.Index; $fromIdx = $m.Index + $m.Length; break }
            }
            if ($fromIdx -lt 0) { break }
            $rest = $Sql.Substring($fromIdx, [Math]::Min(40, $segE - $fromIdx))
            if ($rest -match '^\s*\(') { $pos = $fromIdx + $rest.IndexOf('(') + 1; continue }   # wrapper, go inside
            break
        }
        if ($selIdx -lt 0 -or $fromIdx -lt 0) { continue }   # segment carries no branch

        # The new column always goes LAST in the branch, never next to the
        # Project column. Position has to be identical in every branch of a
        # UNION and some branches carry no Project column at all (Production
        # joins OPRJ on the header and never selects it; the free-text branches
        # of AR Down Payment select nothing). "Last" is the one position every
        # branch can agree on.
        $head = $Sql.Substring($selIdx, $fromStart - $selIdx)
        $anchor = -1; $anchorEnd = -1
        foreach ($m in [regex]::Matches($head, "(?i)AS[ \t]+'Project'")) {
            if ($depth[$selIdx + $m.Index] -eq $L) { $anchor = $selIdx + $m.Index; $anchorEnd = $anchor + $m.Length }
        }
        if ($anchor -lt 0) {
            foreach ($m in [regex]::Matches($head, '(?i)\b[A-Za-z0-9_]+\.Project\b')) {
                if ($depth[$selIdx + $m.Index] -eq $L) { $anchor = $selIdx + $m.Index; $anchorEnd = $anchor + $m.Length }
            }
        }

        $expr = ''
        if ($anchor -ge 0) {
            $itemStart = $selIdx + 6
            for ($i = $anchor - 1; $i -gt $selIdx; $i--) {
                if ($Sql[$i] -eq ',' -and $depth[$i] -eq $L) { $itemStart = $i + 1; break }
            }
            $itemEnd = $fromStart
            for ($i = $anchorEnd; $i -lt $fromStart; $i++) {
                if ($Sql[$i] -eq ',' -and $depth[$i] -eq $L) { $itemEnd = $i; break }
            }
            $expr = [regex]::Replace($Sql.Substring($itemStart, $itemEnd - $itemStart), "(?i)\s+AS\s+'Project'\s*$", '').Trim()
        } else {
            # no Project column in this branch: fall back to the code the branch
            # already joins OPRJ on, e.g. "LEFT JOIN OPRJ ON OWOR.Project = OPRJ.PrjCode"
            $tail = $Sql.Substring($fromStart, $segE - $fromStart)
            $jm = [regex]::Match($tail, '(?i)JOIN\s+OPRJ\s+ON\s+([A-Za-z0-9_]+\.[A-Za-z0-9_]*Project[A-Za-z0-9_]*)\s*=')
            if ($jm.Success) { $expr = $jm.Groups[1].Value }
        }
        if (-not $expr) { $expr = 'NULL' }

        if ($expr -match '(?i)^NULL$') {
            $col = "NULL AS 'ProjectName'"
        } else {
            # a scalar subquery needs no join of its own and stays legal under the
            # GROUP BY that most of these branches end with
            $col = "(SELECT TOP 1 PJN.PrjName FROM OPRJ PJN WHERE PJN.PrjCode = ($expr)) AS 'ProjectName'"
        }
        [void]$edits.Add([pscustomobject]@{ At = $fromStart; Text = ",`r`n    $col`r`n" })
    }
    if ($edits.Count -eq 0) { return $null }
    $out = $Sql
    foreach ($e in ($edits | Sort-Object At -Descending)) {
        $out = $out.Substring(0, $e.At) + $e.Text + $out.Substring($e.At)
    }
    return $out
}
# ------------------------------------------------------------------ enumerate
$rpts = @()
if (Test-Path -LiteralPath $Path -PathType Leaf) {
    if ($Path -like '*.rpt') { $rpts = ,(Get-Item -LiteralPath $Path) }
    else { Log 'ERR' "Path is a file but not .rpt: $Path"; exit 4 }
} elseif (Test-Path -LiteralPath $Path -PathType Container) {
    $args2 = @{ LiteralPath = $Path; Filter = $Filter; File = $true }
    if (-not $NoRecurse) { $args2['Recurse'] = $true }
    $rpts = Get-ChildItem @args2
} else {
    Log 'ERR' "Path not found: $Path"; exit 4
}

$mode = if ($WhatIf) { 'WHATIF' } else { 'APPLY' }
$on = @()
if ($Approver)           { $on += 'T1-Approver' }
if ($ProjectJoin)        { $on += 'T2-ProjectJoin' }
if ($DownPaymentObjType) { $on += 'T3-DownPaymentObjType' }
if ($ProjectName)        { $on += 'T4-ProjectName' }

Log 'INFO' "=== Fix-ReportSql start  mode=$mode  transforms=$($on -join ',') ==="
Log 'INFO' "Path=$Path  files=$($rpts.Count)"

# ------------------------------------------------------------------ counters
$filesScanned = 0; $filesChanged = 0; $filesFailed = 0; $filesSkipped = 0
$t1Count = 0; $t2Count = 0; $t3Count = 0; $t4Count = 0
$cmdChanged = 0; $newlyPoisoned = 0
$poisonFiles = New-Object System.Collections.ArrayList
$failList    = New-Object System.Collections.ArrayList

foreach ($f in $rpts) {
    $full = $f.FullName
    $filesScanned++
    $doc = $null
    $plan = New-Object System.Collections.ArrayList   # what to write in this file

    try {
        $doc = New-Object CrystalDecisions.CrystalReports.Engine.ReportDocument
        $doc.Load($full, [CrystalDecisions.Shared.OpenReportMethod]::OpenReportByTempCopy)

        foreach ($sc in (Get-Scopes -Doc $doc)) {
            $cnt = 0
            try { $cnt = [int]$sc.Dbc.Database.Tables.Count } catch { $cnt = 0 }
            for ($i = 0; $i -lt $cnt; $i++) {
                $t = $sc.Dbc.Database.Tables[$i]
                $alias = [string]$t.Alias
                $sql = ''
                try { $sql = [string]$t.CommandText } catch {}
                if (-not $sql) { continue }

                $cur = $sql
                $tags = @()
                try {
                    if ($Approver)           { $r = Convert-ApproverSql        -Sql $cur; if ($r) { $cur = $r; $tags += 'T1'; $t1Count++ } }
                    if ($ProjectJoin)        { $r = Convert-ProjectJoin        -Sql $cur; if ($r) { $cur = $r; $tags += 'T2'; $t2Count++ } }
                    if ($DownPaymentObjType) { $r = Convert-DownPaymentObjType -Sql $cur; if ($r) { $cur = $r; $tags += 'T3'; $t3Count++ } }
                    if ($ProjectName)        { $r = Convert-ProjectName        -Sql $cur; if ($r) { $cur = $r; $tags += 'T4'; $t4Count++ } }
                } catch {
                    Log 'ERR' "    [$($sc.Scope)] $alias : $($_.Exception.Message) -- file left untouched"
                    $tags = @()
                    $cur  = $sql
                    [void]$failList.Add("$full || $($sc.Scope)|$alias || $($_.Exception.Message)")
                }
                if ($cur -eq $sql) { continue }

                [void]$plan.Add([pscustomobject]@{
                    Scope = $sc.Scope; Index = $i; Alias = $alias
                    OldLen = $sql.Length; NewSql = $cur; Tags = ($tags -join '+')
                    User = (Read-TableUser -Table $t)
                    Fields = (Get-FieldCount -Table $t)
                    WasPoisoned = (Test-TableDllPoison -Table $t)
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

    if ($plan.Count -eq 0) {
        $filesSkipped++
        if ($doc) { try { $doc.Close() } catch {} }
        continue
    }

    Log 'HIT' "--> $full"
    foreach ($p in $plan) {
        Log 'INFO' ("    [{0}] {1} : {2}  {3} -> {4} chars  User='{5}'" -f $p.Scope, $p.Alias, $p.Tags, $p.OldLen, $p.NewSql.Length, $p.User)
    }
    if ($WhatIf) {
        $cmdChanged += $plan.Count
        if ($doc) { try { $doc.Close() } catch {} }
        continue
    }

    # ---- write
    $bak = "$full$BackupSuffix"
    Copy-Item -LiteralPath $full -Destination $bak -Force
    $wrote = 0
    try {
        foreach ($sc in (Get-Scopes -Doc $doc)) {
            foreach ($p in ($plan | Where-Object { $_.Scope -eq $sc.Scope })) {
                $live  = $sc.Dbc.Database.Tables[$p.Index]
                $clone = $live.Clone($true)
                $clone.CommandText = $p.NewSql
                # UserID lives in the nested logon bag; ConnectionInfo.UserName is a
                # floating value that never reaches the file. Write it BEFORE
                # SetTableLocation or the table comes back with a blank login.
                if ($p.User) {
                    $lp = BagItem $clone.ConnectionInfo.Attributes 'QE_LogonProperties'
                    if ($lp) { [void](BagSetStr $lp 'User ID' $p.User) }
                    else { Log 'WARN' "    [$($p.Scope)] $($p.Alias) : no QE_LogonProperties bag -- UserID cannot be carried" }
                }
                $sc.Dbc.SetTableLocation($live, $clone)
                $wrote++
            }
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

    # ---- verify from DISK. An in-memory read-back is exactly how this class of
    #      bug stayed invisible for months.
    $probs = @()
    $poisonedHere = 0
    $d2 = $null
    try {
        $d2 = New-Object CrystalDecisions.CrystalReports.Engine.ReportDocument
        $d2.Load($full, [CrystalDecisions.Shared.OpenReportMethod]::OpenReportByTempCopy)
        foreach ($sc in (Get-Scopes -Doc $d2)) {
            foreach ($p in ($plan | Where-Object { $_.Scope -eq $sc.Scope })) {
                $t2o = $sc.Dbc.Database.Tables[$p.Index]
                if ([string]$t2o.Alias -ne $p.Alias) { $probs += "[$($p.Scope)] index $($p.Index) is now '$($t2o.Alias)' not '$($p.Alias)'"; continue }
                $got = ''
                try { $got = [string]$t2o.CommandText } catch {}
                if ($got -ne $p.NewSql) { $probs += "[$($p.Scope)] $($p.Alias) SQL not written (disk has $($got.Length) chars, wanted $($p.NewSql.Length))" }
                $gotUser = Read-TableUser -Table $t2o
                if ($p.User -and $gotUser -ne $p.User) { $probs += "[$($p.Scope)] $($p.Alias) UserID lost ('$($p.User)' -> '$gotUser')" }
                $gotFields = Get-FieldCount -Table $t2o
                if ($p.Fields -ge 0 -and $gotFields -ge 0 -and $gotFields -ne $p.Fields) {
                    $probs += "[$($p.Scope)] $($p.Alias) field count $($p.Fields) -> $gotFields"
                }
                if ((Test-TableDllPoison -Table $t2o) -and -not $p.WasPoisoned) { $poisonedHere++ }
            }
        }
    } catch {
        $probs += "verify load failed: $($_.Exception.Message.Split([char]10)[0])"
    } finally { if ($d2) { try { $d2.Close() } catch {} } }

    if ($probs.Count) {
        foreach ($pr in $probs) { Log 'ERR' "    VERIFY $pr" }
        Copy-Item -LiteralPath $bak -Destination $full -Force
        Log 'ERR' "    ROLLED BACK from $BackupSuffix"
        $filesFailed++
        [void]$failList.Add("$full || <verify> || $($probs -join '; ')")
        continue
    }

    Log 'OK' "    verified on disk: $wrote command(s) rewritten"
    $filesChanged++
    $cmdChanged += $wrote
    if ($poisonedHere -gt 0) {
        $newlyPoisoned += $poisonedHere
        [void]$poisonFiles.Add($full)
        Log 'WARN' "    $poisonedHere table(s) now carry 'Database DLL' -- this file needs one Designer Set Datasource Location > Update"
    }
}

Log 'INFO' '=============================================================='
if ($failList.Count) {
    Log 'INFO' "FAILURES ($($failList.Count)):"
    foreach ($m in $failList) { Log 'INFO' "  F | $m" }
}
if ($poisonFiles.Count) {
    Log 'INFO' "NEEDS DESIGNER UPDATE ($($poisonFiles.Count)):"
    foreach ($m in $poisonFiles) { Log 'INFO' "  U | $m" }
}
Log 'INFO' ("SUMMARY mode={0}  files={1}  changed={2}  skipped={3}  failed={4}  commands={5}  T1={6}  T2={7}  T3={8}  T4={9}  newlyPoisoned={10}" -f `
    $mode, $filesScanned, $filesChanged, $filesSkipped, $filesFailed, $cmdChanged, $t1Count, $t2Count, $t3Count, $t4Count, $newlyPoisoned)
Log 'INFO' '=============================================================='
Write-Host ("RESULT|{0}|files={1}|changed={2}|skipped={3}|failed={4}|commands={5}|T1={6}|T2={7}|T3={8}|T4={9}|newlyPoisoned={10}" -f `
    $mode, $filesScanned, $filesChanged, $filesSkipped, $filesFailed, $cmdChanged, $t1Count, $t2Count, $t3Count, $t4Count, $newlyPoisoned)

if ($filesFailed -gt 0) { exit 1 }
exit 0
