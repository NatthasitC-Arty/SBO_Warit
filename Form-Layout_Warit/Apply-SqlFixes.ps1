<#
.SYNOPSIS
  Rewrites the SQL of Command tables inside Crystal Report (.rpt) files, in one
  pass, with a verified save.

.DESCRIPTION
  Opens every .rpt under -Root ONCE, applies an ordered list of regex transforms
  to each Command table's SQL, saves the report, then RE-OPENS IT FROM DISK and
  checks table count / field count / connection / exact SQL. Any report that
  fails verification is restored from a pre-save temp copy, so a bad write never
  survives.

  One pass matters: running one transform at a time re-opened all 64 reports per
  transform (1024 loads, ~50 min). One pass is ~70 seconds.

  HOW THE SQL IS WRITTEN -- measured on this repo 2026-08-28, do not change it:
      $table.CommandText = $sql           -> reading it back IN MEMORY shows the
                                             new length, the SAVED FILE keeps the
                                             OLD SQL. Silent data loss.
      InvokeMember('CommandText',...)     -> same silent failure.
      clone = table.Clone($true)
      clone.CommandText = $sql
      Dbc.SetTableLocation($live, $clone) -> ACTUALLY PERSISTS.
  Database.Tables[$i] hands back a COPY of the table, so mutating it changes
  nothing; SetTableLocation is what puts the modified table back in the report.

  KEEPING THE CONNECTION'S User ID -- measured 2026-08-31 on C:\_bakchk\B25.rpt.
  A .rpt whose UserID reads as '' CANNOT BE OPENED BY SAP B1.
      SetTableLocation(live, clone)   -> SQL persists but the UserID is WIPED at
                                         save time. It happens even when the
                                         clone is UNMODIFIED, so SetTableLocation
                                         itself is the poison, not the new SQL.
      clone.ConnectionInfo.UserName   -> in memory only; the saved file has ''.
      ApplyLogOnInfo / ModifyTableConnectionInfo / DataSourceConnections.SetLogon
      / SetDatabaseLogon / SetTableLocationByServerDatabaseName / ReplaceConnection
                                      -> none can put a UserID back once it is ''.
      clone.ConnectionInfo.Attributes['QE_LogonProperties'].Add('User ID',$u)
        + Attributes.Add('QE_LogonProperties',$lp) + clone.ConnectionInfo = $ci
        + SetTableLocation(live, clone)
                                      -> SQL PERSISTS **AND** UserID SURVIVES.
  Set-CommandSql below is the only sanctioned write path. -UserID additionally
  stamps a User ID onto reports that have none.

.PARAMETER Root
  Folder to scan recursively, or a single .rpt.

.PARAMETER Only
  Transform id prefixes to run, e.g. -Only 'T4,T5'. Blank = run them all.
  A comma string is split here on purpose: powershell.exe -File does NOT turn
  "T4,T5" into an array, it passes one string, and every -like test then fails
  silently.

.PARAMETER List
  Print the built-in transform table and exit. Nothing is opened.

.PARAMETER Pattern / .PARAMETER Replacement / .PARAMETER Alias / .PARAMETER FileRx
  Ad-hoc mode. Supplying -Pattern replaces the whole built-in table with a single
  transform, so this script can do one-off jobs without being edited:
     -Pattern 'OUGP\.UgpCode' -Replacement '...' -Alias AR_SQ
  -Alias limits by table alias, -FileRx is a regex on the report FILE NAME (an
  alias is NOT unique across reports -- G_Receipt and G_Issue_BS are used by both
  the Thai and the English form).

.PARAMETER DryRun
  Count and log what would change. Nothing is written.

.EXAMPLE
  # what is still outstanding on this repo:
  Apply-SqlFixes.ps1 -Root <repo> -Only T3 -DryRun          # the pj cartesian joins
  Apply-SqlFixes.ps1 -Root <repo> -Alias AR_SQ -Pattern ... # the 13th UoM report
#>
param(
    [Parameter(Mandatory)] [string] $Root,
    [string[]] $Only        = @(),
    [switch]   $List,
    [string]   $Pattern     = '',
    [string]   $Replacement = '',
    [string[]] $Alias       = @(),
    [string]   $FileRx      = '',
    [switch]   $DryRun,
    [string]   $UserID      = '',
    [string]   $LogPath     = ''
)
$ErrorActionPreference = 'Stop'

# powershell.exe -File does NOT split a comma list into an array: "-Only T4,T5"
# arrives as the single string "T4,T5", every -like test then fails, and the
# script silently does nothing at all. Split it here.
$Only = @($Only | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $PSCommandPath }
if (-not $LogPath) { $LogPath = Join-Path $scriptDir '_ApplySqlFixes.log' }

if (-not [Environment]::Is64BitProcess) { Write-Host '[ERROR] need 64-bit powershell.exe'; exit 2 }
foreach ($s in @(
 @{R='C:\Windows\Microsoft.NET\assembly\GAC_MSIL';N=@('CrystalDecisions.Shared','CrystalDecisions.ReportSource','CrystalDecisions.CrystalReports.Engine')},
 @{R='C:\Windows\Microsoft.NET\assembly\GAC_64';N=@('CrystalDecisions.ReportAppServer.CommLayer','CrystalDecisions.ReportAppServer.DataDefModel','CrystalDecisions.ReportAppServer.Controllers','CrystalDecisions.ReportAppServer.ClientDoc','CrystalDecisions.ReportAppServer.CommonObjectModel','CrystalDecisions.ReportAppServer.ObjectFactory')})) {
  foreach ($n in $s.N) {
    $d = Join-Path $s.R $n
    $dll = Get-ChildItem $d -Recurse -Filter "$n.dll" | Sort-Object { [version](($_.Directory.Name -split '_')[1]) } -Descending | Select-Object -First 1
    [void][System.Reflection.Assembly]::LoadFrom($dll.FullName)
  }
}

function Log {
    param([string]$Lvl,[string]$Msg)
    $line = "[{0}][{1}] {2}" -f (Get-Date).ToString('HH:mm:ss'), $Lvl, $Msg
    switch ($Lvl) {
        'ERR' { Write-Host $line -ForegroundColor Red }
        'OK'  { Write-Host $line -ForegroundColor Green }
        'HIT' { Write-Host $line -ForegroundColor Cyan }
        default { Write-Host $line }
    }
    try { [IO.File]::AppendAllText($LogPath, $line + "`r`n", [Text.Encoding]::UTF8) } catch {}
}

# RAS objects are System.__ComObject; the string indexer is the only reader that
# works and assigning an OBJECT through it poisons every later string write.
$BF = [System.Reflection.BindingFlags]
function BagGet { param($b,[string]$k)
    if ($null -eq $b) { return '' }
    try { $v = $b.GetType().InvokeMember('Item',$BF::GetProperty,$null,$b,@([object]$k)); if ($null -eq $v) { return '' }; return [string]$v } catch { return '' } }

# Writes NewSql into the command table at $Index and puts it back with
# SetTableLocation, carrying the connection's User ID across through the nested
# QE_LogonProperties bag so the save cannot blank it. See header.
# Returns: keepuser | forced | fallback | nouser
function Set-CommandSql {
    param($Dbc, [int]$Index, [string]$NewSql, [string]$Force = '')
    $live = $Dbc.Database.Tables[$Index]
    $usr  = ''
    try { $usr = [string]$live.ConnectionInfo.UserName } catch {}
    $route = 'nouser'
    if ($Force) { $route = 'forced'; $usr = $Force } elseif ($usr) { $route = 'keepuser' }

    $clone = $live.Clone($true)
    $clone.CommandText = $NewSql

    if ($usr) {
        try {
            $ci = $clone.ConnectionInfo
            # Reading the bag through the string indexer is safe; only ASSIGNING
            # an object through it poisons later writes. Put it back with Add().
            $lp = $ci.Attributes['QE_LogonProperties']
            if ($null -eq $lp) { throw 'connection has no QE_LogonProperties bag' }
            $lp.Add('User ID', $usr)
            $ci.Attributes.Add('QE_LogonProperties', $lp)
            $clone.ConnectionInfo = $ci
        } catch {
            Log 'WARN' "        cannot carry User ID '$usr' across: $($_.Exception.Message)"
            $route = 'fallback'
        }
    }
    $Dbc.SetTableLocation($live, $clone)
    return $route
}

# ------------------------------------------------------------------ transforms
# Order is the order Ars asked for: NVARCHAR -> UoM -> pj -> repairs.
# Alias = $null means "every command table".
$UOM = "ISNULL(NULLIF(OUOM.U_SLD_Uomforeign,''), {0}.UomCode)"
$NL  = "`r`n"
$T = New-Object System.Collections.ArrayList
# FileRx: optional regex on the report FILE NAME. Needed because an alias is not
# unique across reports -- G_Receipt / G_Issue_BS are used by BOTH the Thai and
# the English form, and only the English ones are in scope.
function Tr { param([string]$Id,$Alias,[string]$Pat,[string]$Rep,[string]$FileRx='')
    [void]$script:T.Add([pscustomobject]@{ Id=$Id; Alias=$Alias; Pat=$Pat; Rep=$Rep; FileRx=$FileRx; Hits=0; Tables=0 }) }

# --- NVARCHAR(MAX) breaks Crystal's Verify Database (field remaps to Memo) ---
Tr 'T4-nvarchar' $null 'NVARCHAR\s*\(\s*MAX\s*\)' 'NVARCHAR(4000)'

# --- UoM: OUGP.UgpCode -> the foreign UoM UDF on OUOM ---
$uomAliases = @('AP_DownPayment','AR_DownPayment','AR_INV','AR_CN','AR_Delivery_BS','AP_PO','AR_Retrun_BS')
Tr 'T5a-uom-join' $uomAliases 'LEFT\s+JOIN\s+OUGP\s+ON\s+(\w+)\.UomCode\s*=\s*OUGP\.UgpCode' 'LEFT JOIN OUOM ON ${1}.UomCode = OUOM.UomCode'
# AP_DownPayment ENG also lists OUGP.UgpCode in GROUP BY. The SELECT copy has the
# comma on the NEXT line, the GROUP BY copy has it on the same line -- so match
# the SELECT one first with a lookahead. GROUP BY must not get "AS UgpCode";
# "GROUP BY <expr> AS x" is a syntax error.
Tr 'T5b-uom-sel-DPO1'   'AP_DownPayment' 'OUGP\.UgpCode(?=[ \t]*\r?\n[ \t]*,)' (($UOM -f 'DPO1') + ' AS UgpCode')
Tr 'T5c-uom-group-DPO1' 'AP_DownPayment' 'OUGP\.UgpCode'                        ($UOM -f 'DPO1')
foreach ($p in @(@('AR_DownPayment','DPI1'),@('AR_INV','INV1'),@('AR_CN','RIN1'),
                 @('AR_Delivery_BS','DLN1'),@('AP_PO','POR1'),@('AR_Retrun_BS','RDN1'))) {
    Tr ("T5d-uom-sel-" + $p[1]) $p[0] 'OUGP\.UgpCode' (($UOM -f $p[1]) + ' AS UgpCode')
}
# --- two reports never joined a UoM table at all: add the join, alias the select
Tr 'T6-gr-select' 'G_Receipt'  'IGN1\.UomCode,' (($UOM -f 'IGN1') + " AS 'UomCode',") 'ENG'
Tr 'T6-gr-join'   'G_Receipt'  '(LEFT JOIN (?:DRF1 )?IGN1 ON OIGN\.DocEntry = IGN1\.DocEntry[ \t]*\r?\n)' ('${1}LEFT JOIN OUOM ON IGN1.UomCode = OUOM.UomCode' + $NL) 'ENG'
Tr 'T6-gi-select' 'G_Issue_BS' 'IGE1\.UomCode,' (($UOM -f 'IGE1') + " AS 'UomCode',") 'ENG'
Tr 'T6-gi-join'   'G_Issue_BS' '(INNER JOIN (?:IGE1|DRF1) IGE1 ON OIGE\.DocEntry = IGE1\.DocEntry[ \t]*\r?\n)' ('${1}LEFT JOIN OUOM ON IGE1.UomCode = OUOM.UomCode' + $NL) 'ENG'

# --- the draft-line doubling: "LEFT JOIN X pj ON <hdr>.DocEntry = <line>.DocEntry"
# never names pj, so pj is unconstrained and multiplies rows. Idempotent.
Tr 'T3-pj' $null '(LEFT\s+JOIN\s+\w+\s+pj\s+ON\s+\[?\w+\]?\.\[?DocEntry\]?\s*=\s*)\[?\w+\]?(\.\[?DocEntry\]?)' '${1}pj${2}'

# --- text branches INNER JOIN <T>10 then LEFT JOIN it again: "correlation name
# specified multiple times" -- these three commands never executed at all.
Tr 'T1-dup-alias' @('AR_Retrun_BS','G_Receipt') '(AND (?:RDN10|IGN10)\.AftLineNum = (?:RDN1|IGN1)\.VisOrder\r?\n[^\r\n]*\bpj\b[^\r\n]*\r?\n)[ \t]*LEFT JOIN (?:DRF10 )?(?:RDN10|IGN10) ON [^\r\n]*\r?\n' '${1}'

# --- text rows keyed off LineNum instead of VisOrder
Tr 'T2-linenum' 'AR_Retrun_BS' 'RDN1\.LineNum = RDN10\.AftLineNum' 'RDN1.VisOrder = RDN10.AftLineNum'

# Ad-hoc mode: one transform supplied on the command line, built-ins ignored.
if ($Pattern) {
    $T.Clear()
    $aa = if ($Alias.Count) { @($Alias | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ }) } else { $null }
    Tr 'ADHOC' $aa $Pattern $Replacement $FileRx
    $Only = @()
}

if ($List) {
    Write-Host ("{0,-22} {1,-38} {2}" -f 'ID','ALIAS / FILE FILTER','PATTERN')
    foreach ($tr in $T) {
        $flt = if ($null -eq $tr.Alias) { '<all aliases>' } else { (@($tr.Alias) -join ',') }
        if ($tr.FileRx) { $flt += "  file=~$($tr.FileRx)" }
        $pat = $tr.Pat; if ($pat.Length -gt 60) { $pat = $pat.Substring(0,60) + '...' }
        Write-Host ("{0,-22} {1,-38} {2}" -f $tr.Id,$flt,$pat)
    }
    exit 0
}
function Wanted { param($Tr,[string]$Alias,[string]$FileName)
    if ($Tr.FileRx -and ($FileName -notmatch $Tr.FileRx)) { return $false }
    if ($null -eq $Tr.Alias) { return $true }
    foreach ($a in @($Tr.Alias)) { if ($Alias -eq $a) { return $true } }
    return $false }

# Every command table of a loaded report: main plus each subreport.
function Get-Cmds { param($Doc)
    $out = New-Object System.Collections.ArrayList
    $rcd = $Doc.ReportClientDocument
    $scopes = @([pscustomobject]@{ S='main'; D=$rcd.DatabaseController })
    $names = @(); try { $names = $rcd.SubreportController.GetSubreportNames() } catch {}
    foreach ($n in $names) { try { $scopes += [pscustomobject]@{ S="sub:$n"; D=$rcd.SubreportController.GetSubreport($n).DatabaseController } } catch {} }
    foreach ($sc in $scopes) {
        $c = 0; try { $c = [int]$sc.D.Database.Tables.Count } catch {}
        for ($i=0; $i -lt $c; $i++) {
            $t = $sc.D.Database.Tables[$i]
            $sql = ''; try { $sql = [string]$t.CommandText } catch {}
            if ([string]::IsNullOrEmpty($sql)) { continue }
            $at = $null; try { $at = $t.ConnectionInfo.Attributes } catch {}
            [void]$out.Add([pscustomobject]@{
                Scope=$sc.S; Alias=[string]$t.Alias; Fields=[int]$t.DataFields.Count
                Server=(BagGet $at 'QE_ServerDescription'); Db=(BagGet $at 'QE_DatabaseName')
                User=$(try { [string]$t.ConnectionInfo.UserName } catch { '' })
                Sql=$sql; Table=$t; Dbc=$sc.D; Index=$i })
        }
    }
    return ,$out }

# ------------------------------------------------------------------ run
$rpts = Get-ChildItem -LiteralPath $Root -Filter *.rpt -File -Recurse
$mode = if ($DryRun) { 'DRYRUN' } else { 'APPLY' }
Log 'INFO' "=== Apply-SqlFixes  mode=$mode  root=$Root  files=$($rpts.Count) ==="

$filesTouched = 0; $tablesTouched = 0; $errCount = 0
$fieldDiff = 0; $connDiff = 0; $sqlDiff = 0; $userLost = 0
$rKeep = 0; $rForced = 0; $rFallback = 0; $rNoUser = 0

foreach ($f in $rpts) {
    $full = $f.FullName
    $doc  = New-Object CrystalDecisions.CrystalReports.Engine.ReportDocument
    $before = @{}; $intended = @{}; $beforeCount = 0; $fileTables = 0
    $anyChange = $false; $written = @{}
    try {
        $doc.Load($full,[CrystalDecisions.Shared.OpenReportMethod]::OpenReportByTempCopy)
        $cmds = Get-Cmds -Doc $doc
        $beforeCount = $cmds.Count
        foreach ($c in $cmds) {
            $key = "$($c.Scope)|$($c.Alias)"
            $before[$key]   = [pscustomobject]@{ Fields=$c.Fields; Server=$c.Server; Db=$c.Db; User=$c.User }
            $intended[$key] = $c.Sql
            $sql = $c.Sql
            $applied = @()
            foreach ($tr in $T) {
                if ($Only.Count) {
                    $hit = $false
                    foreach ($o in $Only) { if ($tr.Id -like "$o*") { $hit = $true; break } }
                    if (-not $hit) { continue }
                }
                if (-not (Wanted $tr $c.Alias $f.Name)) { continue }
                $h = ([regex]::Matches($sql,$tr.Pat,'IgnoreCase')).Count
                if ($h -eq 0) { continue }
                $new = [regex]::Replace($sql,$tr.Pat,$tr.Rep,'IgnoreCase')
                if ($new -eq $sql) { continue }   # idempotent transform, nothing to do
                $tr.Hits += $h; $tr.Tables++
                $applied += ("{0}x{1}" -f $tr.Id,$h)
                $sql = $new
            }
            if ($sql -ne $c.Sql) {
                Log 'HIT' "    [$($c.Scope)] $($c.Alias): $($applied -join ', ')  (len $($c.Sql.Length) -> $($sql.Length))"
                $intended[$key] = $sql
                $fileTables++
                $anyChange = $true
                if (-not $DryRun) {
                    # Re-fetch by index: an earlier SetTableLocation may have
                    # replaced the item in the collection.
                    $route = Set-CommandSql -Dbc $c.Dbc -Index $c.Index -NewSql $sql -Force $UserID
                    switch ($route) {
                        'keepuser' { $rKeep++ }
                        'forced'   { $rForced++ }
                        'fallback' { $rFallback++; Log 'ERR'  "        *** User ID '$($c.User)' WILL BE LOST on [$($c.Scope)] $($c.Alias) -- SAP B1 may refuse this layout ***" }
                        'nouser'   { $rNoUser++;   Log 'WARN' "        [$($c.Scope)] $($c.Alias) has NO User ID already -- nothing to lose (pass -UserID sa to set one)" }
                    }
                    $written[$key] = $true
                }
            }
        }

        # -UserID re-stamps every OTHER command table of a file that is about to
        # be saved: one blank User ID anywhere stops SAP B1 opening the layout.
        if ($UserID -and -not $DryRun -and $anyChange) {
            foreach ($c in $cmds) {
                $key = "$($c.Scope)|$($c.Alias)"
                if ($written[$key]) { continue }
                $r2 = Set-CommandSql -Dbc $c.Dbc -Index $c.Index -NewSql $c.Sql -Force $UserID
                if ($r2 -eq 'forced') { $rForced++ } else { $rFallback++ }
                Log 'OK' "        [$($c.Scope)] $($c.Alias): User ID stamped '$UserID' (SQL unchanged)"
                $written[$key] = $true
            }
        }
        if (-not $anyChange) { continue }
        if ($DryRun) { $filesTouched++; $tablesTouched += $fileTables; continue }
        Log 'INFO' "--> $full"
        $tmp = Join-Path $env:TEMP ('_asf_' + [guid]::NewGuid().ToString('N') + '.rpt')
        Copy-Item -LiteralPath $full -Destination $tmp -Force
        $doc.SaveAs($full)
    } catch {
        Log 'ERR' "    $($f.Name): $($_.Exception.Message)"
        $errCount++
        try { $doc.Close(); $doc.Dispose() } catch {}
        continue
    } finally {
        try { $doc.Close(); $doc.Dispose() } catch {}
    }
    if ($DryRun) { continue }

    # ---- re-open from disk and prove every table ----
    $fe = 0
    $d2 = New-Object CrystalDecisions.CrystalReports.Engine.ReportDocument
    try {
        $d2.Load($full,[CrystalDecisions.Shared.OpenReportMethod]::OpenReportByTempCopy)
        $after = Get-Cmds -Doc $d2
        if ($after.Count -ne $beforeCount) { Log 'ERR' "    VERIFY table count $beforeCount -> $($after.Count)"; $fe++ }
        foreach ($c in $after) {
            $key = "$($c.Scope)|$($c.Alias)"
            if (-not $before.ContainsKey($key)) { Log 'ERR' "    VERIFY unexpected table $key"; $fe++; continue }
            $b = $before[$key]
            if ($c.Fields -ne $b.Fields) { Log 'ERR' "    VERIFY fieldDiff $key : $($b.Fields) -> $($c.Fields)"; $fieldDiff++; $fe++ }
            if ($c.Server -ne $b.Server -or $c.Db -ne $b.Db) { Log 'ERR' "    VERIFY connDiff $key : '$($b.Server)/$($b.Db)' -> '$($c.Server)/$($c.Db)'"; $connDiff++; $fe++ }
            if ($c.Sql -ne $intended[$key]) { Log 'ERR' "    VERIFY sqlDiff $key : on-disk len $($c.Sql.Length) != intended $($intended[$key].Length)"; $sqlDiff++; $fe++ }
            $wantUser = if ($UserID) { $UserID } else { $b.User }
            if ($c.User -ne $wantUser) { Log 'ERR' "    VERIFY userLost $key : User ID '$wantUser' -> '$($c.User)'"; $userLost++; $fe++ }
        }
    } catch { Log 'ERR' "    VERIFY load failed: $($_.Exception.Message)"; $fe++ }
    finally { try { $d2.Close(); $d2.Dispose() } catch {} }

    if ($fe -gt 0) {
        Copy-Item -LiteralPath $tmp -Destination $full -Force
        Log 'ERR' "    ROLLED BACK $($f.Name) -- $fe verification error(s)"
        $errCount += $fe
    } else {
        Log 'OK' "    verified on disk: $fileTables table(s)"
        $filesTouched++; $tablesTouched += $fileTables
    }
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
}

Log 'INFO' '=============================================================='
foreach ($tr in $T) { Log 'INFO' ("  {0,-22} tables={1,-3} hits={2}" -f $tr.Id,$tr.Tables,$tr.Hits) }
Log 'INFO' "SUMMARY mode=$mode files=$($rpts.Count) touched=$filesTouched tables=$tablesTouched fieldDiff=$fieldDiff connDiff=$connDiff sqlDiff=$sqlDiff userLost=$userLost err=$errCount"
Log 'INFO' "USER ID ROUTES  keepuser=$rKeep forced=$rForced fallback=$rFallback noUserToStart=$rNoUser"
if ($rFallback) { Log 'ERR' "*** $rFallback table(s) lost their User ID -- those layouts may not open in SAP B1 ***" }
Write-Host "RESULT|$mode|files=$($rpts.Count)|touched=$filesTouched|tables=$tablesTouched|fieldDiff=$fieldDiff|connDiff=$connDiff|sqlDiff=$sqlDiff|userLost=$userLost|keep=$rKeep|forced=$rForced|fallback=$rFallback|nouser=$rNoUser|err=$errCount"
exit ([int]([bool]$errCount))
