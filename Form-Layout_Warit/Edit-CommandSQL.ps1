<#
.SYNOPSIS
  Bulk edit SQL Command text inside Crystal Report (.rpt) files via the RAS API.

.DESCRIPTION
  Scans a folder (or a single .rpt), opens each report through the Crystal
  Reports .NET SDK / ReportAppServer (RAS), walks every Command table in the
  main report AND in every subreport, applies a case-insensitive regex
  replacement to the table's CommandText, saves the report, then CLOSES and
  RE-OPENS the file from disk to verify the result.

  Verification after every save compares, against the pre-save snapshot:
    (a) table count is unchanged
    (b) field count of every table is unchanged   -> fieldDiff
    (c) connection info is unchanged              -> connDiff
        (QE_ServerDescription, QE_DatabaseName)
    (d) the new CommandText on disk equals the intended SQL -> errCount

  HOW THE SQL IS WRITTEN -- measured on this repo 2026-08-28, do not change it
  back:
      $table.CommandText = $newSql        -> reading it back IN MEMORY shows the
                                             new length, but the SAVED FILE still
                                             holds the OLD SQL. Silent data loss.
                                             An earlier version of this script
                                             used exactly this and rolled back all
                                             43 reports it touched.
      InvokeMember('CommandText',...)     -> same silent failure.
      clone = table.Clone($true)
      clone.CommandText = $newSql
      Dbc.SetTableLocation($live, $clone) -> ACTUALLY PERSISTS.
  Database.Tables[$i] hands back a COPY of the table, so mutating it changes
  nothing in the report; SetTableLocation is what puts the modified table back.
  Same rule as for connection changes.

  SetSQLCommandTable(ci, alias, sql) is still unusable here -- it needs a live
  database connection.

  KEEPING THE CONNECTION'S User ID -- measured 2026-08-31 on C:\_bakchk\B25.rpt
  (a report whose 3 tables all carry UserID='sa'). A .rpt whose UserID is blank
  CANNOT BE OPENED BY SAP B1, so losing it is real data loss:
      SetTableLocation(live, clone)        -> SQL persists but UserID is WIPED.
                                              Happens even with an UNMODIFIED
                                              clone, so SetTableLocation itself
                                              is what poisons it, not the SQL.
      clone.ConnectionInfo.UserName = 'sa' -> in memory only; saved file has ''.
      table.CommandText = sql, in any order with ApplyLogOnInfo, saved through
      the engine or through rcd.SaveAs, inside BeginTransaction/EndTransaction
                                           -> never reaches the file at all.
      ApplyLogOnInfo / ModifyTableConnectionInfo(alias,ci) / DataSourceConnections
      .SetLogon / .SetConnection / SetDatabaseLogon / SetTableLocationByServer-
      DatabaseName / ReplaceConnection      -> none of them can put a UserID back
                                              once it reads as ''.
      clone.ConnectionInfo.Attributes['QE_LogonProperties'].Add('User ID', $u)
        + Attributes.Add('QE_LogonProperties', $lp)
        + clone.ConnectionInfo = $ci
        + SetTableLocation(live, clone)     -> SQL PERSISTS **AND** UserID SURVIVES.
  The nested QE_LogonProperties bag is the only writable route. Crystal consumes
  the 'User ID' entry on save, so the bag on disk reads exactly as before --
  params, formulas, record selection, table links, field counts and the outer
  attribute bag were all verified identical afterwards.
  The same trick RESTORES a UserID on a report that has none: pass -UserID.

.PARAMETER Root
  Folder to scan recursively, OR a single .rpt path.

.PARAMETER Pattern
  .NET regex, matched case-insensitively against CommandText.

.PARAMETER Replacement
  Replacement string (.NET substitution syntax: $1, ${name} ...).
  Ignored in -Verify mode.

.PARAMETER Alias
  Limit to these table aliases (wildcards allowed). Empty = every alias.

.PARAMETER File
  Limit to reports whose FILE NAME matches one of these wildcard patterns.
  Empty = every .rpt under -Root.

.PARAMETER WhatIf
  Count the matches and show sample lines. Nothing is written.

.PARAMETER Verify
  Read-only. Reports how many times -Pattern currently occurs, per table.

.PARAMETER LogPath
  Log file. Default <script-dir>\_EditCommandSQL.log
#>
param(
    [string]   $Root        = '',
    [string]   $Pattern     = '',
    [string]   $Replacement = '',
    [string[]] $Alias       = @(),
    [string[]] $File        = @(),
    [switch]   $WhatIf,
    [switch]   $Verify,
    [string]   $UserID      = '',
    [string]   $LogPath     = ''
)

$ErrorActionPreference = 'Stop'

if (-not $Root)    { Write-Host '[ERROR] -Root is required.'    -ForegroundColor Red; exit 4 }
if (-not $Pattern) { Write-Host '[ERROR] -Pattern is required.' -ForegroundColor Red; exit 4 }

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $PSCommandPath }
if (-not $LogPath) { $LogPath = Join-Path $scriptDir '_EditCommandSQL.log' }

# ------------------------------------------------------------------ logging
$logDir = Split-Path $LogPath -Parent
if ($logDir -and -not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}
function Log {
    param([string]$Level, [string]$Msg)
    $ts   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$ts][$Level] $Msg"
    switch ($Level) {
        'ERR'  { Write-Host $line -ForegroundColor Red }
        'WARN' { Write-Host $line -ForegroundColor Yellow }
        'OK'   { Write-Host $line -ForegroundColor Green }
        'HIT'  { Write-Host $line -ForegroundColor Cyan }
        default{ Write-Host $line }
    }
    try { [System.IO.File]::AppendAllText($LogPath, $line + "`r`n", [System.Text.Encoding]::UTF8) } catch {}
}

# ------------------------------------------------------------------ bitness
if (-not [Environment]::Is64BitProcess) {
    Write-Host '[ERROR] Crystal RAS assemblies live in GAC_64. Re-run under 64-bit powershell.exe' -ForegroundColor Red
    exit 2
}

# ------------------------------------------------------------------ SDK load
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

# ------------------------------------------------------------------ helpers
# NOTE: RAS objects reached through ReportDocument.ReportClientDocument are
# System.__ComObject (late-bound), not managed types. Attributes has NO
# .Collection property and ConnectionInfo.ServerName always reads back empty --
# the string indexer is the only thing that works. Reading any other way returns
# '' for every table and makes connDiff meaningless.
# Writes NewSql into the command table at $Index and puts the table back with
# SetTableLocation, carrying the connection's User ID across in the nested
# QE_LogonProperties bag so that the save does not blank it. See header.
# Returns the route taken: keepuser | forced | fallback | nouser
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

function Get-ConnAttr {
    param($Table, [string]$Key)
    try {
        $v = $Table.ConnectionInfo.Attributes[$Key]
        if ($null -eq $v) { return '' }
        return [string]$v
    } catch {}
    return ''
}

# Every COMMAND table of the report: the main report plus each subreport.
# Dbc + Index are carried so the caller can re-fetch the LIVE table object --
# the one in this list is a copy and writing to it does nothing.
function Get-CommandTables {
    param($Doc)
    $out = New-Object System.Collections.ArrayList
    $rcd = $Doc.ReportClientDocument

    $scopes = @([pscustomobject]@{ S = 'main'; D = $rcd.DatabaseController })
    $subNames = @()
    try { $subNames = $rcd.SubreportController.GetSubreportNames() } catch {}
    foreach ($sn in $subNames) {
        try { $scopes += [pscustomobject]@{ S = "sub:$sn"; D = $rcd.SubreportController.GetSubreport($sn).DatabaseController } }
        catch { Log 'WARN' "    cannot open subreport '$sn': $($_.Exception.Message)" }
    }

    foreach ($sc in $scopes) {
        $cnt = 0
        try { $cnt = [int]$sc.D.Database.Tables.Count } catch { $cnt = 0 }
        for ($i = 0; $i -lt $cnt; $i++) {
            $t = $sc.D.Database.Tables[$i]
            $sql = ''
            try { $sql = [string]$t.CommandText } catch {}
            if ([string]::IsNullOrEmpty($sql)) { continue }
            [void]$out.Add([pscustomobject]@{
                Scope  = $sc.S
                Alias  = [string]$t.Alias
                Fields = [int]$t.DataFields.Count
                Server = (Get-ConnAttr -Table $t -Key 'QE_ServerDescription')
                Db     = (Get-ConnAttr -Table $t -Key 'QE_DatabaseName')
                User   = $(try { [string]$t.ConnectionInfo.UserName } catch { '' })
                Sql    = $sql
                Table  = $t
                Dbc    = $sc.D
                Index  = $i
            })
        }
    }
    return ,$out
}
function Test-AliasWanted {
    param([string]$A)
    if (-not $Alias -or $Alias.Count -eq 0) { return $true }
    foreach ($p in $Alias) { if ($A -like $p) { return $true } }
    return $false
}

function Show-Sample {
    param([string]$Sql, [string]$Pat, [int]$Max = 3)
    $lines = $Sql -split "`r?`n"
    $n = 0
    foreach ($ln in $lines) {
        if ([regex]::IsMatch($ln, $Pat, 'IgnoreCase')) {
            $txt = $ln.Trim()
            if ($txt.Length -gt 160) { $txt = $txt.Substring(0,160) + ' ...' }
            Log 'HIT' "        | $txt"
            $n++
            if ($n -ge $Max) { break }
        }
    }
}

# ------------------------------------------------------------------ enumerate
$rpts = @()
if (Test-Path -LiteralPath $Root -PathType Leaf) {
    if ($Root -like '*.rpt') { $rpts = ,(Get-Item -LiteralPath $Root) }
    else { Log 'ERR' "Root is a file but not .rpt: $Root"; exit 4 }
} elseif (Test-Path -LiteralPath $Root -PathType Container) {
    $rpts = Get-ChildItem -LiteralPath $Root -Filter '*.rpt' -File -Recurse
} else {
    Log 'ERR' "Root not found: $Root"; exit 4
}

if ($File -and $File.Count -gt 0) {
    $rpts = $rpts | Where-Object {
        $nm = $_.Name
        $hit = $false
        foreach ($p in $File) { if ($nm -like $p) { $hit = $true; break } }
        $hit
    }
}

$mode = if ($Verify) { 'VERIFY' } elseif ($WhatIf) { 'WHATIF' } else { 'APPLY' }
Log 'INFO' "=== Edit-CommandSQL start  mode=$mode ==="
Log 'INFO' "Root=$Root"
Log 'INFO' "Pattern=$Pattern"
if (-not $Verify) { Log 'INFO' "Replacement=$Replacement" }
if ($Alias.Count) { Log 'INFO' "Alias filter: $($Alias -join ', ')" }
if ($File.Count)  { Log 'INFO' "File filter : $($File -join ', ')" }
Log 'INFO' "Matched $($rpts.Count) report file(s)."

# ------------------------------------------------------------------ counters
$filesScanned  = 0
$filesTouched  = 0
$tablesTouched = 0
$replacements  = 0
$fieldDiff     = 0
$connDiff      = 0
$errCount      = 0
$userLost      = 0
$rKeep         = 0
$rForced       = 0
$rFallback     = 0
$rNoUser       = 0

foreach ($f in $rpts) {
    $filesScanned++
    $full = $f.FullName
    Log 'INFO' "--> $full"

    $doc = New-Object CrystalDecisions.CrystalReports.Engine.ReportDocument
    $intended = @{}     # key -> expected new SQL
    $before   = @{}     # key -> snapshot
    $beforeCount = 0
    $fileHits = 0
    $fileTables = 0
    $written  = @{}    # key -> $true for every table pushed through Set-CommandSql

    try {
        $doc.Load($full, [CrystalDecisions.Shared.OpenReportMethod]::OpenReportByTempCopy)
        $cmds = Get-CommandTables -Doc $doc
        $beforeCount = $cmds.Count

        foreach ($c in $cmds) {
            $key = "$($c.Scope)|$($c.Alias)"
            $before[$key] = [pscustomobject]@{
                Fields = $c.Fields; Server = $c.Server; Db = $c.Db; Sql = $c.Sql
                User   = $c.User
            }
            $intended[$key] = $c.Sql

            if (-not (Test-AliasWanted -A $c.Alias)) { continue }

            $hits = ([regex]::Matches($c.Sql, $Pattern, 'IgnoreCase')).Count
            if ($hits -eq 0) { continue }

            $fileHits  += $hits
            $fileTables++
            Log 'HIT' "    [$($c.Scope)] $($c.Alias): $hits match(es)  (sqlLen=$($c.Sql.Length), fields=$($c.Fields))"

            if ($Verify -or $WhatIf) {
                Show-Sample -Sql $c.Sql -Pat $Pattern
                continue
            }

            $newSql = [regex]::Replace($c.Sql, $Pattern, $Replacement, 'IgnoreCase')
            if ($newSql -eq $c.Sql) {
                Log 'WARN' "    [$($c.Scope)] $($c.Alias): replacement produced identical SQL -- skipped"
                $fileHits -= $hits; $fileTables--
                continue
            }
            # Re-fetch the live table by index, modify a clone, put it back.
            # Assigning $c.Table.CommandText would look fine in memory and never
            # reach the file -- see the header.
            $route = Set-CommandSql -Dbc $c.Dbc -Index $c.Index -NewSql $newSql -Force $UserID
            switch ($route) {
                'keepuser' { $rKeep++ }
                'forced'   { $rForced++;   Log 'OK'   "        User ID forced to '$UserID'" }
                'fallback' { $rFallback++; Log 'ERR'  "        *** User ID '$($c.User)' WILL BE LOST on [$($c.Scope)] $($c.Alias) -- SAP B1 may refuse to open this layout ***" }
                'nouser'   { $rNoUser++;   Log 'WARN' "        [$($c.Scope)] $($c.Alias) has NO User ID already -- nothing to lose (pass -UserID sa to set one)" }
            }
            $intended[$key] = $newSql
            $written[$key] = $true
        }

        # -UserID also re-stamps every OTHER command table of a report that is
        # about to be saved. A blank User ID anywhere in the file is enough to
        # stop SAP B1 opening the layout, so fix the whole file or none of it.
        if ($UserID -and -not ($Verify -or $WhatIf) -and $fileHits -gt 0) {
            foreach ($c in $cmds) {
                $key = "$($c.Scope)|$($c.Alias)"
                if ($written[$key]) { continue }
                $r2 = Set-CommandSql -Dbc $c.Dbc -Index $c.Index -NewSql $c.Sql -Force $UserID
                if ($r2 -eq 'forced') { $rForced++ } else { $rFallback++ }
                Log 'OK' "        [$($c.Scope)] $($c.Alias): User ID stamped '$UserID' (SQL unchanged)"
                $written[$key] = $true
            }
        }

        if ($Verify -or $WhatIf) {
            if ($fileHits -gt 0) {
                Log 'INFO' "    ${mode}: $fileHits match(es) in $fileTables table(s) [no write]"
                $filesTouched++; $tablesTouched += $fileTables; $replacements += $fileHits
            }
            continue
        }

        if ($fileHits -eq 0) { Log 'INFO' '    no match -- not saved'; continue }

        $doc.SaveAs($full)
        Log 'OK' "    saved ($fileHits replacement(s) in $fileTables table(s))"
    }
    catch {
        Log 'ERR' "    $($f.Name): $($_.Exception.Message)"
        $errCount++
        continue
    }
    finally {
        try { $doc.Close(); $doc.Dispose() } catch {}
    }

    # ---------------- re-open from disk and verify ----------------
    $doc2 = New-Object CrystalDecisions.CrystalReports.Engine.ReportDocument
    try {
        $doc2.Load($full, [CrystalDecisions.Shared.OpenReportMethod]::OpenReportByTempCopy)
        $after = Get-CommandTables -Doc $doc2

        if ($after.Count -ne $beforeCount) {
            Log 'ERR' "    VERIFY: command-table count changed $beforeCount -> $($after.Count)"
            $errCount++
        }
        foreach ($c in $after) {
            $key = "$($c.Scope)|$($c.Alias)"
            if (-not $before.ContainsKey($key)) {
                Log 'ERR' "    VERIFY: unexpected new table '$key'"; $errCount++; continue
            }
            $b = $before[$key]
            if ($c.Fields -ne $b.Fields) {
                Log 'ERR' "    VERIFY fieldDiff: $key fields $($b.Fields) -> $($c.Fields)"
                $fieldDiff++; $errCount++
            }
            if ($c.Server -ne $b.Server -or $c.Db -ne $b.Db) {
                Log 'ERR' "    VERIFY connDiff: $key '$($b.Server)/$($b.Db)' -> '$($c.Server)/$($c.Db)'"
                $connDiff++; $errCount++
            }
            $wantUser = if ($UserID) { $UserID } else { $b.User }
            if ($c.User -ne $wantUser) {
                Log 'ERR' "    VERIFY userLost: $key User ID '$wantUser' -> '$($c.User)'"
                $userLost++; $errCount++
            }
            if ($c.Sql -ne $intended[$key]) {
                Log 'ERR' "    VERIFY sqlDiff: $key on-disk SQL != intended (len $($c.Sql.Length) vs $($intended[$key].Length))"
                $errCount++
            }
        }
        Log 'OK' "    verified on disk: $($after.Count) command table(s)"
        $filesTouched++
        $tablesTouched += $fileTables
        $replacements  += $fileHits
    }
    catch {
        Log 'ERR' "    VERIFY load failed: $($_.Exception.Message)"
        $errCount++
    }
    finally {
        try { $doc2.Close(); $doc2.Dispose() } catch {}
    }
}

Log 'INFO' '=============================================================='
Log 'INFO' "SUMMARY mode=$mode  filesScanned=$filesScanned  filesTouched=$filesTouched  tablesTouched=$tablesTouched  replacements=$replacements  fieldDiff=$fieldDiff  connDiff=$connDiff  userLost=$userLost  errCount=$errCount"
Log 'INFO' "USER ID ROUTES  keepuser=$rKeep  forced=$rForced  fallback=$rFallback  noUserToStart=$rNoUser"
if ($rFallback) { Log 'ERR' "*** $rFallback table(s) lost their User ID -- those layouts may not open in SAP B1 ***" }
Log 'INFO' '=============================================================='
Write-Host "RESULT|$mode|files=$filesScanned|touched=$filesTouched|tables=$tablesTouched|repl=$replacements|fieldDiff=$fieldDiff|connDiff=$connDiff|userLost=$userLost|keep=$rKeep|forced=$rForced|fallback=$rFallback|nouser=$rNoUser|err=$errCount"

exit ([int]([bool]$errCount))
