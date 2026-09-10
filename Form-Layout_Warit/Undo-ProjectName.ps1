<#
.SYNOPSIS
  Removes the 'ProjectName' column that Fix-ReportSql.ps1 -ProjectName (T4) added,
  putting a report's Commands back to exactly what they were before that step.

.DESCRIPTION
  T4 is a pure insertion. It always writes the same two lines immediately before
  a branch's FROM keyword:

        ,
            <expression> AS 'ProjectName'

  so it can be taken out again character-for-character, without needing the
  .bak files (several of which no longer exist on this machine).

  What is removed, and nothing else:
    - a line whose only content is a comma
    - followed by one line ending in  AS 'ProjectName'
  Both must sit directly above a FROM line. Anything else that mentions
  ProjectName is left alone and reported, so a hand-written column can never be
  deleted by accident.

  Same safety rules as the other scripts here: back up first, write through
  clone + SetTableLocation (a direct CommandText assignment does not persist),
  carry the UserID across in the nested logon bag, then verify by reading the
  file back from DISK and roll back if anything does not match.

.EXAMPLE
  .\Undo-ProjectName.ps1 -Path "C:\GitHub\SDA\Form-Layout_SDA\2. Sales - AR\7. AR Credit Memo" -WhatIf
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $Path,
    [switch] $WhatIf,
    [string] $Filter       = '*.rpt',
    [switch] $NoRecurse,
    [string] $BackupSuffix = '.20260901-undopn.bak',
    [string] $LogFile      = "$PSScriptRoot\_UndoProjectName.log"
)

$ErrorActionPreference = 'Stop'
$script:BF = [System.Reflection.BindingFlags]
if ([IntPtr]::Size -ne 8) { Write-Host '[ERROR] 64-bit PowerShell required.' -ForegroundColor Red; exit 2 }

function Load-CrystalSDK {
    $gacMsil = 'C:\Windows\Microsoft.NET\assembly\GAC_MSIL'
    $gac64   = 'C:\Windows\Microsoft.NET\assembly\GAC_64'
    $sets = @(
        @{ Root = $gacMsil; Names = @('CrystalDecisions.Shared','CrystalDecisions.ReportSource','CrystalDecisions.CrystalReports.Engine') },
        @{ Root = $gac64;   Names = @('CrystalDecisions.ReportAppServer.CommLayer','CrystalDecisions.ReportAppServer.DataDefModel',
                                      'CrystalDecisions.ReportAppServer.Controllers','CrystalDecisions.ReportAppServer.ClientDoc',
                                      'CrystalDecisions.ReportAppServer.CommonObjectModel','CrystalDecisions.ReportAppServer.ObjectFactory') }
    )
    foreach ($set in $sets) {
        foreach ($n in $set.Names) {
            $dir = Join-Path $set.Root $n
            if (-not (Test-Path $dir)) { throw "GAC folder not found: $dir" }
            $dll = Get-ChildItem $dir -Recurse -Filter "$n.dll" |
                   Sort-Object { [version](($_.Directory.Name -split '_')[1]) } -Descending | Select-Object -First 1
            [void][System.Reflection.Assembly]::LoadFrom($dll.FullName)
        }
    }
}
try { Load-CrystalSDK } catch { Write-Host "[ERROR] SDK load failed: $($_.Exception.Message)" -ForegroundColor Red; exit 3 }

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

function BagItem { param($B,$K) if ($null -eq $B) { return $null }; try { return $B.Item($K) } catch { return $null } }
function BagHas  { param($B,$K) if ($null -eq $B) { return $false }; try { return [bool]$B.Contains($K) } catch { return $false } }
function BagSetStr {
    param($B,$K,$V)
    if ($null -eq $B) { return $false }
    try { if (BagHas $B $K) { [void]$B.Remove($K) } } catch {}
    try { [void]$B.Add($K, [string]$V); return $true } catch { return $false }
}
function Read-TableUser {
    param($Table)
    try { $ci = $Table.ConnectionInfo; return [string]$ci.GetType().InvokeMember('UserName', $script:BF::GetProperty, $null, $ci, $null) }
    catch { return '' }
}

# the exact shape T4 wrote: a lone comma line, the column line, then FROM
$UndoRx = "(?m)^[ \t]*,[ \t]*\r?\n[ \t]*.*AS[ \t]+'ProjectName'[ \t]*\r?\n(?=[ \t]*FROM\b)"

function Remove-ProjectName {
    param([string]$Sql)
    if ($Sql -notmatch "(?i)AS\s+'ProjectName'") { return $null }
    $rx  = New-Object System.Text.RegularExpressions.Regex $UndoRx
    $hit = $rx.Matches($Sql).Count
    if ($hit -eq 0) { return $null }
    $new = $rx.Replace($Sql, '')
    $left = [regex]::Matches($new, "(?i)AS\s+'ProjectName'").Count
    if ($left -gt 0) { throw "Undo: $left 'ProjectName' column(s) do not match the shape T4 writes -- leaving this command alone" }
    return $new
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
Log 'INFO' "=== Undo-ProjectName start  mode=$mode ==="
Log 'INFO' "Path=$Path  files=$($rpts.Count)"

$scanned = 0; $changed = 0; $failed = 0; $cmds = 0
foreach ($f in $rpts) {
    $full = $f.FullName
    $scanned++
    $doc = $null
    $plan = New-Object System.Collections.ArrayList
    try {
        $doc = New-Object CrystalDecisions.CrystalReports.Engine.ReportDocument
        $doc.Load($full, [CrystalDecisions.Shared.OpenReportMethod]::OpenReportByTempCopy)
        $dbc = $doc.ReportClientDocument.DatabaseController
        $n = 0; try { $n = [int]$dbc.Database.Tables.Count } catch {}
        for ($i = 0; $i -lt $n; $i++) {
            $t = $dbc.Database.Tables[$i]
            $sql = ''; try { $sql = [string]$t.CommandText } catch {}
            if (-not $sql) { continue }
            $new = $null
            try { $new = Remove-ProjectName -Sql $sql }
            catch { Log 'ERR' "    [$($t.Alias)] $($_.Exception.Message)"; continue }
            if (-not $new) { continue }
            [void]$plan.Add([pscustomobject]@{
                Index = $i; Alias = [string]$t.Alias; OldLen = $sql.Length; NewSql = $new
                User = (Read-TableUser -Table $t)
            })
        }
    } catch {
        Log 'ERR' "--> $full : load failed: $($_.Exception.Message.Split([char]10)[0])"
        $failed++
        if ($doc) { try { $doc.Close() } catch {} }
        continue
    }

    if ($plan.Count -eq 0) { if ($doc) { try { $doc.Close() } catch {} }; continue }

    Log 'HIT' "--> $full"
    foreach ($p in $plan) { Log 'INFO' ("    [{0}] {1} -> {2} chars  User='{3}'" -f $p.Alias, $p.OldLen, $p.NewSql.Length, $p.User) }
    if ($WhatIf) { $cmds += $plan.Count; if ($doc) { try { $doc.Close() } catch {} }; continue }

    $bak = "$full$BackupSuffix"
    Copy-Item -LiteralPath $full -Destination $bak -Force
    try {
        $dbc = $doc.ReportClientDocument.DatabaseController
        foreach ($p in $plan) {
            $live  = $dbc.Database.Tables[$p.Index]
            $clone = $live.Clone($true)
            $clone.CommandText = $p.NewSql
            if ($p.User) {
                $lp = BagItem $clone.ConnectionInfo.Attributes 'QE_LogonProperties'
                if ($lp) { [void](BagSetStr $lp 'User ID' $p.User) }
            }
            $dbc.SetTableLocation($live, $clone)
        }
        $doc.SaveAs([string]$full)
    } catch {
        Log 'ERR' "    write failed: $($_.Exception.Message.Split([char]10)[0])"
        try { $doc.Close() } catch {}
        Copy-Item -LiteralPath $bak -Destination $full -Force
        Log 'ERR' "    ROLLED BACK"
        $failed++
        continue
    }
    try { $doc.Close() } catch {}

    $probs = @()
    $d2 = $null
    try {
        $d2 = New-Object CrystalDecisions.CrystalReports.Engine.ReportDocument
        $d2.Load($full, [CrystalDecisions.Shared.OpenReportMethod]::OpenReportByTempCopy)
        $c2 = $d2.ReportClientDocument.DatabaseController
        foreach ($p in $plan) {
            $t2 = $c2.Database.Tables[$p.Index]
            if ([string]$t2.Alias -ne $p.Alias) { $probs += "index moved for $($p.Alias)"; continue }
            $got = ''; try { $got = [string]$t2.CommandText } catch {}
            if ($got -ne $p.NewSql) { $probs += "$($p.Alias): disk has $($got.Length) chars, wanted $($p.NewSql.Length)" }
            $gu = Read-TableUser -Table $t2
            if ($p.User -and $gu -ne $p.User) { $probs += "$($p.Alias): UserID '$($p.User)' -> '$gu'" }
        }
    } catch { $probs += "verify failed: $($_.Exception.Message.Split([char]10)[0])" }
    finally { if ($d2) { try { $d2.Close() } catch {} } }

    if ($probs.Count) {
        foreach ($pr in $probs) { Log 'ERR' "    VERIFY $pr" }
        Copy-Item -LiteralPath $bak -Destination $full -Force
        Log 'ERR' "    ROLLED BACK"
        $failed++
        continue
    }
    Log 'OK' "    verified on disk: $($plan.Count) command(s) restored"
    $changed++; $cmds += $plan.Count
}

Log 'INFO' '=============================================================='
Log 'INFO' ("SUMMARY mode={0}  files={1}  changed={2}  failed={3}  commands={4}" -f $mode, $scanned, $changed, $failed, $cmds)
Log 'INFO' '=============================================================='
Write-Host ("RESULT|{0}|files={1}|changed={2}|failed={3}|commands={4}" -f $mode, $scanned, $changed, $failed, $cmds)
if ($failed -gt 0) { exit 1 }
exit 0
