<#
.SYNOPSIS
  Prints every GROUP BY clause found in .rpt Command SQL and flags any that
  contains a column alias.

.DESCRIPTION
  "GROUP BY <expr> AS <name>" is a syntax error, so a find/replace that appends
  an alias to a column will break any report that also lists that column in its
  GROUP BY. Two reports here do exactly that (AP Down Payment ENG lists the UoM
  column in both the SELECT and the GROUP BY), which is why the UoM work needed a
  separate no-alias transform for the GROUP BY copy.

  The clause is bounded BY LINES, not by regex lookahead. A regex that scans
  forward to the next UNION / ORDER BY runs straight past a subquery's
  "GROUP BY DocEntry, Parent_VisOrder" into the outer SELECT list and then flags
  every aliased column it finds there -- that produced 9 false positives the
  first time this check was written.

  Read-only. Nothing is ever written to the reports.

.PARAMETER Root
  Folder to scan recursively, or a single .rpt.

.PARAMETER FileRx
  Regex on the report FILE NAME. Blank = every report.

.PARAMETER AliasRx
  Regex on the table alias. Blank = every command table.

.PARAMETER AliasToken
  What counts as a leaked alias inside a GROUP BY. Default 'UgpCode'.

.PARAMETER Quiet
  Only print the clauses that are flagged.

.EXAMPLE
  gbcheck.ps1 -Root <repo>
  gbcheck.ps1 -Root <repo> -FileRx 'Down Payment' -Quiet
#>
param(
    [Parameter(Mandatory)] [string] $Root,
    [string] $FileRx     = '',
    [string] $AliasRx    = '',
    [string] $AliasToken = 'UgpCode',
    [switch] $Quiet
)
$ErrorActionPreference = 'Stop'
if (-not [Environment]::Is64BitProcess) { Write-Host '[ERROR] need 64-bit powershell.exe'; exit 2 }
foreach ($s in @(
 @{R='C:\Windows\Microsoft.NET\assembly\GAC_MSIL';N=@('CrystalDecisions.Shared','CrystalDecisions.ReportSource','CrystalDecisions.CrystalReports.Engine')},
 @{R='C:\Windows\Microsoft.NET\assembly\GAC_64';N=@('CrystalDecisions.ReportAppServer.CommLayer','CrystalDecisions.ReportAppServer.DataDefModel','CrystalDecisions.ReportAppServer.Controllers','CrystalDecisions.ReportAppServer.ClientDoc','CrystalDecisions.ReportAppServer.CommonObjectModel','CrystalDecisions.ReportAppServer.ObjectFactory')})) {
  foreach ($n in $s.N) {
    $d = Join-Path $s.R $n
    $dll = Get-ChildItem $d -Recurse -Filter "$n.dll" | Sort-Object { [version](($_.Directory.Name -split '_')[1]) } -Descending | Select-Object -First 1
    if (-not $dll) { Write-Host "[ERROR] assembly not found: $n"; exit 3 }
    [void][System.Reflection.Assembly]::LoadFrom($dll.FullName)
  }
}

$rpts = @()
if (Test-Path -LiteralPath $Root -PathType Leaf) { $rpts = ,(Get-Item -LiteralPath $Root) }
else { $rpts = Get-ChildItem -LiteralPath $Root -Filter *.rpt -File -Recurse }

$bad = 0; $seen = 0
foreach ($f in $rpts) {
    if ($FileRx -and ($f.Name -notmatch $FileRx)) { continue }
    $doc = New-Object CrystalDecisions.CrystalReports.Engine.ReportDocument
    try {
        $doc.Load($f.FullName,[CrystalDecisions.Shared.OpenReportMethod]::OpenReportByTempCopy)
        $rcd = $doc.ReportClientDocument
        $ctrls = @($rcd.DatabaseController)
        try { foreach ($n in $rcd.SubreportController.GetSubreportNames()) { $ctrls += $rcd.SubreportController.GetSubreport($n).DatabaseController } } catch {}
        foreach ($dbc in $ctrls) {
            $n = 0; try { $n = [int]$dbc.Database.Tables.Count } catch {}
            for ($i = 0; $i -lt $n; $i++) {
                $t = $dbc.Database.Tables[$i]
                $al = [string]$t.Alias
                if ($AliasRx -and ($al -notmatch $AliasRx)) { continue }
                $sql = ''; try { $sql = [string]$t.CommandText } catch {}
                if (-not $sql) { continue }
                $ln = $sql -split "`r?`n"
                for ($k = 0; $k -lt $ln.Count; $k++) {
                    if ($ln[$k] -notmatch '^\s*GROUP\s+BY') { continue }
                    $buf = @($ln[$k])
                    for ($m = $k+1; $m -lt $ln.Count; $m++) {
                        if ($ln[$m] -match '^\s*(UNION|ORDER\s+BY|HAVING|SELECT|FROM|WHERE|\))') { break }
                        if ($ln[$m].Trim() -eq '') { continue }
                        $buf += $ln[$m]
                    }
                    $seen++
                    $clause = ($buf -join ' ') -replace '\s+',' '
                    $isBad  = $clause -match ("AS\s+" + [regex]::Escape($AliasToken))
                    if ($isBad) { $bad++ }
                    if ($isBad -or -not $Quiet) {
                        $short = $clause
                        if ($short.Length -gt 150) { $short = $short.Substring(0,150) + ' ...' }
                        $tag = if ($isBad) { 'BAD  ' } else { 'ok   ' }
                        Write-Host ("{0}{1} [{2}] L{3}: {4}" -f $tag,$f.Name,$al,($k+1),$short)
                    }
                }
            }
        }
    } catch { Write-Host "LOAD-ERR $($f.Name): $($_.Exception.Message)" -ForegroundColor Red }
    finally { try { $doc.Close(); $doc.Dispose() } catch {} }
}
Write-Host ''
Write-Host ("GROUPBY-CLAUSES={0}  WITH-ALIAS-{1}={2}" -f $seen,$AliasToken.ToUpper(),$bad)
exit ([int]([bool]$bad))
