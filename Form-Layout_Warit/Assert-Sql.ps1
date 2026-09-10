<#
.SYNOPSIS
  Counts SQL patterns across every Command table of every .rpt, straight off disk.

.DESCRIPTION
  An INDEPENDENT check. It deliberately does not share code with
  Apply-SqlFixes.ps1: if a later transform silently undid an earlier one, or a
  write never reached the file, this is what catches it.

  Typical use -- take a baseline before touching anything, then compare after:
      Assert-Sql.ps1 -Root <copy> -Save before.json
      Apply-SqlFixes.ps1 -Root <copy> -Only T4
      Assert-Sql.ps1 -Root <copy> -Compare before.json -MustBeZero nvmax

  Counters:
      files, tables      how much was scanned
      nvmax              NVARCHAR(MAX)
      ougp               OUGP.UgpCode still present
      pjUnbound          "LEFT JOIN <t> pj ON <hdr>.DocEntry = <line>.DocEntry"
                         where the ON clause never names pj -- an unconstrained
                         self-join that multiplies rows
      dupAlias           a text branch that INNER JOINs <T>10 and then LEFT JOINs
                         it again ("correlation name specified multiple times")
      linenum            RDN1.LineNum used where VisOrder is meant
      gbAlias            a GROUP BY clause containing "AS UgpCode" -- a syntax
                         error. The clause is bounded BY LINES: a regex that runs
                         forward to the next UNION/ORDER BY sails past a
                         subquery's GROUP BY into the outer SELECT list and
                         reports a false positive on every aliased column there.
      asUgpCode, asUomCode, joinOuom, uomForeign
                         end-state counts for the UoM work. These are NOT
                         expected to be zero and they are NOT zero to begin with
                         -- always read them against a -Compare baseline.

.PARAMETER Root
  Folder to scan recursively, or a single .rpt.

.PARAMETER Save
  Write the counters to this JSON file.

.PARAMETER Compare
  Read a previously saved JSON and print before -> after for every counter.

.PARAMETER MustBeZero
  Comma list of counter names that must be 0, e.g. -MustBeZero 'nvmax,gbAlias'.
  Exit code is 1 if any of them is not 0. Nothing fails by default -- what counts
  as a failure depends on the job.

.PARAMETER Detail
  Also list the file/alias behind ougp, pjUnbound and gbAlias.
#>
param(
    [Parameter(Mandatory)] [string] $Root,
    [string]   $Save       = '',
    [string]   $Compare    = '',
    [string[]] $MustBeZero = @(),
    [switch]   $Detail
)
$ErrorActionPreference = 'Stop'
# powershell.exe -File does not split a comma list into an array.
$MustBeZero = @($MustBeZero | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })

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

$c = [ordered]@{
    files=0; tables=0; nvmax=0; ougp=0; pjUnbound=0; dupAlias=0; linenum=0
    gbAlias=0; asUgpCode=0; asUomCode=0; joinOuom=0; uomForeign=0
}
$whereOugp = @(); $wherePj = @(); $whereGb = @()

$P_DUP = '(AND (?:RDN10|IGN10)\.AftLineNum = (?:RDN1|IGN1)\.VisOrder\r?\n[^\r\n]*\bpj\b[^\r\n]*\r?\n)[ \t]*LEFT JOIN (?:DRF10 )?(?:RDN10|IGN10) ON'
$P_PJ  = 'LEFT\s+JOIN\s+\w+\s+pj\s+ON\s+\[?\w+\]?\.\[?DocEntry\]?\s*=\s*\[?(\w+)\]?\.\[?DocEntry\]?'

$rpts = @()
if (Test-Path -LiteralPath $Root -PathType Leaf) { $rpts = ,(Get-Item -LiteralPath $Root) }
else { $rpts = Get-ChildItem -LiteralPath $Root -Filter *.rpt -File -Recurse }

foreach ($f in $rpts) {
    $c.files++
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
                $sql = ''; try { $sql = [string]$t.CommandText } catch {}
                if (-not $sql) { continue }
                $c.tables++
                $c.nvmax      += ([regex]::Matches($sql,'NVARCHAR\s*\(\s*MAX\s*\)','IgnoreCase')).Count
                $c.dupAlias   += ([regex]::Matches($sql,$P_DUP,'IgnoreCase')).Count
                $c.linenum    += ([regex]::Matches($sql,'RDN1\.LineNum = RDN10\.AftLineNum','IgnoreCase')).Count
                $c.asUgpCode  += ([regex]::Matches($sql,'AS\s+UgpCode','IgnoreCase')).Count
                $c.asUomCode  += ([regex]::Matches($sql,"AS\s+'UomCode'",'IgnoreCase')).Count
                $c.joinOuom   += ([regex]::Matches($sql,'LEFT\s+JOIN\s+OUOM\s+ON','IgnoreCase')).Count
                $c.uomForeign += ([regex]::Matches($sql,'U_SLD_Uomforeign','IgnoreCase')).Count

                $o = ([regex]::Matches($sql,'OUGP\.UgpCode','IgnoreCase')).Count
                if ($o) { $c.ougp += $o; $whereOugp += ("{0} [{1}] x{2}" -f $f.Name,$t.Alias,$o) }

                foreach ($m in [regex]::Matches($sql,$P_PJ,'IgnoreCase')) {
                    if ($m.Groups[1].Value -ne 'pj') {
                        $c.pjUnbound++
                        $wherePj += ("{0} [{1}] {2}" -f $f.Name,$t.Alias,($m.Value -replace '\s+',' '))
                    }
                }

                # GROUP BY bounded by lines -- see the header for why.
                $ln = $sql -split "`r?`n"
                for ($k = 0; $k -lt $ln.Count; $k++) {
                    if ($ln[$k] -notmatch '^\s*GROUP\s+BY') { continue }
                    $buf = @($ln[$k])
                    for ($m2 = $k+1; $m2 -lt $ln.Count; $m2++) {
                        if ($ln[$m2] -match '^\s*(UNION|ORDER\s+BY|HAVING|SELECT|FROM|WHERE|\))') { break }
                        $buf += $ln[$m2]
                    }
                    if (($buf -join ' ') -match 'AS\s+UgpCode') {
                        $c.gbAlias++
                        $whereGb += ("{0} [{1}] L{2}" -f $f.Name,$t.Alias,($k+1))
                    }
                }
            }
        }
    } catch { Write-Host "LOAD-ERR $($f.Name): $($_.Exception.Message)" -ForegroundColor Red }
    finally { try { $doc.Close(); $doc.Dispose() } catch {} }
}

$base = $null
if ($Compare) {
    if (Test-Path -LiteralPath $Compare) { $base = Get-Content -Raw -LiteralPath $Compare | ConvertFrom-Json }
    else { Write-Host "[WARN] baseline not found: $Compare" -ForegroundColor Yellow }
}

Write-Host ''
foreach ($k in $c.Keys) {
    if ($base -and ($base.PSObject.Properties.Name -contains $k)) {
        $b = [int]$base.$k; $d = [int]$c[$k] - $b
        $tag = if ($d -eq 0) { 'same' } elseif ($d -gt 0) { "+$d" } else { "$d" }
        Write-Host ("  {0,-12} {1,6}  (was {2}, {3})" -f $k,$c[$k],$b,$tag)
    } else {
        Write-Host ("  {0,-12} {1,6}" -f $k,$c[$k])
    }
}
if ($Detail) {
    if ($whereOugp) { Write-Host ''; Write-Host '  OUGP.UgpCode:'; $whereOugp | ForEach-Object { Write-Host "    $_" } }
    if ($whereGb)   { Write-Host ''; Write-Host '  GROUP BY with alias:'; $whereGb | ForEach-Object { Write-Host "    $_" } }
    if ($wherePj)   { Write-Host ''; Write-Host ("  pj unbound (first 10 of {0}):" -f $wherePj.Count); $wherePj | Select-Object -First 10 | ForEach-Object { Write-Host "    $_" } }
}

if ($Save) {
    ($c | ConvertTo-Json) | Set-Content -LiteralPath $Save -Encoding UTF8
    Write-Host ''
    Write-Host "  baseline saved -> $Save"
}

$fail = 0
foreach ($k in $MustBeZero) {
    if (-not $c.Contains($k)) { Write-Host "[WARN] unknown counter '$k'" -ForegroundColor Yellow; continue }
    if ([int]$c[$k] -ne 0) { Write-Host ("[FAIL] {0} = {1}, expected 0" -f $k,$c[$k]) -ForegroundColor Red; $fail++ }
}
Write-Host ''
Write-Host ('ASSERT|' + (($c.Keys | ForEach-Object { "$_=$($c[$_])" }) -join '|') + "|FAIL=$fail")
exit ([int]([bool]$fail))
