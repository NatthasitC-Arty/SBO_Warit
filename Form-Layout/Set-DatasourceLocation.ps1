<#
.SYNOPSIS
  Bulk Set Datasource Location for Crystal Report (.rpt) files.

.DESCRIPTION
  Scans a folder (or single .rpt), loads each report via Crystal Reports .NET SDK,
  and rewrites ServerName / DatabaseName / Logon credentials on every Table --
  including all subreports -- then saves the report back.

  CR Runtime on this machine is installed in GAC_64, so the script runs as 64-bit
  PowerShell. Falls back to relaunch under SysWOW64 if only GAC_32 is present.

.PARAMETER Path
  Folder to scan recursively, OR a single .rpt path.

.PARAMETER NewServer
  New SQL Server / data source name (required).

.PARAMETER OldServer
  Optional. If set, only tables whose current ServerName matches (case-insensitive)
  are rewritten. Leave empty to rewrite every table.

.PARAMETER NewDatabase
  Optional. New database name. Leave empty to keep existing per-table DB.

.PARAMETER NewUser
  Optional. SQL login. Leave empty to keep existing UserID.

.PARAMETER NewPassword
  Optional. SQL password (not persisted into .rpt by Crystal; only used during
  the in-memory ApplyLogOnInfo).

.PARAMETER Filter
  File mask. Default *.rpt

.PARAMETER NoRecurse
  Do not recurse into subfolders.

.PARAMETER BackupSuffix
  If set, original .rpt is copied to "<name>.rpt<suffix>" before save.

.PARAMETER WhatIf
  Show what would change without saving.

.PARAMETER LogFile
  Log file path. Default <script-dir>\_SetLocation.log
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory)] [string] $Path,
    [Parameter(Mandatory)] [string] $NewServer,
    [string] $OldServer       = '',
    [string] $NewDatabase     = '',
    [string] $NewUser         = '',
    [string] $NewPassword     = '',
    [string] $Filter          = '*.rpt',
    [switch] $NoRecurse,
    [string] $BackupSuffix    = '',
    [string] $LogFile         = ''
)

if (-not $LogFile) {
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $PSCommandPath }
    $LogFile = Join-Path $scriptDir '_SetLocation.log'
}

# --- Bitness check: CR Runtime native libs (ReportAppServer) live in GAC_32 or GAC_64
$arch = if ([Environment]::Is64BitProcess) { 'x64' } else { 'x86' }
Write-Host "[arch] running as $arch PowerShell" -ForegroundColor DarkGray
$gacRas32 = Test-Path 'C:\Windows\Microsoft.NET\assembly\GAC_32\CrystalDecisions.ReportAppServer.CommLayer'
$gacRas64 = Test-Path 'C:\Windows\Microsoft.NET\assembly\GAC_64\CrystalDecisions.ReportAppServer.CommLayer'
if (-not $gacRas32 -and -not $gacRas64) {
    Write-Error "Crystal Reports Runtime for .NET 4.0 not found (no ReportAppServer.CommLayer in GAC_32/GAC_64). Install CRRuntime_*_13_0_xx.msi from SAP."
    exit 2
}
if ([Environment]::Is64BitProcess -and -not $gacRas64 -and $gacRas32) {
    $ps86 = "$env:WINDIR\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    Write-Host "[relaunch] CR runtime is 32-bit only - switching to $ps86" -ForegroundColor Yellow
    $argList = @('-NoProfile','-ExecutionPolicy','Bypass','-File', $PSCommandPath)
    foreach ($k in $PSBoundParameters.Keys) {
        $v = $PSBoundParameters[$k]
        if ($v -is [switch]) { if ($v.IsPresent) { $argList += "-$k" } }
        else                 { $argList += "-$k"; $argList += [string]$v }
    }
    & $ps86 @argList; exit $LASTEXITCODE
}
if (-not [Environment]::Is64BitProcess -and -not $gacRas32 -and $gacRas64) {
    Write-Error "Running 32-bit but CR Runtime is 64-bit only. Re-run from 64-bit PowerShell, or install CRRuntime_32bit_13_0_xx.msi."
    exit 2
}

# --- Locate & load Crystal Reports SDK from GAC ---
function Load-CrystalSDK {
    $gacRoot = 'C:\Windows\Microsoft.NET\assembly\GAC_MSIL'
    $names = 'CrystalDecisions.Shared','CrystalDecisions.ReportSource','CrystalDecisions.CrystalReports.Engine'
    foreach ($n in $names) {
        $dir = Join-Path $gacRoot $n
        if (-not (Test-Path $dir)) { throw "GAC folder not found: $dir" }
        $dll = Get-ChildItem $dir -Recurse -Filter "$n.dll" |
               Sort-Object { [version](($_.Directory.Name -split '_')[1]) } -Descending |
               Select-Object -First 1
        if (-not $dll) { throw "DLL not found under $dir" }
        [void][System.Reflection.Assembly]::LoadFrom($dll.FullName)
        Write-Host "[sdk] loaded $($dll.FullName)" -ForegroundColor DarkGray
    }
}
try { Load-CrystalSDK } catch { Write-Error "SDK load failed: $($_.Exception.Message)"; exit 3 }

# --- Logging ---
$logDir = Split-Path $LogFile -Parent
if ($logDir -and -not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force -WhatIf:$false -Confirm:$false | Out-Null
}
function Log {
    param([string]$Level,[string]$Msg)
    $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$ts][$Level] $Msg"
    switch ($Level) {
        'ERR'  { Write-Host $line -ForegroundColor Red }
        'WARN' { Write-Host $line -ForegroundColor Yellow }
        'OK'   { Write-Host $line -ForegroundColor Green }
        default{ Write-Host $line }
    }
    try { [System.IO.File]::AppendAllText($LogFile, $line + "`r`n", [System.Text.Encoding]::UTF8) }
    catch { Write-Host "  (log write failed: $($_.Exception.Message))" -ForegroundColor DarkRed }
}

Log 'INFO' "=== Set-DatasourceLocation start ==="
Log 'INFO' "Path=$Path  NewServer=$NewServer  OldServer=$OldServer  NewDatabase=$NewDatabase  NewUser=$NewUser  WhatIf=$($WhatIfPreference)"

# --- Enumerate .rpt files ---
$rpts = @()
if (Test-Path -LiteralPath $Path -PathType Leaf) {
    if ($Path -like '*.rpt') { $rpts = ,(Get-Item -LiteralPath $Path) }
    else { Log 'ERR' "Path is a file but not .rpt: $Path"; exit 4 }
} elseif (Test-Path -LiteralPath $Path -PathType Container) {
    $gciArgs = @{ LiteralPath = $Path; Filter = $Filter; File = $true }
    if (-not $NoRecurse) { $gciArgs['Recurse'] = $true }
    $rpts = Get-ChildItem @gciArgs
} else {
    Log 'ERR' "Path not found: $Path"; exit 4
}
Log 'INFO' "Found $($rpts.Count) report file(s)."

# --- Per-table rewrite ---
function Set-TableLocation {
    param($Table, [string]$Scope)
    $li = $Table.LogOnInfo
    $ci = $li.ConnectionInfo
    $curServer = $ci.ServerName
    $curDb     = $ci.DatabaseName

    if ($OldServer -and ($curServer -ne $OldServer)) {
        Log 'INFO' "  [$Scope] skip table '$($Table.Name)' (server='$curServer' != OldServer)"
        return $false
    }

    $ci.ServerName = $NewServer
    if ($NewDatabase) { $ci.DatabaseName = $NewDatabase }
    if ($NewUser)     { $ci.UserID       = $NewUser }
    if ($NewPassword) { $ci.Password     = $NewPassword }
    $Table.ApplyLogOnInfo($li)

    try {
        $loc = $Table.Location
        if ($NewDatabase -and $loc) {
            $parts = $loc -split '\.'
            if ($parts.Count -ge 3) {
                $tableLeaf = $parts[-1]; $schema = $parts[-2]
                $Table.Location = "$NewDatabase.$schema.$tableLeaf"
            }
        }
    } catch { Log 'WARN' "  [$Scope] set Location failed on '$($Table.Name)': $($_.Exception.Message)" }

    Log 'OK'   "  [$Scope] '$($Table.Name)' server '$curServer' -> '$NewServer'$(if($NewDatabase){' db '+$curDb+' -> '+$NewDatabase})"
    return $true
}

# --- Process each report ---
$okCount = 0; $errCount = 0; $skipCount = 0
foreach ($f in $rpts) {
    Log 'INFO' "--> $($f.FullName)"
    $doc = New-Object CrystalDecisions.CrystalReports.Engine.ReportDocument
    try {
        $doc.Load($f.FullName, [CrystalDecisions.Shared.OpenReportMethod]::OpenReportByTempCopy)
        $changed = 0
        foreach ($t in $doc.Database.Tables) {
            if (Set-TableLocation -Table $t -Scope 'main') { $changed++ }
        }
        foreach ($sub in $doc.Subreports) {
            foreach ($t in $sub.Database.Tables) {
                if (Set-TableLocation -Table $t -Scope "sub:$($sub.Name)") { $changed++ }
            }
        }
        if ($changed -eq 0) {
            Log 'WARN' "  no tables matched -- nothing to save"
            $skipCount++
        } elseif ($PSCmdlet.ShouldProcess($f.FullName, "SaveAs $changed table(s) changed")) {
            if ($BackupSuffix) {
                $bak = "$($f.FullName)$BackupSuffix"
                Copy-Item -LiteralPath $f.FullName -Destination $bak -Force
                Log 'INFO' "  backup -> $bak"
            }
            $doc.SaveAs($f.FullName)
            Log 'OK'   "  saved $changed table(s)"
            $okCount++
        } else {
            Log 'INFO' "  WhatIf: would save $changed table change(s) [no save]"
        }
    } catch {
        Log 'ERR' "  $($f.Name): $($_.Exception.Message)"
        $errCount++
    } finally {
        try { $doc.Close(); $doc.Dispose() } catch {}
    }
}

Log 'INFO' "=== Done. ok=$okCount  err=$errCount  skipped=$skipCount  total=$($rpts.Count) ==="
exit ([int]([bool]$errCount))
