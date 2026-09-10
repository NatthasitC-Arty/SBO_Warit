<#
.SYNOPSIS
  Bulk "Set Datasource Location" for Crystal Report (.rpt) files -- RAS edition.

.DESCRIPTION
  Scans a folder (or a single .rpt), opens each report through the Crystal
  Reports .NET SDK / ReportAppServer (RAS), and rewrites the connection of every
  table -- in the main report AND in every subreport -- then saves the report and
  RE-OPENS it from disk to prove the change actually landed.

  ApplyLogOnInfo DOES NOT MOVE A TABLE OFFLINE -- measured 2026-08-31 on copies,
  always read back FROM DISK. With a stored UserID present on BOTH the engine and
  the RAS layer, ApplyLogOnInfo still left the saved file on the OLD server and
  database (B25.rpt: logonApplied=3, changed=0, rolled back; two Return reports:
  logonApplied=6, changed=0, rolled back). The same moves through
  SetTableLocation succeeded with err=0 and every UserID intact. It does not
  throw and reading it back in memory looks right, which is exactly how the old
  "82 of 207 tables failed" note and the later "4 of 4 tables took it" note were
  BOTH produced. Judge these calls only by the file on disk.
  ApplyLogOnInfo is therefore an engine-layer helper, never the write.

  APIs proven on this repo -- do not substitute:
    OK    SetTableLocation(currentTable, modifiedClone) + doc.SaveAs()
    BAD   ApplyLogOnInfo -- silently persists nothing offline, any UserID state
    BAD   SetTableLocationByServerDatabaseName -- wipes server/database
    BAD   ModifyTableConnectionInfo            -- memory only, never persisted
    BAD   ReplaceConnection                    -- never persists, blank UserName
                                                  or not

  THE User ID RULE -- measured 2026-08-31, always read back FROM DISK.
  A .rpt whose UserID is blank CANNOT BE OPENED BY SAP B1, so losing it is real
  data loss. UserID does NOT live where it looks like it lives:
    clone.ConnectionInfo.UserName = 'sa'      -> in memory only. The saved file
                                                 has ''. This is the trap the
                                                 fallback path used to fall in.
    SetTableLocation(live, clone)             -> WIPES UserID, even when the
                                                 clone is UNMODIFIED, unless the
                                                 line below is done first.
    ApplyLogOnInfo / ModifyTableConnectionInfo / DataSourceConnections.SetLogon /
    SetConnection / SetDatabaseLogon / SetTableLocationByServerDatabaseName /
    ReplaceConnection                         -> none of them can put a UserID
                                                 back once it reads as ''.
    Set-BagString -Bag <QE_LogonProperties> -Key 'User ID' -Value <user>
      then SetTableLocation(live, clone)      -> UserID SURVIVES, and is CREATED
                                                 on a table that had none.
  The nested QE_LogonProperties bag is the only writable door. Crystal consumes
  the 'User ID' entry while saving, so the bag on disk reads exactly as before.
  Verified on C:\_bakchk\B25.rpt (3 tables, UserID='sa') and on a repo report
  whose UserID was blank (4 tables): server, database, UserID and field counts
  all correct on reload. Writing the bag in place is enough -- do NOT assign the
  bag object back through the indexer (see rule 3 below).
  This is what -NewUser now does, and it is why -NewUser finally works on the
  reports whose UserID was already empty.

  Every table is READ BACK from the freshly re-opened file and compared with the
  intended values. A file whose verification fails is RESTORED from a pre-save
  temp copy, so a bad write never survives.

  All tables in this repo are SQL Commands (Table.Name = 'Command'), so the old
  "Table.Location = db.schema.table" block was removed -- a command table's
  Location is always the literal string 'Command' and writing it does nothing.

.PARAMETER Path
  Folder to scan recursively, OR a single .rpt path.

.PARAMETER NewServer
  Target SQL Server / data source name (required; also the comparison target in
  -Verify mode).

.PARAMETER OldServer
  Optional. Only tables whose current QE_ServerDescription matches are rewritten.
  Blank = rewrite every table.

.PARAMETER NewDatabase
  Optional. Target database name. Blank = keep the existing per-table database.

.PARAMETER NewUser
  Optional SQL login. Written into the nested QE_LogonProperties bag as
  'User ID', which is the only place that survives SaveAs. Works even when the
  table's current UserID is blank, so this is also the way to REPAIR a report
  that SAP B1 refuses to open. Tables that already carry a UserID keep it when
  -NewUser is not supplied.

.PARAMETER NewPassword
  Optional SQL password written into the cloned ConnectionInfo.

.PARAMETER Filter
  File mask. Default *.rpt

.PARAMETER NoRecurse
  Do not recurse into subfolders.

.PARAMETER Verify
  READ-ONLY. Opens every report and lists each table whose QE_ServerDescription /
  QE_DatabaseName does not match -NewServer / -NewDatabase (MISMATCH), and each
  table whose PreQEServerName / PreQEDatabaseName still points elsewhere (STALE).
  Nothing is written.

.PARAMETER BackupSuffix
  If set, the original .rpt is copied to "<name>.rpt<suffix>" before saving.

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
    [switch] $NoDesignerNormalize,
    [switch] $Connect,
    [ValidateRange(1,5)] [int] $ConnectMode = 1,
    [switch] $AllowFieldChange,
    [ValidateSet('Auto','ApplyLogOnInfo','SetTableLocation')] [string] $Method = 'Auto',
    [string] $ConnectionFrom = '',
    [switch] $ListConnections,
    [switch] $Force,
    [string] $Filter          = '*.rpt',
    [switch] $NoRecurse,
    [switch] $Verify,
    [string] $BackupSuffix    = '',
    [string] $LogFile         = ''
)

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $PSCommandPath }
if (-not $LogFile) { $LogFile = Join-Path $scriptDir '_SetLocation.log' }

# ------------------------------------------------------------------ bitness
# The ReportAppServer assemblies on this machine exist ONLY in GAC_64.
if (-not [Environment]::Is64BitProcess) {
    Write-Host '[ERROR] Crystal RAS assemblies live in GAC_64. Re-run under 64-bit powershell.exe' -ForegroundColor Red
    exit 2
}
if (-not (Test-Path 'C:\Windows\Microsoft.NET\assembly\GAC_64\CrystalDecisions.ReportAppServer.CommLayer')) {
    Write-Host '[ERROR] Crystal Reports Runtime for .NET 4.0 not found in GAC_64. Install CRRuntime_64bit_13_0_xx.msi from SAP.' -ForegroundColor Red
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

# ------------------------------------------------------------------ logging
$logDir = Split-Path $LogFile -Parent
if ($logDir -and -not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force -WhatIf:$false -Confirm:$false | Out-Null
}
function Log {
    param([string]$Level,[string]$Msg)
    $ts   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = "[$ts][$Level] $Msg"
    switch ($Level) {
        'ERR'  { Write-Host $line -ForegroundColor Red }
        'WARN' { Write-Host $line -ForegroundColor Yellow }
        'OK'   { Write-Host $line -ForegroundColor Green }
        'HIT'  { Write-Host $line -ForegroundColor Cyan }
        default{ Write-Host $line }
    }
    try { [System.IO.File]::AppendAllText($LogFile, $line + "`r`n", [System.Text.Encoding]::UTF8) } catch {}
}


# ------------------------------------------------------------------ property bag
# HARD-WON RULES about the RAS PropertyBag -- all three were proven on this repo
# and every one of them caused a SILENT wrong result before it was understood:
#
# 1. Objects reached through ReportDocument.ReportClientDocument are
#    System.__ComObject (late-bound IDispatch), NOT managed CrystalDecisions
#    types. ConnectionInfo.ServerName / .DatabaseName always read back EMPTY and
#    Attributes has no .Collection -- the string indexer is the only door.
#
# 2. The real server/database live in TWO places. Writing only the outer pair
#    leaves the report pointing at the old machine:
#      outer  QE_ServerDescription / QE_DatabaseName
#      inner  QE_LogonProperties -> "Data Source" / "Initial Catalog"
#                                   "PreQEServerName" / "PreQEDatabaseName"
#    PreQE* are NESTED, not top level. Looking for them at the top level finds
#    nothing and makes every report look clean.
#
# 3. NEVER assign an object through the PowerShell indexer, i.e. never
#      $attributes['QE_LogonProperties'] = $bag
#    Doing it once makes PowerShell cache the Item setter signature as
#    __ComObject, and EVERY later string write on ANY bag throws
#    "Unable to cast System.String to System.__ComObject". Wrapped in try/catch
#    that turns into a silent no-op -- exactly the failure mode this rewrite
#    exists to kill. The nested bag is a live reference anyway, so mutating it
#    in place is enough; it never needs writing back.
#
# Every access below goes through InvokeMember, which bypasses the PowerShell
# COM adapter and its signature cache entirely.
# THE catch{} CONVENTION -- added 2026-08-31 after a bare `catch {}` plus a
# null-tolerant `[int]$null` hid a whole unmoved subreport on every Return
# report while the run still printed OK.
# A catch block in this script may be silent ONLY when the exception carries no
# information the operator could act on:
#   Get-BagValue / Test-BagKey        a missing key IS the answer ($null/$false)
#   Get-ExDetail HResult formatting   best-effort decoration of an error already
#                                     being reported
#   $doc.Close()/.Dispose() in finally  the work is done; a dispose warning per
#                                     file would bury the real output
# EVERYTHING else -- anything that decides whether a table gets written, or that
# turns a failure into a default value -- must Log at least WARN, and must name
# the scope and alias so the line is actionable.
$script:BF = [System.Reflection.BindingFlags]

function Get-BagValue {
    param($Bag, [string]$Key)
    if ($null -eq $Bag) { return $null }
    try { return $Bag.GetType().InvokeMember('Item', $script:BF::GetProperty, $null, $Bag, @([object]$Key)) }
    catch { return $null }
}
function Get-BagString {
    param($Bag, [string]$Key)
    $v = Get-BagValue -Bag $Bag -Key $Key
    if ($null -eq $v) { return '' }
    return [string]$v
}
function Test-BagKey {
    param($Bag, [string]$Key)
    if ($null -eq $Bag) { return $false }
    try { return [bool]$Bag.GetType().InvokeMember('Contains', $script:BF::InvokeMethod, $null, $Bag, @([object]$Key)) }
    catch { return $false }
}
# Throws on failure ON PURPOSE -- a swallowed write is how this script used to lie.
function Set-BagString {
    param($Bag, [string]$Key, [string]$Value)
    $Bag.GetType().InvokeMember('Item', $script:BF::SetProperty, $null, $Bag, @([object]$Key, [object]$Value))
}

# THE one line that keeps a UserID alive across SetTableLocation. Measured
# 2026-08-31: ConnectionInfo.UserName is a floating value that never reaches the
# file; the nested QE_LogonProperties key 'User ID' is what Crystal actually
# stores. Writing it in place on the CLONE's bag, before SetTableLocation, both
# PRESERVES an existing UserID and CREATES one on a table that had none.
# All four ways of writing it were tried (Item-set / Add, with and without
# assigning the bag back) and all four persist, so use the in-place write and
# never hand the bag object back through the indexer -- see rule 3 above.
function Set-BagUser {
    param($LogonBag, [string]$User)
    if ($null -eq $LogonBag -or -not $User) { return $false }
    try { Set-BagString -Bag $LogonBag -Key 'User ID' -Value $User; return $true }
    catch { return $false }
}

# QE_SQLDB is a BOOLEAN, not a string. Writing 'True' as text leaves Crystal
# treating the source as non-SQL. InvokeMember takes the real [bool].
function Set-BagBool {
    param($Bag, [string]$Key, [bool]$Value)
    $Bag.GetType().InvokeMember('Item', $script:BF::SetProperty, $null, $Bag, @([object]$Key, [object]$Value))
}
# The RAS property bag exposes Contains + Remove as METHODS (no Delete/Clear).
# Returns $true only when a key was actually there and is now gone.
function Remove-BagKey {
    param($Bag, [string]$Key)
    if ($null -eq $Bag) { return $false }
    if (-not (Test-BagKey -Bag $Bag -Key $Key)) { return $false }
    try {
        [void]$Bag.GetType().InvokeMember('Remove', $script:BF::InvokeMethod, $null, $Bag, @([object]$Key))
        return (-not (Test-BagKey -Bag $Bag -Key $Key))
    } catch { return $false }
}

# Everything that decides whether a table points at the right place, outer and inner.
function Read-TableConn {
    param($Table)
    $at = $null
    # A table with no Attributes bag cannot be read OR written -- it must not
    # look like a table that is simply on the wrong server.
    try { $at = $Table.ConnectionInfo.Attributes }
    catch { Log 'WARN' "    cannot read ConnectionInfo.Attributes on '$($Table.Alias)': $($_.Exception.Message.Split([char]10)[0])" }
    $lp = Get-BagValue -Bag $at -Key 'QE_LogonProperties'
    $usr = ''
    try { $ci0 = $Table.ConnectionInfo
          $usr = [string]$ci0.GetType().InvokeMember('UserName', $script:BF::GetProperty, $null, $ci0, $null) }
    catch { Log 'WARN' "    cannot read UserName on '$($Table.Alias)' -- treated as blank: $($_.Exception.Message.Split([char]10)[0])" }
    [pscustomobject]@{
        Server = (Get-BagString -Bag $at -Key 'QE_ServerDescription')
        Db     = (Get-BagString -Bag $at -Key 'QE_DatabaseName')
        DS     = (Get-BagString -Bag $lp -Key 'Data Source')
        IC     = (Get-BagString -Bag $lp -Key 'Initial Catalog')
        PreSrv = (Get-BagString -Bag $lp -Key 'PreQEServerName')
        PreDb  = (Get-BagString -Bag $lp -Key 'PreQEDatabaseName')
        SqlDb  = (Get-BagString -Bag $at -Key 'QE_SQLDB')
        User   = $usr
        # 'Database DLL' has one legitimate home: the OUTER Attributes bag.
        # When it also appears INSIDE QE_LogonProperties it is poison -- see the
        # DLL-IN-LOGON-BAG note near the verify block. True = poisoned.
        DllInLp = (Test-BagKey -Bag $lp -Key 'Database DLL')
    }
}

# Every (scope, DatabaseController) pair of a loaded report: main + one per subreport.
#
# WHY THIS IS SO DEFENSIVE -- the bug Ars caught in Crystal Designer 2026-08-31.
# Every report in "2. Sales - AR\4. Return" moved its 3 main tables to the new
# server and left its subreport sitting on 192.168.0.216. The script reported
# tables=6 for the pair, i.e. 3 per file, i.e. THE SUBREPORT SCOPE WAS NEVER
# ENUMERATED -- and it said nothing, because the whole lookup was wrapped in a
# bare `catch {}`.
#
# Two independent doors are used now and the wider one wins:
# THE ROOT CAUSE, measured 2026-08-31 by logging every step of the old code:
#   GetSubreportNames() was FINE  -- it returned 'Batch/Serial' every time.
#   The foreach body DID run       -- the scope list came back with 2 entries.
#   The scope's .Dbc was NULL      -- and the caller's
#       try { $cnt = [int]$dbc.Database.Tables.Count } catch { $cnt = 0 }
#     does NOT throw on a null $dbc under non-strict mode: $null.Database is
#     $null, [int]$null is 0. Zero tables, zero iterations, ZERO log lines.
#     The subreport was skipped in complete silence and the run still said "OK".
#
# WHY .Dbc came back null -- the one API rule to remember here:
#     $rcd.SubreportController.GetSubreport($n).DatabaseController   -> NULL
#     $ctl = $rcd.SubreportController                                       #
#     $ctl.GetSubreport($n).DatabaseController                       -> OK
# Calling GetSubreport on a TEMPORARY SubreportController hands back a
# SubreportClientDocument whose DatabaseController is already gone. The
# controller must be held in a variable that outlives the subreport object.
# That is why $sub_ctl is a local here, and why Ctl/Sub are carried on the
# returned descriptor: they keep the COM references rooted for the caller,
# which dereferences .Dbc long after this function has returned.
#
# Two independent doors are used and the wider one wins:
#   RAS     rcd.SubreportController.GetSubreportNames()  -- returns a COM
#           collection (System.__ComObject), NOT a string[]. It is materialised
#           into a real [string[]] IMMEDIATELY, never left as a live COM
#           collection that can go empty later in the run.
#   ENGINE  Doc.Subreports -- the engine-side list. Used as the CROSS-CHECK: it
#           is what Crystal Designer's tree shows, so if RAS hands back fewer
#           names than the engine sees, RAS is the one that is wrong and the
#           engine names are used instead.
# Anything that goes wrong here is LOGGED. A silent subreport is exactly the
# failure that shipped the half-moved Return reports Ars caught in Designer.
function Get-Scopes {
    param($Doc)
    $out = New-Object System.Collections.ArrayList
    $rcd = $Doc.ReportClientDocument
    [void]$out.Add([pscustomobject]@{ Scope = 'main'; Dbc = $rcd.DatabaseController; SubName = '' })

    # ---- engine cross-check: how many subreports does Crystal Designer show?
    $engNames = @()
    try {
        $ec = [int]$Doc.Subreports.Count
        for ($i = 0; $i -lt $ec; $i++) {
            try { $engNames += [string]$Doc.Subreports[$i].Name }
            catch { Log 'WARN' "    engine Subreports[$i].Name unreadable: $($_.Exception.Message.Split([char]10)[0])" }
        }
    } catch {
        Log 'WARN' "    engine Doc.Subreports unreadable: $($_.Exception.Message.Split([char]10)[0])"
    }

    # ---- RAS door
    $rasNames = @()
    $sub_ctl  = $null
    try { $sub_ctl = $rcd.SubreportController }
    catch { Log 'WARN' "    RAS SubreportController unavailable: $($_.Exception.Message.Split([char]10)[0])" }
    if ($null -ne $sub_ctl) {
        try {
            $raw = $sub_ctl.GetSubreportNames()
            if ($null -eq $raw) {
                Log 'WARN' '    RAS GetSubreportNames() returned null'
            } else {
                foreach ($n in $raw) { if ($n) { $rasNames += [string]$n } }
            }
        } catch {
            Log 'WARN' "    RAS GetSubreportNames() failed: $($_.Exception.Message.Split([char]10)[0])"
        }
    }

    # ---- pick the wider list, and say so when they disagree
    $subNames = $rasNames
    if ($engNames.Count -gt $rasNames.Count) {
        Log 'WARN' "    subreport list disagrees: RAS=$($rasNames.Count) engine=$($engNames.Count) -- using the engine names"
        $subNames = $engNames
    }
    if ($subNames.Count -eq 0 -and $engNames.Count -eq 0) { return ,$out }

    foreach ($sn in $subNames) {
        if ($null -eq $sub_ctl) {
            Log 'ERR' "    subreport '$sn' exists but RAS SubreportController is unavailable -- IT WILL NOT BE MOVED"
            continue
        }
        try {
            # $sub_ctl MUST be the held local, never $rcd.SubreportController inline.
            $sub = $sub_ctl.GetSubreport($sn)
            if ($null -eq $sub) { Log 'ERR' "    GetSubreport('$sn') returned null -- SUBREPORT NOT MOVED"; continue }
            $sdbc = $sub.DatabaseController
            if ($null -eq $sdbc) {
                Log 'ERR' "    subreport '$sn' has a null DatabaseController -- SUBREPORT NOT MOVED"
                continue
            }
            # Ctl and Sub are dead weight to the caller and that is the point:
            # they root the COM references so .Dbc is still alive out there.
            [void]$out.Add([pscustomobject]@{
                Scope = "sub:$sn"; Dbc = $sdbc; SubName = [string]$sn; Ctl = $sub_ctl; Sub = $sub })
        } catch {
            Log 'ERR' "    cannot open subreport '$sn' -- IT WILL NOT BE MOVED: $($_.Exception.Message.Split([char]10)[0])"
        }
    }
    return ,$out
}

# How many tables a scope has -- and it is NEVER allowed to answer 0 quietly.
# The old inline `try { [int]$dbc.Database.Tables.Count } catch { 0 }` is what
# hid the whole subreport bug: a null $dbc yields 0 without throwing, so the
# scope was skipped and nothing was written to the log. Every zero and every
# throw now names the scope.
function Get-ScopeTableCount {
    param($Scope, [string]$Where = '')
    $tag = if ($Where) { "$Where " } else { '' }
    if ($null -eq $Scope.Dbc) {
        Log 'ERR' "    ${tag}[$($Scope.Scope)] DatabaseController is NULL -- this scope CANNOT be processed"
        return 0
    }
    $n = -1
    try { $n = [int]$Scope.Dbc.Database.Tables.Count }
    catch {
        Log 'ERR' "    ${tag}[$($Scope.Scope)] Tables.Count failed -- scope SKIPPED: $($_.Exception.Message.Split([char]10)[0])"
        return 0
    }
    if ($n -le 0) {
        Log 'WARN' "    ${tag}[$($Scope.Scope)] reports 0 table(s) -- nothing to do for this scope"
        return 0
    }
    return $n
}

# ------------------------------------------------------------------ write methods
# TWO ways to move a table to a new server/database. They are NOT equivalent.
# Measured 2026-08-31 on the report SAP B1 accepts, always read back FROM DISK:
#
#   ApplyLogOnInfo   writes NOTHING to the saved file offline, whatever the
#                    UserID state -- verified from disk on three reports. It only
#                    updates the in-memory engine logon, which -Connect uses.
#                    Never let its result decide whether the RAS write runs.
#   SetTableLocation server/db persist. UserID is WIPED to blank UNLESS
#                    'User ID' is written into the nested QE_LogonProperties bag
#                    of the clone first -- see Set-BagUser below. With that one
#                    line it persists, and it also CREATES a UserID on a table
#                    that never had one, which ApplyLogOnInfo cannot do.
#
# Both reset QE_SQLDB to False, and that is not a bug in either call: the flag
# means "I have verified this connection is a SQL database", so pointing the
# table somewhere else legitimately voids it. Only a successful VerifyDatabase()
# can set it again (see -Connect).
#
# WHAT THE OLD "82 of 207 tables failed" NOTE WAS ACTUALLY SEEING. Two separate
# things, neither of them a 40% API defect:
#   a) reading ConnectionInfo back IN MEMORY right after the call shows the OLD
#      values. That readback is what failed, not the write. Never judge one of
#      these calls by an in-memory readback -- only by the file on disk.
#   b) ApplyLogOnInfo really does write nothing, silently, when the table's
#      stored UserID is blank. Those tables are handled by the RAS pass below,
#      which since 2026-08-31 keeps the UserID instead of destroying it.
function Set-TableByLogOnInfo {
    param($EngineTable, [string]$Server, [string]$Database, [string]$User, [string]$Password)
    $li = $EngineTable.LogOnInfo
    $ci = $li.ConnectionInfo
    $ci.ServerName = $Server
    if ($Database) { $ci.DatabaseName = $Database }
    if ($User)     { $ci.UserID       = $User }
    if ($Password) { $ci.Password     = $Password }
    $EngineTable.ApplyLogOnInfo($li)
}

# ------------------------------------------------------------------ credentials
# The password is never logged, never echoed and never put in an error message.
# It lives in one local for the length of a run and the BSTR is zeroed after use.
function Get-PlainPassword {
    param([string]$Supplied, [string]$User)
    if ($Supplied) { return $Supplied }
    $sec = Read-Host -Prompt "SQL password for '$User' (typing is hidden)" -AsSecureString
    if (-not $sec -or $sec.Length -eq 0) { return '' }
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

# ------------------------------------------------------------------ live connect
# WHY THIS EXISTS -- measured 2026-08-31.
# QE_SQLDB is not ordinary metadata. Crystal only stamps QE_SQLDB=True once it
# has really opened the connection -- which is what Designer's Set Datasource
# Location -> Update does. A report left at QE_SQLDB=False is treated as "not a
# SQL datasource", its SQL Commands never run, and SAP B1 refuses to open it.
# So: log on for real, prove it, and only then save.
#
# THE UserID HALF OF THIS NOTE IS OBSOLETE. It used to say a UserID could not be
# set offline either. That was wrong -- the five routes tried back then all
# missed the nested QE_LogonProperties 'User ID' key. Since 2026-08-31 -NewUser
# sets it offline, with no live connection, on tables that had no UserID at all
# (see Set-BagUser). -Connect is still needed for QE_SQLDB, nothing else.
#
# VerifyTableConnectivity returns a BOOLEAN and does not throw -- that is the
# reliable probe. VerifyDatabase() throws "No error." on a bad logon, which is a
# useless message but a fine signal, and it is the call that re-stamps metadata.
function Get-ExDetail {
    param($ErrRec)
    $e = $ErrRec.Exception
    $bits = @()
    $bits += "msg='" + $e.Message.Split([char]10)[0] + "'"
    try { $bits += ("hr=0x{0:X8}" -f $e.HResult) } catch {}
    $inner = $e.InnerException
    $depth = 0
    while ($inner -and $depth -lt 3) {
        $bits += "inner[$depth]='" + $inner.Message.Split([char]10)[0] + "'"
        try { $bits += ("innerHr[$depth]=0x{0:X8}" -f $inner.HResult) } catch {}
        try { if ($null -ne $inner.ErrorCode) { $bits += "errorCode[$depth]=$($inner.ErrorCode)" } } catch {}
        $inner = $inner.InnerException; $depth++
    }
    return ($bits -join '  ')
}

# ConnectMode -- what to do after the logon, before saving. Measured 2026-08-31:
#   LogonEx does NOT throw on bad credentials (it is lazy).
#   VerifyTableConnectivity returns a BOOLEAN and is the only honest gate.
#   VerifyDatabase() re-runs the report SQL. On this repo it throws the useless
#   string "No error." even when the logon was fine -- and it can legitimately
#   fail when the SQL names UDFs/UDTs the target database does not have
#   (Sale Order references U_SLD_Dis_Amount, U_SLD_FullName, U_SLD_LVatBranch,
#   U_SLD_Title and [@SLDT_SET_BRANCH]). Either way it must not block the save:
#   the goal is only to make Crystal stamp QE_SQLDB, and the file read back from
#   disk is what decides success -- never an exception.
#     1 = logon + connectivity, then save            (default)
#     2 = 1 + VerifyDatabase, failures NON-fatal
#     3 = logon BEFORE SetTableLocation (see caller) + connectivity, then save
#     4 = 1 + ReadRecords, failures NON-fatal
#     5 = 1 + VerifyDatabase, failures FATAL         (the old behaviour)
function Invoke-LiveConnect {
    param($Doc, [string]$Server, [string]$Database, [string]$User, [string]$Password, [int]$Mode = 1)
    $rcd = $Doc.ReportClientDocument
    $dbc = $rcd.DatabaseController
    try { $dbc.LogonEx($Server, $Database, $User, $Password) }
    catch { return [pscustomobject]@{ Ok=$false; Stage='LogonEx'; Msg=(Get-ExDetail $_) } }

    # Subreports need their own logon: QE_SQLDB is stamped per DatabaseController,
    # so logging on the main report only leaves every subreport table flagged
    # "not a SQL database", which is one of the things that makes SAP B1 refuse
    # a layout.
    # BEST-EFFORT ON PURPOSE -- this never decides the credential gate below.
    # The gate answers "are these credentials good?" and the main report already
    # answers it, so a subreport quirk must not be able to block a good move.
    # NOT exercised offline: the test rig has no reachable server, which is
    # exactly why this is wrapped and non-fatal.
    foreach ($lsc in (Get-Scopes -Doc $Doc)) {
        if ($lsc.Scope -eq 'main') { continue }
        try {
            $lsc.Dbc.LogonEx($Server, $Database, $User, $Password)
            Log 'INFO' "    live logon OK [$($lsc.Scope)]"
        } catch {
            Log 'WARN' "    live logon failed [$($lsc.Scope)] -- QE_SQLDB may stay False there: $($_.Exception.Message.Split([char]10)[0])"
        }
    }

    # THE gate -- but scored across ALL tables, not "first failure wins".
    # Measured 2026-08-31: with identical connection settings on all four tables
    # (verified: same Data Source, Initial Catalog and UserName at both the engine
    # and the RAS layer) VerifyTableConnectivity returned True for AR_SO and False
    # for Address. LogonEx is lazy, so the first probe can answer from cached state
    # while the next one is the first to actually put a request on the wire.
    # A single False therefore does NOT mean the credentials are wrong.
    # What genuinely proves bad credentials is NONE of them passing.
    $n = 0
    try { $n = [int]$dbc.Database.Tables.Count }
    catch { Log 'WARN' "    connectivity probe: Tables.Count unreadable: $($_.Exception.Message.Split([char]10)[0])" }
    if ($n -le 0) { Log 'WARN' '    connectivity probe: 0 table(s) to probe' }
    $pass = 0; $fail = @()
    for ($i = 0; $i -lt $n; $i++) {
        $tbl = $dbc.Database.Tables[$i]
        $ok = $false
        try { $ok = [bool]$dbc.VerifyTableConnectivity($tbl) } catch { $ok = $false }
        if ($ok) { $pass++ } else { $fail += [string]$tbl.Alias }
    }
    if ($pass -eq 0) {
        return [pscustomobject]@{ Ok=$false; Stage='VerifyTableConnectivity'; Msg="all $n table(s) refused -- server unreachable or credentials rejected" }
    }
    if ($fail.Count) {
        Log 'WARN' "    VerifyTableConnectivity: $pass/$n passed; not confirmed for $($fail -join ', ')."
        Log 'WARN' "    At least one table connected, so the logon itself is good -- continuing."
    }

    if ($Mode -eq 2 -or $Mode -eq 5) {
        try { $rcd.VerifyDatabase(); Log 'INFO' '    VerifyDatabase OK' }
        catch {
            $d = Get-ExDetail $_
            if ($Mode -eq 5) { return [pscustomobject]@{ Ok=$false; Stage='VerifyDatabase'; Msg=$d } }
            Log 'WARN' "    VerifyDatabase threw but connectivity already passed -- continuing. $d"
        }
    }
    if ($Mode -eq 4) {
        try { $Doc.ReadRecords(); Log 'INFO' '    ReadRecords OK' }
        catch { Log 'WARN' ("    ReadRecords threw -- continuing. " + (Get-ExDetail $_)) }
    }
    return [pscustomobject]@{ Ok=$true; Stage=''; Msg='' }
}

# ------------------------------------------------------------------ list mode
# WHY -ConnectionFrom EXISTS -- measured 2026-08-31, always read back FROM DISK.
#   ApplyLogOnInfo only persists on a connection that ALREADY has a stored
#   UserID. Point it at a table whose UserID is blank and it writes nothing at
#   all: no exception, no return value, no warning. That is the real cause of
#   the "82 of 207 tables failed" this script used to blame on the API.
#   Reusing the LogOnInfo of a table that DOES have credentials works, and the
#   borrowed UserID survives on every table (4 of 4 verified).
#   Reusing one from a DIFFERENT .rpt does nothing -- the object is bound to its
#   own document. So the donor has to live in the same report.
#   SINCE 2026-08-31 a donor is no longer the only way out: -NewUser writes the
#   UserID straight into the nested logon bag and works on a report where NO
#   table has credentials, which -ConnectionFrom cannot do. Keep this switch for
#   borrowing a whole connection; use -NewUser when all you need is the login.
# This is Crystal Designer's "My Connections" -> Update, rather than
# "Create New Connection", expressed through the SDK.
# -ListConnections is the tool Ars checks a report with, so it MUST show exactly
# what -Verify and the apply pass see -- which means it walks Get-Scopes, not
# just Doc.Database.Tables.
#
# The old version listed the ENGINE's main tables only. That is why the Return
# reports looked fully moved here while Crystal Designer's own Set Datasource
# Location tree showed the subreport still on 192.168.0.216: the subreport was
# never in this listing to begin with. Scope is printed in front of every row.
function Show-Connections {
    param($Doc, [string]$File)
    Write-Host ''
    Write-Host "=== $File ==="
    Write-Host ("  {0,-22} {1,-3} {2,-22} {3,-18} {4,-24} {5}" -f 'scope','idx','alias','server','database','userID (blank = ApplyLogOnInfo will NOT persist)')
    $rows = 0
    foreach ($sc in (Get-Scopes -Doc $Doc)) {
        $label = if ($sc.Scope -eq 'main') { 'main' } else { "[$($sc.Scope)]" }
        $cnt   = Get-ScopeTableCount -Scope $sc -Where 'list'
        for ($i = 0; $i -lt $cnt; $i++) {
            $t    = $sc.Dbc.Database.Tables[$i]
            $conn = Read-TableConn -Table $t
            $u    = [string]$conn.User
            $flag = if ($u) { $u } else { '<blank>  <-- unusable as -ConnectionFrom donor' }
            # Read-TableConn is the RAS view (QE_ServerDescription / QE_DatabaseName),
            # the same pair -Verify compares, so this listing can never disagree
            # with the verdict the script gives.
            Write-Host ("  {0,-22} {1,-3} {2,-22} {3,-18} {4,-24} {5}" -f $label,$i,[string]$t.Alias,$conn.Server,$conn.Db,$flag)
            $rows++
        }
    }
    if ($rows -eq 0) { Log 'WARN' "    $File : no tables listed in any scope" }
}

# ------------------------------------------------------------------ enumerate
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

$mode = if ($Verify) { 'VERIFY' } elseif ($WhatIfPreference) { 'WHATIF' } else { 'APPLY' }
Log 'INFO' "=== Set-DatasourceLocation start  mode=$mode ==="
$pwShown = if ($NewPassword) { '***' } elseif ($Connect) { '<prompt>' } else { '<none>' }
$cmShown = if ($Connect) { "ConnectMode=$ConnectMode" } else { 'Connect=off' }
# VerifyDatabase / ReadRecords deliberately re-read the schema from the target
# database, so the report's field list is SUPPOSED to change. Measured
# 2026-08-31: the hand-repaired report SAP B1 accepts has 62 fields on AR_SO
# where the broken one has 59 -- i.e. 59 -> 62 is the CORRECT result, and the
# fieldDiff guard was rolling back the very thing we were trying to achieve.
if ($Connect -and ($ConnectMode -eq 2 -or $ConnectMode -eq 4 -or $ConnectMode -eq 5)) {
    if (-not $AllowFieldChange) {
        $AllowFieldChange = $true
        Log 'INFO' "ConnectMode $ConnectMode refreshes the schema -- field-count changes allowed (no rollback)."
    }
}
Log 'INFO' "Path=$Path  NewServer=$NewServer  OldServer=$OldServer  NewDatabase=$NewDatabase  NewUser=$NewUser  Password=$pwShown  $cmShown"

$livePw = ''
if ($Connect) {
    if ($mode -ne 'APPLY') { Log 'WARN' '-Connect only applies in APPLY mode -- ignored.'; $Connect = $false }
    elseif (-not $NewUser)     { Log 'ERR' '-Connect requires -NewUser.';     exit 4 }
    elseif (-not $NewDatabase) { Log 'ERR' '-Connect requires -NewDatabase.'; exit 4 }
    else {
        $livePw = Get-PlainPassword -Supplied $NewPassword -User $NewUser
        if (-not $livePw) { Log 'ERR' 'No password supplied -- cannot open a live connection. Nothing was written.'; exit 4 }
    }
}
Log 'INFO' "Found $($rpts.Count) report file(s)."

# ------------------------------------------------------------------ counters
$filesScanned = 0
$okCount      = 0
$errCount     = 0
$skipCount    = 0
$tblTotal     = 0
# subScopes / subTables exist so the summary can never again look healthy while
# every subreport is being skipped. The Return reports reported tables=6 (3 per
# file, main only) for a pair that really holds 8 tables. If a report has
# subreports these must be non-zero.
$subScopes    = 0
$subTables    = 0
$tblChanged   = 0
$tblSkipped   = 0
$vMismatch    = 0
$vStale       = 0
$fieldDiff    = 0
$removedKeys  = 0
$sqlDbUnset   = 0
$dllInLp      = 0
$userUnset    = 0
$warnedSqlDb  = $false
$warnedDllLp  = $false
$warnedUser   = $false
$connectOk    = 0
$connectFail  = 0
$fieldChanged = 0
$logonApplied = 0
$fallbackUsed = 0
$alreadyOk    = 0
$userKept     = 0     # table had a UserID and the bag write carried it across
$userStamped  = 0     # table had NO UserID and -NewUser created one
$userDropped  = 0     # bag write was impossible -- UserID really was lost

# -Verify detail lists: "<file> || <scope>|<alias> || <what is wrong>"
$mismatchList = New-Object System.Collections.ArrayList
# -Verify detail list for tables poisoned with 'Database DLL' inside QE_LogonProperties
$poisonList   = New-Object System.Collections.ArrayList
$vPoison      = 0
$staleList    = New-Object System.Collections.ArrayList

foreach ($f in $rpts) {
    $filesScanned++
    $full = $f.FullName
    Log 'INFO' "--> $full"

    $doc      = New-Object CrystalDecisions.CrystalReports.Engine.ReportDocument
    $before   = @{}      # key -> field count before the write
    $beforeUser = @{}    # key -> UserID before the write (verification target)
    $intended = @{}      # key -> expected server/db after the write
    $changed  = 0
    $loadOk   = $false

    try {
        $doc.Load($full, [CrystalDecisions.Shared.OpenReportMethod]::OpenReportByTempCopy)
        $loadOk = $true

        if ($ListConnections) {
            Show-Connections -Doc $doc -File $f.Name
            try { $doc.Close(); $doc.Dispose() } catch {}
            continue
        }

        # -ConnectionFrom <alias>: borrow that table's LogOnInfo for every table
        # instead of building a fresh connection per table.
        $donorLi = $null
        if ($ConnectionFrom) {
            foreach ($et in $doc.Database.Tables) {
                if ([string]$et.Name -eq $ConnectionFrom) { $donorLi = $et.LogOnInfo; break }
            }
            if ($null -eq $donorLi) {
                Log 'ERR' "    -ConnectionFrom '$ConnectionFrom' not found in $($f.Name) -- skipping file."
                $errCount++
                try { $doc.Close(); $doc.Dispose() } catch {}
                continue
            }
            if (-not [string]$donorLi.ConnectionInfo.UserID) {
                Log 'WARN' "    donor '$ConnectionFrom' has a blank UserID -- ApplyLogOnInfo will not persist."
            }
            $donorLi.ConnectionInfo.ServerName = $NewServer
            if ($NewDatabase) { $donorLi.ConnectionInfo.DatabaseName = $NewDatabase }
            if ($NewUser)     { $donorLi.ConnectionInfo.UserID       = $NewUser }
            if ($livePw)      { $donorLi.ConnectionInfo.Password     = $livePw }
            Log 'INFO' "    reusing connection of '$ConnectionFrom' for every table"
        }

        # ENGINE PASS -- best effort only, NOT the write that persists.
        # ApplyLogOnInfo keeps the engine layer's logon in step with where the
        # table is going, which matters for -Connect in this same session. It does
        # NOT move the table in the saved file (measured -- see the RAS block
        # below), so its result is never allowed to skip the RAS write.
        # -Method ApplyLogOnInfo runs this pass ALONE and therefore writes nothing
        # to disk offline; it is kept only for diagnosing the engine layer.
        # ApplyLogOnInfo writes NOTHING, silently, on a table whose stored UserID
        # is blank. Those tables must not be reported as "moved by ApplyLogOnInfo"
        # -- they have to go down the RAS path, which now sets the UserID through
        # the logon bag. Snapshot the blank ones BEFORE anything is applied.
        if ($Method -eq 'ApplyLogOnInfo' -and -not $script:warnedEngineOnly) {
            Log 'WARN' '    -Method ApplyLogOnInfo does NOT persist a server/database change offline.'
            Log 'WARN' '    Use -Method Auto (default) or -Method SetTableLocation to actually move the tables.'
            $script:warnedEngineOnly = $true
        }
        $logonFailed = @{}
        $logonBlank  = @{}
        foreach ($et in $doc.Database.Tables) {
            $u0 = ''
            try { $u0 = [string]$et.LogOnInfo.ConnectionInfo.UserID }
            catch { Log 'WARN' "    cannot read engine UserID on '$($et.Name)' -- treated as blank: $($_.Exception.Message.Split([char]10)[0])" }
            if (-not $u0) { $logonBlank[[string]$et.Name] = $true }
        }
        if ($logonBlank.Count) {
            Log 'INFO' "    $($logonBlank.Count) table(s) have a blank UserID -- ApplyLogOnInfo cannot write them; using the RAS logon-bag path"
        }
        if ($Method -ne 'SetTableLocation') {
            foreach ($et in $doc.Database.Tables) {
                if ($logonBlank.ContainsKey([string]$et.Name)) { continue }
                try {
                    # Already where it should be: leave it alone. Rewriting a
                    # correct table gains nothing and costs its UserID.
                    $cci = $et.LogOnInfo.ConnectionInfo
                    $sameSrv = ($cci.ServerName -eq $NewServer)
                    $sameDb  = ((-not $NewDatabase) -or ($cci.DatabaseName -eq $NewDatabase))
                    if ($sameSrv -and $sameDb -and -not $Force -and -not $ConnectionFrom) {
                        Log 'INFO' "    [$($et.Name)] already on $NewServer/$($cci.DatabaseName) -- left untouched"
                        $alreadyOk++
                        continue
                    }
                    if ($donorLi) { $et.ApplyLogOnInfo($donorLi) }
                    else { Set-TableByLogOnInfo -EngineTable $et -Server $NewServer -Database $NewDatabase -User $NewUser -Password $livePw }
                    $logonApplied++
                } catch {
                    $logonFailed[[string]$et.Name] = $true
                    Log 'WARN' "    ApplyLogOnInfo failed on '$($et.Name)': $($_.Exception.Message.Split([char]10)[0])"
                }
            }
        }

        # ConnectMode 3: log on BEFORE any SetTableLocation, in case Crystal only
        # keeps the SQL-database flag for a connection it already knows is live.
        if ($Connect -and $ConnectMode -eq 3) {
            try {
                $doc.ReportClientDocument.DatabaseController.LogonEx($NewServer, $NewDatabase, $NewUser, $livePw)
                Log 'INFO' '    pre-logon done (ConnectMode 3)'
            } catch { Log 'WARN' ("    pre-logon failed -- carrying on: " + (Get-ExDetail $_)) }
        }

        foreach ($sc in (Get-Scopes -Doc $doc)) {
            $dbc = $sc.Dbc
            $cnt = Get-ScopeTableCount -Scope $sc
            if ($cnt -le 0) { continue }
            if ($sc.Scope -ne 'main') { $subScopes++; $subTables += $cnt }

            for ($i = 0; $i -lt $cnt; $i++) {
                # Re-fetch by index: SetTableLocation replaces the item in place.
                $t     = $dbc.Database.Tables[$i]
                $alias = [string]$t.Alias
                $key   = "$($sc.Scope)|$alias"
                $cur   = Read-TableConn -Table $t
                $tblTotal++

                $wantDb = if ($NewDatabase) { $NewDatabase } else { $cur.Db }

                # ---------------- read-only mode ----------------
                if ($Verify) {
                    $bad = @()
                    if ($cur.Server -ne $NewServer) { $bad += "QE_ServerDescription='$($cur.Server)'" }
                    if ($cur.DS -and $cur.DS -ne $NewServer) { $bad += "DataSource='$($cur.DS)'" }
                    if ($NewDatabase) {
                        if ($cur.Db -ne $NewDatabase) { $bad += "QE_DatabaseName='$($cur.Db)'" }
                        if ($cur.IC -and $cur.IC -ne $NewDatabase) { $bad += "InitialCatalog='$($cur.IC)'" }
                    }
                    if ($bad.Count) {
                        Log 'ERR' "    MISMATCH [$($sc.Scope)] $alias -- $($bad -join ' ')"
                        [void]$mismatchList.Add("$full || $($sc.Scope)|$alias || $($bad -join ' ')")
                        $vMismatch++
                    }
                    $stale = @()
                    if ($cur.PreSrv -and $cur.PreSrv -ne $NewServer) { $stale += "PreQEServerName='$($cur.PreSrv)'" }
                    if ($NewDatabase -and $cur.PreDb -and $cur.PreDb -ne $NewDatabase) { $stale += "PreQEDatabaseName='$($cur.PreDb)'" }
                    if ($stale.Count) {
                        Log 'WARN' "    STALE    [$($sc.Scope)] $alias -- $($stale -join ' ')"
                        [void]$staleList.Add("$full || $($sc.Scope)|$alias || $($stale -join ' ')")
                        $vStale++
                    }
                    # A table can sit on exactly the right server and database and still
                    # be unopenable in SAP B1. See the DLL-IN-LOGON-BAG note further down:
                    # 'Database DLL' inside QE_LogonProperties breaks the OLE DB connection
                    # string. -Verify used to report mismatch=0 on precisely those files,
                    # which is the same silent lie this script was rewritten to stop telling.
                    if ($cur.DllInLp) {
                        Log 'ERR' "    POISONED [$($sc.Scope)] $alias -- 'Database DLL' inside QE_LogonProperties (SAP B1 will refuse this report)"
                        [void]$poisonList.Add("$full || $($sc.Scope)|$alias || Database DLL in QE_LogonProperties")
                        $vPoison++
                    }
                    continue
                }

                # ---------------- OldServer filter ----------------
                if ($OldServer -and ($cur.Server -ne $OldServer)) {
                    Log 'INFO' "    skip [$($sc.Scope)] $alias (server='$($cur.Server)' != OldServer)"
                    $tblSkipped++
                    continue
                }

                # -1 disables the field-count guard for this table, so it has to say so.
                try { $before[$key] = [int]$t.DataFields.Count }
                catch {
                    $before[$key] = -1
                    Log 'WARN' "    [$($sc.Scope)] $alias : DataFields.Count unreadable -- field-count guard DISABLED for this table: $($_.Exception.Message.Split([char]10)[0])"
                }
                $beforeUser[$key] = $cur.User
                $intended[$key] = [pscustomobject]@{ Server = $NewServer; Db = $wantDb }

                if ($WhatIfPreference) {
                    Log 'HIT' "    WHATIF [$($sc.Scope)] $alias : '$($cur.Server)/$($cur.Db)' -> '$NewServer/$wantDb'"
                    $changed++
                    continue
                }

                # ---------------- clone, rewrite, apply ----------------
                # Order matters: mutate the nested logon bag IN PLACE first, then the
                # outer keys. Never assign the bag object back (see rule 3 above).
                #
                # AUTO ALWAYS WRITES HERE -- measured 2026-08-31, read back FROM DISK.
                # Auto used to `continue` at this point whenever ApplyLogOnInfo had
                # not thrown, on the belief that the engine call had already moved
                # the table. It had not. ApplyLogOnInfo does NOT persist a server or
                # database change to the file at all when the report is edited
                # OFFLINE -- not even with a stored UserID on both the engine and
                # the RAS layer. Proven twice on copies:
                #   B25.rpt      sa on both layers, Auto -> 'AUTOSRV/AUTODB'
                #                logonApplied=3, changed=0, disk still 10.0.70.61
                #                /SBO_SDA_REAL, all 3 tables ROLLED BACK.
                #   Return x2    sa on both layers, Auto -> SBO_Seoul_XX
                #                logonApplied=6, changed=0, disk still SBO_Seoul_Re,
                #                6 tables ROLLED BACK.
                # The identical move with -Method SetTableLocation succeeded with
                # err=0 and kept every UserID. So the engine pass above is now only
                # a best-effort extra that keeps the engine layer in step; the write
                # that actually reaches the file is the RAS one below, and Auto must
                # never skip it. Do not re-add a `continue` here.
                # -Method ApplyLogOnInfo means "engine pass only". It is kept for
                # diagnosing the engine layer and is PROVEN not to reach the file,
                # so let it save, fail verification and roll back with the hint --
                # that is a far clearer signal than a silent no-op.
                if ($Method -eq 'ApplyLogOnInfo') { $changed++; continue }
                if ($Method -eq 'Auto') { $fallbackUsed++ }

                try {
                    $clone = $t.Clone($true)
                    $at    = $clone.ConnectionInfo.Attributes
                    $lp    = Get-BagValue -Bag $at -Key 'QE_LogonProperties'

                    if ($lp) {
                        Set-BagString -Bag $lp -Key 'Data Source' -Value $NewServer
                        if ($NewDatabase) { Set-BagString -Bag $lp -Key 'Initial Catalog' -Value $NewDatabase }

                        if (-not $NoDesignerNormalize) {
                            # Match what Crystal Designer's Set Datasource Location -> Update
                            # leaves behind. Diffed property-by-property against the one report
                            # Ars repaired by hand -- the only version SAP B1 would open.
                            #   PreQEServerName / PreQEDatabaseName -> DELETED, not rewritten.
                            #     They record the PREVIOUS location, so carrying them forward
                            #     with the new value is meaningless; the working file has none.
                            #   Database DLL inside the logon bag -> DELETED. Stale duplicate;
                            #     the real one lives on the outer Attributes bag.
                            foreach ($dead in @('PreQEServerName','PreQEDatabaseName','Database DLL')) {
                                if (Remove-BagKey -Bag $lp -Key $dead) { $script:removedKeys++ }
                            }
                            # Designer normalises the casing to 'False'.
                            if (Test-BagKey -Bag $lp -Key 'Integrated Security') {
                                Set-BagString -Bag $lp -Key 'Integrated Security' -Value 'False'
                            }
                        }
                    }
                    Set-BagString -Bag $at -Key 'QE_ServerDescription' -Value $NewServer
                    if ($NewDatabase) { Set-BagString -Bag $at -Key 'QE_DatabaseName' -Value $NewDatabase }

                    if (-not $NoDesignerNormalize) {
                        # THE bug. QE_SQLDB=False tells Crystal "this is not a SQL database",
                        # so a SQL Command cannot be executed against it and SAP B1 fails to
                        # open the report. The hand-fixed report has True on every table; every
                        # table this script wrote had False. It is a BOOLEAN, not the string
                        # 'True' -- Set-BagString here would keep the bug and look correct.
                        Set-BagBool -Bag $at -Key 'QE_SQLDB' -Value $true
                    }

                    # THE User ID. SetTableLocation blanks it unless it is written
                    # into the nested logon bag FIRST -- ConnectionInfo.UserName is
                    # a floating value that never reaches the file. -NewUser wins;
                    # otherwise carry whatever the table already had. See header.
                    $wantUser = if ($NewUser) { $NewUser } else { $cur.User }
                    $userOk   = $false
                    if ($wantUser) {
                        if ($null -eq $lp) {
                            Log 'ERR' "    [$($sc.Scope)] $alias : connection has no QE_LogonProperties bag -- User ID '$wantUser' CANNOT be carried across"
                        } else {
                            $userOk = Set-BagUser -LogonBag $lp -User $wantUser
                            if (-not $userOk) {
                                Log 'ERR' "    [$($sc.Scope)] $alias : writing 'User ID' into the logon bag failed -- User ID '$wantUser' will be LOST"
                            }
                        }
                        if ($userOk) {
                            if ($cur.User) { $userKept++ }
                            else { $userStamped++; Log 'OK' "    [$($sc.Scope)] $alias : User ID stamped '$wantUser' (was blank)" }
                        } else {
                            $userDropped++
                            Log 'ERR' "    *** [$($sc.Scope)] $alias : SAP B1 may refuse to open this layout ***"
                        }
                    } else {
                        Log 'WARN' "    [$($sc.Scope)] $alias : no User ID on this table and no -NewUser -- it stays blank (SAP B1 may refuse the layout)"
                    }

                    # ConnectionInfo.UserName / .Password are memory-only: the save
                    # throws them away. They are still set because -Connect reads
                    # them for the live logon in this same session.
                    $ci = $clone.ConnectionInfo
                    if ($wantUser) {
                        $ci.GetType().InvokeMember('UserName', $script:BF::SetProperty, $null, $ci, @([object]$wantUser))
                    }
                    if ($NewPassword) {
                        $ci.GetType().InvokeMember('Password', $script:BF::SetProperty, $null, $ci, @([object]$NewPassword))
                    }

                    $dbc.SetTableLocation($t, $clone)
                    $changed++
                    Log 'INFO' "    set [$($sc.Scope)] $alias : '$($cur.Server)/$($cur.Db)' -> '$NewServer/$wantDb'"
                } catch {
                    Log 'ERR' "    apply failed [$($sc.Scope)] $alias : $($_.Exception.Message)"
                    $errCount++
                }
            }
        }
    } catch {
        Log 'ERR' "    $($f.Name): $($_.Exception.Message)"
        $errCount++
    }

    if (-not $loadOk) {
        try { $doc.Close(); $doc.Dispose() } catch {}
        continue
    }

    # ---------------- read-only / whatif: nothing to save ----------------
    if ($Verify -or $WhatIfPreference) {
        try { $doc.Close(); $doc.Dispose() } catch {}
        if (-not $Verify) {
            if ($changed -eq 0) { Log 'WARN' '    no tables matched -- nothing to save'; $skipCount++ }
            else {
                Log 'INFO' "    WhatIf: would save $changed table change(s) [no save]"
                $tblChanged += $changed
            }
        }
        continue
    }

    if ($changed -eq 0) {
        Log 'WARN' '    no tables matched -- nothing to save'
        $skipCount++
        try { $doc.Close(); $doc.Dispose() } catch {}
        continue
    }

    # ---------------- live connect before saving ----------------
    if ($Connect) {
        $res = Invoke-LiveConnect -Doc $doc -Server $NewServer -Database $NewDatabase -User $NewUser -Password $livePw -Mode $ConnectMode
        if (-not $res.Ok) {
            Log 'ERR' "    LIVE CONNECT FAILED at $($res.Stage): $($res.Msg)"
            Log 'ERR' "    NOT SAVING $($f.Name) -- the file on disk is untouched."
            Log 'ERR' "    Check server / database / user / password, and that $NewServer answers on 1433."
            $connectFail++
            $errCount++
            try { $doc.Close(); $doc.Dispose() } catch {}
            continue
        }
        Log 'OK' "    live connection to $NewServer/$NewDatabase verified as '$NewUser'"
        $connectOk++
    }

    # ---------------- save (with a rollback copy) ----------------
    $tmpCopy = $null
    $saved   = $false
    try {
        if ($BackupSuffix) {
            $bak = "$full$BackupSuffix"
            Copy-Item -LiteralPath $full -Destination $bak -Force -WhatIf:$false -Confirm:$false
            Log 'INFO' "    backup -> $bak"
        }
        $tmpCopy = Join-Path $env:TEMP ("_sdl_" + [guid]::NewGuid().ToString('N') + '.rpt')
        Copy-Item -LiteralPath $full -Destination $tmpCopy -Force -WhatIf:$false -Confirm:$false

        $doc.SaveAs($full)
        $saved = $true
        Log 'INFO' "    saved $changed table change(s) -- verifying from disk"
    } catch {
        Log 'ERR' "    save failed: $($_.Exception.Message)"
        $errCount++
    } finally {
        try { $doc.Close(); $doc.Dispose() } catch {}
    }

    if (-not $saved) {
        if ($tmpCopy -and (Test-Path -LiteralPath $tmpCopy)) { Remove-Item -LiteralPath $tmpCopy -Force -ErrorAction SilentlyContinue }
        continue
    }

    # ---------------- re-open from disk and prove it ----------------
    # The whole point of the rewrite: never trust the API, read the file back.
    $fileErr = 0
    $doc2 = New-Object CrystalDecisions.CrystalReports.Engine.ReportDocument
    try {
        $doc2.Load($full, [CrystalDecisions.Shared.OpenReportMethod]::OpenReportByTempCopy)
        $seen = 0
        foreach ($sc in (Get-Scopes -Doc $doc2)) {
            $cnt = Get-ScopeTableCount -Scope $sc -Where 'verify'
            for ($i = 0; $i -lt $cnt; $i++) {
                $t2    = $sc.Dbc.Database.Tables[$i]
                $alias = [string]$t2.Alias
                $key   = "$($sc.Scope)|$alias"
                if (-not $intended.ContainsKey($key)) { continue }   # skipped by -OldServer
                $seen++
                $got  = Read-TableConn -Table $t2
                $want = $intended[$key]

                $probs = @()
                if ($got.Server -ne $want.Server)                { $probs += "QE_ServerDescription='$($got.Server)'" }
                if ($got.DS -and $got.DS -ne $want.Server)       { $probs += "DataSource='$($got.DS)'" }
                if ($got.Db -ne $want.Db)                        { $probs += "QE_DatabaseName='$($got.Db)'" }
                if ($got.IC -and $got.IC -ne $want.Db)           { $probs += "InitialCatalog='$($got.IC)'" }
                if ($got.PreSrv -and $got.PreSrv -ne $want.Server) { $probs += "PreQEServerName='$($got.PreSrv)'" }
                if ($got.PreDb  -and $got.PreDb  -ne $want.Db)     { $probs += "PreQEDatabaseName='$($got.PreDb)'" }
                if ($probs.Count) {
                    Log 'ERR' "    VERIFY [$($sc.Scope)] $alias : wanted '$($want.Server)/$($want.Db)' but disk has $($probs -join ' ')"
                    $fileErr++
                }

                # ---------------------------------------------------- DLL-IN-LOGON-BAG
                # THE reason SAP B1 refuses a report this script wrote offline.
                # Measured 2026-09-01 by dumping every property of a file Ars repaired
                # by hand against the same file after a script run:
                #
                #   file                            'Database DLL' in QE_LogonProperties
                #   SO(Dis)  repaired in Designer   NO   -> opens in SAP B1
                #   SBO_Seoul reference copy        NO   -> opens in SAP B1
                #   SO(Bom)  script WITH -Connect   NO   -> opens in SAP B1
                #   Return   script offline         YES  -> SAP B1 refuses it
                #
                # QE_LogonProperties is the OLE DB connection string. 'Database DLL' is
                # a Crystal-internal key, not an OLE DB keyword, so the provider rejects
                # the whole string. Its legitimate home is the OUTER Attributes bag.
                #
                # SetTableLocation() injects it, at SAVE time, every time:
                #   load + SaveAs, nothing modified        -> stays clean
                #   load + SetTableLocation + SaveAs       -> key appears
                # It cannot be removed offline. Proven dead ends:
                #   - Remove it from the clone before SetTableLocation -> re-injected.
                #   - Remove it from the outer bag too -> "Invalid argument for database."
                #   - Remove it from the live table after SetTableLocation -> the key is
                #     not even there in memory; it is added by the serializer.
                #   - Reload the saved file, remove it, SaveAs -> reverts on reload.
                #   - ReplaceConnection() (the API behind Designer's Update button)
                #     CRASHES the process. Blacklisted alongside Get/SetConnectionInfos.
                # A live VerifyDatabase is the only thing that rewrites the bag cleanly,
                # so -Connect is not optional for a file that must open in SAP B1.
                if ($got.DllInLp) {
                    if (-not $script:warnedDllLp) {
                        Log 'ERR' "    *** 'Database DLL' is inside QE_LogonProperties -- SAP B1 WILL REFUSE THIS REPORT ***"
                        Log 'ERR' "    SetTableLocation injects it and no offline edit can remove it."
                        Log 'ERR' "    Re-run with -Connect -NewUser <user> (needs the server reachable), or:"
                        Log 'ERR' "      Database > Set Datasource Location > select the connection > Update > Save."
                        $script:warnedDllLp = $true
                    }
                    $dllInLp++
                }
                # QE_SQLDB: only a successful VerifyDatabase sets it True. It is NOT the
                # cause of the SAP B1 failures -- both hand-repaired reference files that
                # open fine carry False. Kept as an informational counter only; do not
                # re-promote this to an error.
                if ($got.SqlDb -ne 'True') { $sqlDbUnset++ }
                # UserID is now a HARD check: a blank one stops SAP B1 opening the
                # layout, so a report that lost it must be rolled back, not warned
                # about. $wantU is -NewUser when given, otherwise whatever the
                # table carried before the write.
                $wantU = if ($NewUser) { $NewUser } else { $beforeUser[$key] }
                if ($wantU -and $got.User -ne $wantU) {
                    Log 'ERR' "    VERIFY userLost [$($sc.Scope)] $alias : User ID '$wantU' -> '$($got.User)' on disk"
                    $userUnset++; $fileErr++
                } elseif (-not $wantU -and -not $got.User) {
                    if (-not $script:warnedUser) {
                        Log 'WARN' "    User ID is blank on disk and no -NewUser was given -- SAP B1 may refuse this report."
                        Log 'WARN' "    Re-run with -NewUser <login> to stamp one (it works even on a blank table)."
                        $script:warnedUser = $true
                    }
                    $userUnset++
                }

                $fc = -1
                try { $fc = [int]$t2.DataFields.Count }
                catch { Log 'WARN' "    [$($sc.Scope)] $alias : DataFields.Count unreadable -- field-count check skipped: $($_.Exception.Message.Split([char]10)[0])" }
                if ($before[$key] -ge 0 -and $fc -ne $before[$key]) {
                    if ($AllowFieldChange) {
                        Log 'INFO' "    field count [$($sc.Scope)] $alias : $($before[$key]) -> $fc (schema refreshed, allowed)"
                        $fieldChanged++
                    } else {
                        Log 'ERR' "    VERIFY fieldDiff [$($sc.Scope)] $alias : $($before[$key]) -> $fc"
                        $fieldDiff++; $fileErr++
                    }
                }
            }
        }
        if ($seen -ne $intended.Count) {
            Log 'ERR' "    VERIFY: expected $($intended.Count) table(s) on disk, found $seen"
            $fileErr++
        }
    } catch {
        Log 'ERR' "    VERIFY load failed: $($_.Exception.Message)"
        $fileErr++
    } finally {
        try { $doc2.Close(); $doc2.Dispose() } catch {}
    }

    if ($fileErr -gt 0) {
        # A bad write must never survive -- put the original back.
        try {
            Copy-Item -LiteralPath $tmpCopy -Destination $full -Force -WhatIf:$false -Confirm:$false
            Log 'ERR' "    ROLLED BACK $($f.Name) -- $fileErr verification error(s), file left unchanged"
            if ($Method -eq 'ApplyLogOnInfo') {
                Log 'ERR' "    -Method ApplyLogOnInfo cannot write this offline. Re-run with -Method SetTableLocation."
            }
        } catch {
            Log 'ERR' "    ROLLBACK FAILED for $($f.Name): $($_.Exception.Message)"
        }
        $errCount += $fileErr
    } else {
        $dbShown = if ($NewDatabase) { $NewDatabase } else { '<per-table>' }
        Log 'OK' "    verified on disk: $changed table(s) now point to $NewServer/$dbShown"
        $okCount++
        $tblChanged += $changed
    }

    if ($tmpCopy -and (Test-Path -LiteralPath $tmpCopy)) { Remove-Item -LiteralPath $tmpCopy -Force -ErrorAction SilentlyContinue }
}

Log 'INFO' '=============================================================='
if ($Verify) {
    if ($mismatchList.Count) {
        Log 'INFO' "MISMATCH detail ($($mismatchList.Count)):"
        foreach ($m in $mismatchList) { Log 'INFO' "  M | $m" }
    }
    if ($staleList.Count) {
        Log 'INFO' "STALE detail ($($staleList.Count)):"
        foreach ($m in $staleList) { Log 'INFO' "  S | $m" }
    }
    if ($poisonList.Count) {
        Log 'INFO' "POISONED detail ($($poisonList.Count)) -- these need -Connect or a Designer Update:"
        foreach ($m in $poisonList) { Log 'INFO' "  P | $m" }
    }
    Log 'INFO' "SUMMARY mode=VERIFY  files=$filesScanned  tables=$tblTotal  subreports=$subScopes  subTables=$subTables  mismatch=$vMismatch  stale=$vStale  poisoned=$vPoison"
    Log 'INFO' '=============================================================='
    Write-Host "RESULT|VERIFY|files=$filesScanned|tables=$tblTotal|subreports=$subScopes|subTables=$subTables|mismatch=$vMismatch|stale=$vStale|poisoned=$vPoison"
    exit ([int]([bool]($vMismatch + $vStale)))
}
Log 'INFO' "SUMMARY mode=$mode  files=$filesScanned  ok=$okCount  err=$errCount  skippedFiles=$skipCount  tables=$tblTotal  subreports=$subScopes  subTables=$subTables  changed=$tblChanged  skippedTables=$tblSkipped  fieldDiff=$fieldDiff  removedStaleKeys=$removedKeys  sqlDbUnset=$sqlDbUnset  dllInLogonBag=$dllInLp  userUnset=$userUnset  connectOk=$connectOk  connectFail=$connectFail  fieldChanged=$fieldChanged  logonApplied=$logonApplied  fallback=$fallbackUsed  alreadyOk=$alreadyOk  method=$Method"
Log 'INFO' "USER ID  kept=$userKept  stamped=$userStamped  dropped=$userDropped"
if ($userDropped) { Log 'ERR' "*** $userDropped table(s) could not keep a User ID -- those layouts may not open in SAP B1 ***" }
Log 'INFO' '=============================================================='
Write-Host "RESULT|$mode|files=$filesScanned|ok=$okCount|err=$errCount|tables=$tblTotal|subreports=$subScopes|subTables=$subTables|changed=$tblChanged|skippedTables=$tblSkipped|fieldDiff=$fieldDiff|removedStaleKeys=$removedKeys|sqlDbUnset=$sqlDbUnset|dllInLogonBag=$dllInLp|userUnset=$userUnset|userKept=$userKept|userStamped=$userStamped|userDropped=$userDropped|connectOk=$connectOk|connectFail=$connectFail|fieldChanged=$fieldChanged|logonApplied=$logonApplied|fallback=$fallbackUsed|alreadyOk=$alreadyOk|method=$Method"

exit ([int]([bool]$errCount))
