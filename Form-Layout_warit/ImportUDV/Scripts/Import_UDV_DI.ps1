# ============================================================
# Batch Import User-Defined Values (FMS) to SAP B1 via DI API
# Reads UDV_Map.csv -> oUserQueries (OUQR) + oFormattedSearches (CSHS)
#
# Requirements:
#   - SAP B1 client + DI API installed on this machine
#   - PowerShell architecture must match installed DI API:
#       B1 v10+   -> 64-bit DI API -> 64-bit PowerShell (System32)
#       B1 v9.x   -> 32-bit DI API -> 32-bit PowerShell (SysWOW64)
#     The accompanying .bat auto-selects the right one.
# ============================================================
param(
    [string]$Server      = "SLD-C072",
    [string]$CompanyDB   = "SBO_SDA",
    [string]$DBUser      = "sa",
    [string]$DBPassword  = "1q2w3e4r",
    [string]$SapUser     = "manager",
    [string]$SapPassword = "1q2w3e4r",
    [ValidateSet("MSSQL","HANA")]
    [string]$DBType      = "MSSQL",
    [string]$DIAPIVersion = "",                                  # auto-detect if empty
    [string]$MapFile     = "$PSScriptRoot\..\Config\UDV_Map.csv",
    [string]$LogFile     = "$PSScriptRoot\..\Import_UDV_Log.txt",
    [switch]$DryRun
)

# ------------------------------------------------------------
# Logging
# ------------------------------------------------------------
function Write-Log {
    param([string]$Msg, [string]$Level = "INFO")
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Msg
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

# ------------------------------------------------------------
# Locate & load DI API interop assembly
# Searches Program Files and Program Files (x86) under SAP\
# Returns the first Interop.SAPbobsCOM.dll matching this process bitness.
# ------------------------------------------------------------
function Find-DIAPIAssembly {
    param([string]$VersionHint)

    $is64 = [Environment]::Is64BitProcess
    $roots = @()
    if ($is64) {
        # 64-bit PS: prefer Program Files (no x86)
        $roots += "${env:ProgramFiles}\SAP"
        $roots += "${env:ProgramW6432}\SAP"
    } else {
        # 32-bit PS: only Program Files (x86) DLLs will load
        $roots += "${env:ProgramFiles(x86)}\SAP"
    }
    $roots = $roots | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique

    foreach ($root in $roots) {
        # Try DI API folder with specific version first
        if ($VersionHint) {
            $p = Join-Path $root "SAP Business One DI API\DI API $VersionHint\Interop.SAPbobsCOM.dll"
            if (Test-Path $p) { return $p }
        }
        # Then any Interop.SAPbobsCOM.dll under this SAP root
        $hit = Get-ChildItem -Path $root -Recurse -Filter "Interop.SAPbobsCOM.dll" -ErrorAction SilentlyContinue |
            Sort-Object -Property @{Expression = { $_.FullName -match "DI API" }; Descending = $true} |
            Select-Object -First 1 -ExpandProperty FullName
        if ($hit) { return $hit }
    }
    return $null
}

$diDll = Find-DIAPIAssembly -VersionHint $DIAPIVersion
if (-not $diDll) {
    $arch = if ([Environment]::Is64BitProcess) { "64-bit" } else { "32-bit" }
    Write-Log "DI API interop (Interop.SAPbobsCOM.dll) not found for $arch PowerShell." "ERROR"
    Write-Log "Searched under \$env:ProgramFiles\SAP and \$env:ProgramFiles(x86)\SAP." "ERROR"
    Write-Log "If DI API is installed but DLL is missing, install Data Transfer Workbench (ships Interop wrapper)." "ERROR"
    exit 2
}
try {
    Add-Type -Path $diDll -ErrorAction Stop
} catch {
    Write-Log "Failed to load $diDll : $($_.Exception.Message)" "ERROR"
    Write-Log "Likely an architecture mismatch. Use the matching PowerShell (System32 for 64-bit, SysWOW64 for 32-bit)." "ERROR"
    exit 2
}
Write-Log "Loaded DI API: $diDll"

# ------------------------------------------------------------
# CSV input
#   Columns (header row required):
#     Action          ADD | UPDATE | DELETE | UPSERT
#     FormID          e.g. 139  (Sales Order)
#     ItemID          ItemUID (System Info)
#     ColumnID        ColUID for matrix; blank/0 for header field
#     FMSAction       Q (Query) | F (FixedValue)
#     QueryName       free-text description (reused if already in OUQR)
#     QueryCategory   -1 = General, else category IntrnalKey
#     QueryBody       the saved SQL (use $[$..] tokens; escape $ in CSV as needed)
#     FixedValue      used when FMSAction = F
#     Refresh         Y | N  (auto-refresh on)
#     TriggerID       ItemUID of trigger field (when Refresh=Y)
#     TriggerColumn   ColUID of trigger column (when matrix)
#     ForceRefresh    Y | N  (display saved values immediately)
#
#   For DELETE rows: only Action + FormID + ItemID + ColumnID required.
# ------------------------------------------------------------
if (-not (Test-Path $MapFile)) {
    Write-Log "MapFile not found: $MapFile" "ERROR"
    exit 2
}
$rows = Import-Csv -Path $MapFile

Write-Log "=== Start UDV/FMS Import (DI API) ==="
Write-Log "Server   : $Server"
Write-Log "CompanyDB: $CompanyDB"
Write-Log "DBType   : $DBType"
Write-Log "MapFile  : $MapFile  ($($rows.Count) rows)"
Write-Log "DryRun   : $DryRun"

# ------------------------------------------------------------
# Connect
# ------------------------------------------------------------
$company = New-Object -ComObject SAPbobsCOM.Company
$company.Server       = $Server
$company.CompanyDB    = $CompanyDB
$company.UserName     = $SapUser
$company.Password     = $SapPassword
$company.DbUserName   = $DBUser
$company.DbPassword   = $DBPassword
$company.UseTrusted   = $false
$company.language     = [SAPbobsCOM.BoSuppLangs]::ln_English

if ($DBType -eq "HANA") {
    $company.DbServerType = [SAPbobsCOM.BoDataServerTypes]::dst_HANADB
} else {
    # Best-effort MSSQL version — try newest first
    $tryServerTypes = @(
        [SAPbobsCOM.BoDataServerTypes]::dst_MSSQL2019,
        [SAPbobsCOM.BoDataServerTypes]::dst_MSSQL2017,
        [SAPbobsCOM.BoDataServerTypes]::dst_MSSQL2016,
        [SAPbobsCOM.BoDataServerTypes]::dst_MSSQL2014,
        [SAPbobsCOM.BoDataServerTypes]::dst_MSSQL2012
    )
}

if ($DryRun) {
    Write-Log "DryRun mode — skipping Connect / writes. Will validate CSV only."
} else {
    $connected = $false
    if ($DBType -eq "HANA") {
        $rc = $company.Connect()
        if ($rc -ne 0) {
            Write-Log "Connect failed (HANA): $($company.GetLastErrorDescription())" "ERROR"
            exit 3
        }
        $connected = $true
    } else {
        foreach ($t in $tryServerTypes) {
            $company.DbServerType = $t
            $rc = $company.Connect()
            if ($rc -eq 0) { $connected = $true; Write-Log "Connected (DbServerType=$t)"; break }
        }
        if (-not $connected) {
            Write-Log "Connect failed (MSSQL): $($company.GetLastErrorDescription())" "ERROR"
            exit 3
        }
    }
}

# ------------------------------------------------------------
# Helper: find existing query IntrnalKey by description (OUQR.QName)
# ------------------------------------------------------------
function Find-QueryIDByName {
    param([string]$Name)
    if ($DryRun) { return $null }
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    $rs = $company.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::BoRecordset)
    try {
        $safe = $Name.Replace("'", "''")
        $rs.DoQuery("SELECT TOP 1 ""IntrnalKey"" FROM OUQR WHERE ""QName"" = '$safe'")
        if ($rs.RecordCount -gt 0) { return [int]$rs.Fields.Item(0).Value }
        return $null
    } finally {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($rs)
    }
}

# ------------------------------------------------------------
# Helper: create or reuse Query
# ------------------------------------------------------------
function Ensure-Query {
    param($Row)
    if ($Row.FMSAction -ne "Q") { return $null }
    $name = [string]$Row.QueryName
    if ([string]::IsNullOrWhiteSpace($name)) {
        throw "QueryName is required when FMSAction=Q"
    }
    $existing = Find-QueryIDByName -Name $name
    if ($existing) {
        Write-Log "  Reuse existing Query '$name' (IntrnalKey=$existing)"
        return $existing
    }
    if ($DryRun) {
        Write-Log "  [DryRun] would CREATE Query '$name'"
        return -1
    }
    $uq = $company.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oUserQueries)
    try {
        $cat = if ($Row.QueryCategory) { [int]$Row.QueryCategory } else { -1 }
        $uq.QueryCategory    = $cat
        $uq.QueryDescription = $name
        $uq.Query            = [string]$Row.QueryBody
        if ($uq.Add() -ne 0) {
            throw "UserQueries.Add failed: $($company.GetLastErrorDescription())"
        }
        $newKey = $company.GetNewObjectKey()             # "<cat>,<IntrnalKey>"
        $qid = [int]($newKey -split ',')[1]
        Write-Log "  Created Query '$name' (IntrnalKey=$qid)"
        return $qid
    } finally {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($uq)
    }
}

# ------------------------------------------------------------
# Helper: YN -> BoYesNoEnum
# ------------------------------------------------------------
function To-YN {
    param([string]$Val)
    if ($Val -and $Val.ToUpper() -in @("Y","YES","TRUE","1")) {
        return [SAPbobsCOM.BoYesNoEnum]::tYES
    }
    return [SAPbobsCOM.BoYesNoEnum]::tNO
}

# ------------------------------------------------------------
# Helper: apply FMS properties on a fetched/new object
# ------------------------------------------------------------
function Set-FMSFields {
    param($Fs, $Row, [int]$QueryId)
    $Fs.FormID   = [string]$Row.FormID
    $Fs.ItemID   = [string]$Row.ItemID
    $Fs.ColumnID = [string]$Row.ColumnID
    if ($Row.FMSAction -eq "F") {
        $Fs.Action = [SAPbobsCOM.BoFormattedSearchActionEnum]::bofsa_FixedValue
        $fv = [string]$Row.FixedValue
        if (-not [string]::IsNullOrEmpty($fv)) {
            # DI API property name differs by version: try several
            $set = $false
            foreach ($prop in @('FixedValue','DefaultValue','Value','StringValue','SearchString')) {
                try {
                    $Fs.PSObject.Properties[$prop].Value = $fv
                    $set = $true; break
                } catch {
                    try { $Fs.$prop = $fv; $set = $true; break } catch {}
                }
            }
            if (-not $set) {
                Write-Log "  WARN: cannot set fixed value '$fv' (no compatible DI API property)" "WARN"
            }
        }
    } else {
        $Fs.Action  = [SAPbobsCOM.BoFormattedSearchActionEnum]::bofsa_QueryByQueryID
        $Fs.QueryID = $QueryId
    }
    $Fs.Refresh        = To-YN $Row.Refresh
    $Fs.ForceRefresh   = To-YN $Row.ForceRefresh
    if ($Row.TriggerID) {
        $Fs.TriggerByField = [SAPbobsCOM.BoYesNoEnum]::tYES
        $Fs.TriggerID      = [string]$Row.TriggerID
        if ($Row.TriggerColumn) { $Fs.TriggerColumn = [string]$Row.TriggerColumn }
    }
}

# ------------------------------------------------------------
# Pre-build CSHS key map: (FormID|ItemID|ColID) -> IndexID
# DI API v10 oFormattedSearches.GetByKey takes IndexID (int),
# not (FormID, ItemID, ColumnID).
# ------------------------------------------------------------
function Build-CshsKeyMap {
    $rs = $company.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::BoRecordset)
    $map = @{}
    try {
        $rs.DoQuery("SELECT IndexID, FormID, ItemID, ColID FROM CSHS")
        while (-not $rs.EoF) {
            $f  = $rs.Fields
            $fi = [string]$f.Item("FormID").Value
            $ii = [string]$f.Item("ItemID").Value
            $ci = [string]$f.Item("ColID").Value
            $idx = [int]$f.Item("IndexID").Value
            $map["$fi|$ii|$ci"] = $idx
            $rs.MoveNext()
        }
    } finally {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($rs)
    }
    return $map
}

# ------------------------------------------------------------
# Helper: load existing FMS by composite key via IndexID lookup
# Returns $true if loaded, $false if not found.
# ------------------------------------------------------------
function Get-ExistingFMS {
    param($Fs, $Row, $KeyMap)
    $col = [string]$Row.ColumnID
    if ([string]::IsNullOrEmpty($col)) { $col = "0" }
    $key = "{0}|{1}|{2}" -f [string]$Row.FormID, [string]$Row.ItemID, $col
    if (-not $KeyMap.ContainsKey($key)) { return $false }
    try {
        return [bool]$Fs.GetByKey($KeyMap[$key])
    } catch {
        Write-Log "  GetByKey(IndexID=$($KeyMap[$key])) failed: $($_.Exception.Message)" "WARN"
        return $false
    }
}

# ------------------------------------------------------------
# Build key map for existence checks (skipped on DryRun)
# ------------------------------------------------------------
$script:CshsKeyMap = @{}
if (-not $DryRun) {
    $script:CshsKeyMap = Build-CshsKeyMap
    Write-Log "Pre-loaded $($script:CshsKeyMap.Count) existing CSHS rows for IndexID lookup"
}

# ------------------------------------------------------------
# Main loop
# ------------------------------------------------------------
$ok = 0; $fail = 0; $skip = 0
$rowNo = 0
foreach ($row in $rows) {
    $rowNo++
    $label = "row#$rowNo Form=$($row.FormID) Item=$($row.ItemID) Col=$($row.ColumnID) Action=$($row.Action)"
    try {
        $action = ([string]$row.Action).ToUpper()
        if ($action -notin @("ADD","UPDATE","DELETE","UPSERT")) {
            Write-Log "SKIP unknown Action '$($row.Action)' ($label)" "WARN"; $skip++; continue
        }
        if (-not $row.FormID -or -not $row.ItemID) {
            Write-Log "SKIP missing FormID/ItemID ($label)" "WARN"; $skip++; continue
        }

        # 1. Query side (skip for DELETE / FixedValue)
        $qid = 0
        if ($action -ne "DELETE" -and $row.FMSAction -eq "Q") {
            $qid = Ensure-Query -Row $row
            if (-not $qid) { throw "Failed to resolve QueryID" }
        }

        # 2. FMS side
        if ($DryRun) {
            Write-Log "[DryRun] would $action FMS ($label)"
            $ok++; continue
        }

        $fs = $company.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oFormattedSearches)
        try {
            $exists = Get-ExistingFMS -Fs $fs -Row $row -KeyMap $script:CshsKeyMap

            switch ($action) {
                "ADD" {
                    if ($exists) {
                        Write-Log "SKIP ADD — FMS already exists ($label). Use UPDATE/UPSERT." "WARN"
                        $skip++; continue
                    }
                    # GetByKey leaves the object dirty if not found — rebuild a fresh one
                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($fs)
                    $fs = $company.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oFormattedSearches)
                    Set-FMSFields -Fs $fs -Row $row -QueryId $qid
                    if ($fs.Add() -ne 0) { throw "FMS Add failed: $($company.GetLastErrorDescription())" }
                    Write-Log "OK   ADD FMS ($label)"
                    $ok++
                }
                "UPDATE" {
                    if (-not $exists) {
                        Write-Log "SKIP UPDATE — FMS not found ($label). Use ADD/UPSERT." "WARN"
                        $skip++; continue
                    }
                    Set-FMSFields -Fs $fs -Row $row -QueryId $qid
                    if ($fs.Update() -ne 0) { throw "FMS Update failed: $($company.GetLastErrorDescription())" }
                    Write-Log "OK   UPDATE FMS ($label)"
                    $ok++
                }
                "UPSERT" {
                    if ($exists) {
                        Set-FMSFields -Fs $fs -Row $row -QueryId $qid
                        if ($fs.Update() -ne 0) { throw "FMS Update failed: $($company.GetLastErrorDescription())" }
                        Write-Log "OK   UPSERT->UPDATE FMS ($label)"
                    } else {
                        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($fs)
                        $fs = $company.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oFormattedSearches)
                        Set-FMSFields -Fs $fs -Row $row -QueryId $qid
                        if ($fs.Add() -ne 0) { throw "FMS Add failed: $($company.GetLastErrorDescription())" }
                        Write-Log "OK   UPSERT->ADD FMS ($label)"
                    }
                    $ok++
                }
                "DELETE" {
                    if (-not $exists) {
                        Write-Log "SKIP DELETE — FMS not found ($label)" "WARN"
                        $skip++; continue
                    }
                    if ($fs.Remove() -ne 0) { throw "FMS Remove failed: $($company.GetLastErrorDescription())" }
                    Write-Log "OK   DELETE FMS ($label)"
                    $ok++
                }
            }
        } finally {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($fs)
        }
    } catch {
        Write-Log "FAIL ($label): $($_.Exception.Message)" "ERROR"
        $fail++
    }
}

if (-not $DryRun -and $company.Connected) {
    $company.Disconnect() | Out-Null
}
[void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($company)

Write-Log "=== Done. OK=$ok  FAIL=$fail  SKIP=$skip ==="
if ($fail -gt 0) { exit 1 } else { exit 0 }
