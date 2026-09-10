@echo off
REM ============================================================
REM  Set-DatasourceLocation.bat
REM  Wrapper for Set-DatasourceLocation.ps1 -- runs under 64-bit
REM  PowerShell (Crystal Reports Runtime is installed in GAC_64).
REM
REM  Interactive mode (no args): asks Server + Database + mode + SQL user +
REM  write method, then lets you pick which subfolder(s) to apply to.
REM
REM  -Method picks HOW the change is written:
REM    Auto             engine logon pass + table rewrite. Default, recommended.
REM    SetTableLocation table rewrite only. Same result on disk.
REM    ApplyLogOnInfo   engine logon only -- DIAGNOSTIC. Measured 2026-08-31:
REM                     it persists NOTHING offline, so the run will roll back.
REM
REM  -NewUser <u> stamps a SQL user into every table. A report whose user is
REM  blank cannot be opened by SAP B1, and this now works even when the report
REM  has no user at all. It does NOT require -Connect.
REM
REM  Pass-through mode: forwards all args to the .ps1 unchanged.
REM    Set-DatasourceLocation.bat -Path "..." -NewServer NEWSRV -WhatIf
REM    Set-DatasourceLocation.bat -Path "..." -NewServer NEWSRV -NewDatabase DB -Verify
REM
REM  -Verify is read-only: it opens every report and lists the tables whose
REM  server/database do not match, plus the ones whose PreQE* still point
REM  somewhere else. Run it before and after to prove the change landed.
REM
REM  -Connect -NewUser <u> logs on to the database for real before saving.
REM  Required: QE_SQLDB is only stamped True after a successful connection, and
REM  SAP B1 will not open a report whose QE_SQLDB is False. If the logon fails
REM  the report is NOT saved. The password is prompted for by the .ps1.
REM ============================================================

setlocal EnableDelayedExpansion
set "SCRIPT_DIR=%~dp0"
set "ROOT=%SCRIPT_DIR:~0,-1%"
set "PS1=%SCRIPT_DIR%Set-DatasourceLocation.ps1"
set "PS64=%WINDIR%\System32\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%PS64%" ( echo [ERROR] 64-bit PowerShell not found at %PS64% & pause & exit /b 2 )
if not exist "%PS1%"  ( echo [ERROR] PS1 script not found: %PS1%               & pause & exit /b 3 )

if not "%~1"=="" goto :runWithArgs

echo.
echo === Set Datasource Location (Crystal Reports .rpt) ===
echo Root folder: %ROOT%
echo.

set /p NEW_SERVER=New SQL Server name (required):
if "!NEW_SERVER!"=="" ( echo [ERROR] NewServer is required. & pause & exit /b 4 )

set /p NEW_DB=New database name (blank = keep existing):

echo.
echo Mode:
echo   1. Apply   -- write the change, then re-open each file to verify it
echo   2. Verify  -- read-only, just report what does not match
echo   3. WhatIf  -- show what would change, write nothing
set /p MODE=Pick mode [1]:
if "!MODE!"=="" set "MODE=1"
set "EXTRA="
if "!MODE!"=="2" set "EXTRA=-Verify"
if "!MODE!"=="3" set "EXTRA=-WhatIf"

set "SQL_USER="
set "DO_CONNECT="
set "METHOD="
if "!MODE!"=="1" (
    echo.
    echo A report whose SQL user is blank cannot be opened by SAP B1.
    echo A user given here is stamped into every table -- it works even on
    echo reports that have no user at all. Blank = keep whatever each table has.
    set /p SQL_USER=SQL user to write into the reports ^(e.g. sa, blank = keep^):

    echo.
    echo Write method:
    echo   1. Auto             -- engine logon pass + table rewrite. Recommended.
    echo   2. SetTableLocation -- table rewrite only. Same result on disk.
    echo   3. ApplyLogOnInfo   -- engine logon only. DIAGNOSTIC: writes nothing.
    set /p METHOD=Pick method [1]:

    echo.
    echo Live connect logs on to the database for real before saving. That is
    echo the ONLY way QE_SQLDB gets set to True -- without it SAP B1 may still
    echo refuse the report. Needs a reachable server and the right password.
    set /p DO_CONNECT=Log on to the database before saving? y/N:
)
if "!METHOD!"=="" set "METHOD=1"
set "METHOD_ARG="
if "!METHOD!"=="2" set "METHOD_ARG=-Method SetTableLocation"
if "!METHOD!"=="3" set "METHOD_ARG=-Method ApplyLogOnInfo"

echo.
echo Subfolders under %ROOT%:
echo   0. ^<ALL^> (entire root, recursive)
set i=0
for /f "delims=" %%D in ('dir /b /ad "%ROOT%" 2^>nul') do (
    set "NAME=%%D"
    if /I not "!NAME:~0,1!"=="." (
        set /a i+=1
        set "F_!i!=%%D"
        echo   !i!. %%D
    )
)
set MAX=!i!
if !MAX! EQU 0 (
    echo [ERROR] No subfolders found under %ROOT%.
    pause
    exit /b 5
)
echo.
set /p PICK=Pick folders (number, comma-separated, or blank=ALL):
if "!PICK!"=="" set "PICK=0"

set "COMMON=-NewServer "!NEW_SERVER!""
if not "!NEW_DB!"=="" set "COMMON=!COMMON! -NewDatabase "!NEW_DB!""
if not "!EXTRA!"==""  set "COMMON=!COMMON! !EXTRA!"
REM The PASSWORD is never handled here -- the .ps1 asks for it with
REM Read-Host -AsSecureString so it is never echoed, logged or stored in a
REM cmd variable. One prompt per folder processed.
if not "!SQL_USER!"=="" set "COMMON=!COMMON! -NewUser "!SQL_USER!""
if not "!METHOD_ARG!"=="" set "COMMON=!COMMON! !METHOD_ARG!"
if /I "!DO_CONNECT!"=="y" (
    if "!SQL_USER!"=="" (
        echo [WARN] -Connect needs a SQL user. No user was given -- skipping the live logon.
    ) else (
        set "COMMON=!COMMON! -Connect"
    )
)

set "GLOBAL_RC=0"
for %%T in (!PICK!) do (
    set "N=%%T"
    set "N=!N: =!"
    if "!N!"=="0" (
        set "TARGET=!ROOT!"
    ) else (
        call set "SUB=%%F_!N!%%"
        if "!SUB!"=="" (
            echo [WARN] selection !N! out of range -- skipping
            set "TARGET="
        ) else (
            set "TARGET=!ROOT!\!SUB!"
        )
    )
    if not "!TARGET!"=="" (
        echo.
        echo --- Processing: !TARGET! ---
        "%PS64%" -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Path "!TARGET!" !COMMON!
        if errorlevel 1 set "GLOBAL_RC=!ERRORLEVEL!"
    )
)

echo.
echo Final exit code: !GLOBAL_RC!
pause
exit /b !GLOBAL_RC!

:runWithArgs
"%PS64%" -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %*
set "RC=%ERRORLEVEL%"
echo.
echo Exit code: %RC%
pause
exit /b %RC%
