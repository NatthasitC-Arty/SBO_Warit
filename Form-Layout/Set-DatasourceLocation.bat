@echo off
REM ============================================================
REM  Set-DatasourceLocation.bat
REM  Wrapper for Set-DatasourceLocation.ps1 -- runs under 64-bit
REM  PowerShell (Crystal Reports Runtime is installed in GAC_64).
REM
REM  Interactive mode (no args): asks only Server + Database, then
REM  lets you pick which subfolder(s) to apply to.
REM
REM  Pass-through mode: forwards all args to the .ps1 unchanged.
REM    Set-DatasourceLocation.bat -Path "..." -NewServer NEWSRV -WhatIf
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
