@echo off
REM ============================================================
REM  Edit-CommandSQL.bat
REM  Wrapper for Edit-CommandSQL.ps1 -- always runs under 64-bit
REM  PowerShell (Crystal Reports RAS assemblies are in GAC_64).
REM
REM  Pass-through mode: forwards all args to the .ps1 unchanged.
REM
REM  Examples:
REM    Edit-CommandSQL.bat -Root "%~dp0." -Pattern "NVARCHAR\s*\(\s*MAX\s*\)" ^
REM         -Replacement "NVARCHAR(4000)" -WhatIf
REM
REM    Edit-CommandSQL.bat -Root "%~dp0." -Pattern "NVARCHAR\s*\(\s*MAX\s*\)" ^
REM         -Replacement "NVARCHAR(4000)"
REM
REM    Edit-CommandSQL.bat -Root "%~dp0." -File "1.Purchase Order_ENG*.rpt" ^
REM         -Alias AP_PO -Pattern "OUGP\.UgpCode" -Verify
REM ============================================================

setlocal EnableDelayedExpansion
set "SCRIPT_DIR=%~dp0"
set "ROOT=%SCRIPT_DIR:~0,-1%"
set "PS1=%SCRIPT_DIR%Edit-CommandSQL.ps1"
set "PS64=%WINDIR%\System32\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%PS64%" ( echo [ERROR] 64-bit PowerShell not found at %PS64% & pause & exit /b 2 )
if not exist "%PS1%"  ( echo [ERROR] PS1 script not found: %PS1%          & pause & exit /b 3 )

if not "%~1"=="" goto :runWithArgs

echo.
echo === Edit SQL Command text (Crystal Reports .rpt) ===
echo Root folder: %ROOT%
echo.
echo This wrapper is pass-through only -- supply arguments:
echo.
echo   -Root ^<folder or .rpt^>     required
echo   -Pattern ^<regex^>           required, case-insensitive
echo   -Replacement ^<string^>      .NET substitution syntax
echo   -Alias ^<a1,a2^>             limit to these table aliases
echo   -File ^<mask1,mask2^>        limit to these report file names
echo   -WhatIf                     count only, no write
echo   -Verify                     read-only report of current matches
echo   -LogPath ^<file^>            default _EditCommandSQL.log
echo.
echo Example:
echo   Edit-CommandSQL.bat -Root "%ROOT%" -Pattern "NVARCHAR\s*\(\s*MAX\s*\)" -Replacement "NVARCHAR(4000)" -WhatIf
echo.
pause
exit /b 0

:runWithArgs
"%PS64%" -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %*
set "RC=%ERRORLEVEL%"
echo.
echo Exit code: %RC%
exit /b %RC%
