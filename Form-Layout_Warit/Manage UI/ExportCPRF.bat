@echo off
chcp 65001 >nul
REM ============================================================
REM  Export CPRF (Form Settings) from SAP B1 to CSV.
REM  Default: exports ALL rows. Edit FILTER vars below to limit
REM  by FormID / ItemID / ColID.
REM ============================================================
if not exist "%~dp0_settings.bat" (
    echo ERROR: _settings.bat not found.
    echo Copy _settings.bat.example to _settings.bat and edit it.
    pause
    exit /b 1
)
call "%~dp0_settings.bat"

REM --- Optional filters (leave empty for "all") ---
set FORMID=139
set ITEMID=
set COLID=

set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Export-CPRF.ps1" ^
    -Server "%SERVER%" ^
    -CompanyDB "%COMPANYDB%" ^
    -DBUser "%DBUSER%" ^
    -DBPassword "%DBPASSWORD%" ^
    -FormID "%FORMID%" ^
    -ItemID "%ITEMID%" ^
    -ColID "%COLID%"

echo.
echo CSV saved to: %~dp0Config\
pause
