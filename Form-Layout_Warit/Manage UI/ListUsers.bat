@echo off
chcp 65001 >nul
REM ============================================================
REM  List SAP B1 users from OUSR + (optional) per-user CPRF
REM  row counts for a given FormID.
REM ============================================================
if not exist "%~dp0_settings.bat" (
    echo ERROR: _settings.bat not found.
    pause & exit /b 1
)
call "%~dp0_settings.bat"

REM --- Optional: show CPRF row counts per user for this FormID ---
set FORMID=139

set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\List-Users.ps1" ^
    -Server "%SERVER%" ^
    -CompanyDB "%COMPANYDB%" ^
    -DBUser "%DBUSER%" ^
    -DBPassword "%DBPASSWORD%" ^
    -FormID "%FORMID%"

echo.
pause
