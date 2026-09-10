@echo off
chcp 65001 >nul
if not exist "%~dp0_settings.bat" ( echo ERROR: _settings.bat not found. & pause & exit /b 1 )
call "%~dp0_settings.bat"
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Probe-UICU.ps1" ^
    -Server "%SERVER%" -CompanyDB "%COMPANYDB%" -DBUser "%DBUSER%" -DBPassword "%DBPASSWORD%"
echo.
pause
