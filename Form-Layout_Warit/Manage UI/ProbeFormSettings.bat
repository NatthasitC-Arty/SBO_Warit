@echo off
chcp 65001 >nul
REM ============================================================
REM  Probe SAP B1 schema for matrix Form Settings table.
REM  Reuses connection settings from ..\ImportUDV\_settings.bat
REM  (avoids credential duplication).
REM ============================================================
if not exist "%~dp0_settings.bat" (
    echo ERROR: _settings.bat not found.
    echo Copy _settings.bat.example to _settings.bat and edit it.
    pause
    exit /b 1
)
call "%~dp0_settings.bat"

set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Probe-FormSettings.ps1" ^
    -Server "%SERVER%" ^
    -CompanyDB "%COMPANYDB%" ^
    -DBUser "%DBUSER%" ^
    -DBPassword "%DBPASSWORD%"

echo.
pause
