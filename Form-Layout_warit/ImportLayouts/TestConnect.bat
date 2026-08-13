@echo off
chcp 65001 >nul
REM ============================================================
REM  Test SQL connection to SAP B1 Company DB
REM  Connection settings are in _settings.bat (shared, gitignored).
REM ============================================================
if not exist "%~dp0_settings.bat" (
    echo ERROR: _settings.bat not found.
    echo Copy _settings.bat.example to _settings.bat and edit it.
    pause
    exit /b 1
)
call "%~dp0_settings.bat"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Test-SQLConnect.ps1" ^
    -Server "%SERVER%" ^
    -CompanyDB "%COMPANYDB%" ^
    -DBUser "%DBUSER%" ^
    -DBPassword "%DBPASSWORD%"

echo.
pause
