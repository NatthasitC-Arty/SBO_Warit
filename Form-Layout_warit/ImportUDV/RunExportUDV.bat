@echo off
chcp 65001 >nul
REM ============================================================
REM  Export all User-Defined Values (FMS) from SAP B1 to CSV.
REM  Uses .NET SqlClient direct (no DI API required).
REM  Output: Config\UDV_Export_<timestamp>.csv (Import-ready).
REM  Connection settings in _settings.bat (shared, gitignored).
REM ============================================================
if not exist "%~dp0_settings.bat" (
    echo ERROR: _settings.bat not found.
    echo Copy _settings.bat.example to _settings.bat and edit it.
    pause
    exit /b 1
)
call "%~dp0_settings.bat"

REM 64-bit PowerShell - no DI API needed for SQL-direct path
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

echo ============================================
echo  Export UDV/FMS to CSV (SQL Direct)
echo  Server    : %SERVER%
echo  Database  : %COMPANYDB%
echo ============================================
echo.

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Export-UDV.ps1" ^
    -Server "%SERVER%" ^
    -CompanyDB "%COMPANYDB%" ^
    -DBUser "%DBUSER%" ^
    -DBPassword "%DBPASSWORD%"

echo.
echo ============================================
echo  Done. Check Config\UDV_Export_*.csv
echo  Log: %~dp0Export_UDV_Log.txt
echo ============================================
pause
