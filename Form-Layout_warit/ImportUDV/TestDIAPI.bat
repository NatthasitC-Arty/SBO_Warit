@echo off
chcp 65001 >nul
REM ============================================================
REM  Verify DI API install + B1 login. No data is modified.
REM  Auto-selects PowerShell arch to match installed DI API.
REM ============================================================
if not exist "%~dp0_settings.bat" (
    echo ERROR: _settings.bat not found.
    echo Copy _settings.bat.example to _settings.bat and edit it.
    pause
    exit /b 1
)
call "%~dp0_settings.bat"

if "%SAPUSER%"==""     set SAPUSER=manager
if "%SAPPASSWORD%"=="" set SAPPASSWORD=%DBPASSWORD%
if "%DBTYPE%"==""      set DBTYPE=MSSQL

REM Auto-select PowerShell arch to match installed DI API
set "PS64=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "PS32=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
set "PFX64=%ProgramFiles%"
set "PFX86=%ProgramFiles(x86)%"
set "PS="
if exist "%PFX64%\SAP\SAP Business One DI API\" set "PS=%PS64%"
if defined PS goto :ps_found
if exist "%PFX86%\SAP\SAP Business One DI API\" set "PS=%PS32%"
if defined PS goto :ps_found
echo ERROR: DI API not detected.
echo   Looked in: %PFX64%\SAP\SAP Business One DI API\
echo   Looked in: %PFX86%\SAP\SAP Business One DI API\
pause
exit /b 2
:ps_found
echo Using PowerShell: %PS%

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Test-DIAPI.ps1" ^
    -Server "%SERVER%" ^
    -CompanyDB "%COMPANYDB%" ^
    -DBUser "%DBUSER%" ^
    -DBPassword "%DBPASSWORD%" ^
    -SapUser "%SAPUSER%" ^
    -SapPassword "%SAPPASSWORD%" ^
    -DBType "%DBTYPE%"

pause
