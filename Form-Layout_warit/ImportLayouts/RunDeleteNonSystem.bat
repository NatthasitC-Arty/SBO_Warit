@echo off
chcp 65001 >nul
REM ============================================================
REM  Delete ALL layouts in RDOC EXCEPT Author='System'
REM  ** DESTRUCTIVE ** - backup RDOC first if needed
REM  Connection settings are in _settings.bat (shared, gitignored).
REM ============================================================
if not exist "%~dp0_settings.bat" (
    echo ERROR: _settings.bat not found.
    echo Copy _settings.bat.example to _settings.bat and edit it.
    pause
    exit /b 1
)
call "%~dp0_settings.bat"

set SYSTEMAUTHOR=System

REM ============================================================
REM  MODE:
REM    -DryRun        = preview only (no delete)
REM    (leave empty)  = delete for real (will ask 'yes' confirm)
REM    -Force         = delete without asking
REM ============================================================
set MODE=

echo ============================================
echo  Delete NON-SYSTEM Layouts from RDOC
echo  Server      : %SERVER%
echo  Database    : %COMPANYDB%
echo  Keep Author : %SYSTEMAUTHOR%
echo  Mode        : %MODE%
echo ============================================
echo.
pause

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Delete-NonSystemLayouts.ps1" ^
    -Server "%SERVER%" ^
    -CompanyDB "%COMPANYDB%" ^
    -DBUser "%DBUSER%" ^
    -DBPassword "%DBPASSWORD%" ^
    -SystemAuthor "%SYSTEMAUTHOR%" ^
    %MODE%

echo.
pause
