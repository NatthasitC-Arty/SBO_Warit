@echo off
chcp 65001 >nul
REM ============================================================
REM  Rollback Crystal Layouts: list non-system layouts, pick which to delete.
REM  Connection settings are in _settings.bat (shared, gitignored).
REM ============================================================
if not exist "%~dp0_settings.bat" (
    echo ERROR: _settings.bat not found.
    echo Copy _settings.bat.example to _settings.bat and edit it.
    pause
    exit /b 1
)
call "%~dp0_settings.bat"

REM ============================================================
REM  MODE:
REM    -DryRun        = preview selection only (no delete)
REM    (leave empty)  = delete for real (will ask 'yes' to confirm)
REM    -Force         = delete without confirmation
REM ============================================================
set MODE=

REM ============================================================
REM  SYSTEMAUTHOR: rows with this Author are protected (always kept).
REM  Default 'System' matches SAP B1 v10. Verify with:
REM    SELECT DISTINCT Author FROM RDOC
REM ============================================================
set SYSTEMAUTHOR=System

echo ============================================
echo  Rollback Layouts (by selection)
echo  Server   : %SERVER%
echo  Database : %COMPANYDB%
echo  KeepAuth : %SYSTEMAUTHOR%
echo  Mode     : %MODE%
echo ============================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Rollback-BySelection.ps1" ^
    -Server "%SERVER%" ^
    -CompanyDB "%COMPANYDB%" ^
    -DBUser "%DBUSER%" ^
    -DBPassword "%DBPASSWORD%" ^
    -SystemAuthor "%SYSTEMAUTHOR%" ^
    %MODE%

echo.
pause
