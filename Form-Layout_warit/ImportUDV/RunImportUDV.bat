@echo off
chcp 65001 >nul
REM ============================================================
REM  Import User-Defined Values (FMS) to SAP B1 via DIRECT SQL.
REM  Replaces the DI API path (Import_UDV_DI.ps1, deprecated)
REM  because DI API v10 doesn't persist QueryId/Refresh/FieldID.
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
REM    -DryRun        = preview only (no writes, no transaction)
REM    (leave empty)  = real import
REM ============================================================
set MODE=

setlocal enabledelayedexpansion
set "CFG=%~dp0Config"

echo ============================================
echo  Select UDV mapping CSV from:
echo  !CFG!
echo ============================================
set IDX=0
for %%F in ("!CFG!\UDV_*.csv") do (
    set /a IDX+=1
    set "FILE_!IDX!=%%~nxF"
    echo   !IDX!. %%~nxF
)
if %IDX%==0 (
    echo No UDV_*.csv files found in !CFG!
    echo Copy UDV_Map.csv.example -^> UDV_Map.csv and edit it.
    pause
    exit /b
)
echo.
set /p PICK=Enter number (1-%IDX%):
if "%PICK%"=="" ( echo No selection. & pause & exit /b )
call set "MAPFILE=%%FILE_%PICK%%%"
if "%MAPFILE%"=="" ( echo Invalid selection. & pause & exit /b )

echo.
echo ============================================
echo  SAP B1 UDV/FMS Import (SQL Direct)
echo  Server    : %SERVER%
echo  Database  : %COMPANYDB%
echo  MapFile   : !MAPFILE!
echo  Mode      : %MODE% (empty=REAL run, writes CSHS+OUQR)
echo ============================================
echo.
echo This writes directly to CSHS (and OUQR for new queries).
echo Press Ctrl+C to abort, or any key to continue.
pause

REM 64-bit PowerShell ? no DI API required for SQL-direct path
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Import_UDV_SQL.ps1" ^
    -Server "%SERVER%" ^
    -CompanyDB "%COMPANYDB%" ^
    -DBUser "%DBUSER%" ^
    -DBPassword "%DBPASSWORD%" ^
    -MapFile "!CFG!\!MAPFILE!" ^
    %MODE%

endlocal

echo.
echo ============================================
echo  Done. Check log: %~dp0Import_UDV_SQL_Log.txt
echo ============================================
pause
