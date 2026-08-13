@echo off
chcp 65001 >nul
REM ============================================================
REM  Import Crystal Layouts to SAP B1 (SQL Direct)
REM  Connection settings are in _settings.bat (shared, gitignored).
REM
REM  Choose at runtime:
REM    1 = Import ALL rows from the mapping file
REM    2 = Import by KEYWORD (filter by RPT_FileName, loop)
REM ============================================================
if not exist "%~dp0_settings.bat" (
    echo ERROR: _settings.bat not found.
    echo Copy _settings.bat.example to _settings.bat and edit it.
    pause
    exit /b 1
)
call "%~dp0_settings.bat"

REM ============================================================
REM  AUTHOR: stored on NEW rows only (INSERT). Updates to existing
REM  layouts keep the original Author -- dedup matches across authors.
REM ============================================================
set AUTHOR=manager

REM ============================================================
REM  MODE: choose one of these for MODE
REM    -DryRun        = preview only (no changes)
REM    (leave empty)  = real import
REM ============================================================
set MODE=

REM ============================================================
REM  ONDUP: how to handle duplicate layouts (DocName+TypeCode match, Author ignored)
REM    Update  = overwrite existing (recommended, default)
REM    Skip    = leave existing alone, only insert new
REM    Insert  = always insert new row (creates duplicates - careful!)
REM ============================================================
set ONDUP=Update

REM ============================================================
REM  CONFIGDIR: folder to scan for mapping .xlsx files
REM  Can be relative (Config) or absolute (C:\path\...)
REM ============================================================
set CONFIGDIR=Config

setlocal enabledelayedexpansion
REM Resolve CONFIGDIR / RPTROOT to absolute paths
set "CFG=!CONFIGDIR!"
if not "!CFG:~1,1!"==":" set "CFG=%~dp0!CFG!"
set "RPT=!RPTROOT!"
if not "!RPT:~1,1!"==":" set "RPT=%~dp0!RPT!"

echo ============================================
echo  Select mapping Excel file from:
echo  !CFG!
echo ============================================
set IDX=0
for %%F in ("!CFG!\*.xlsx") do (
    set /a IDX+=1
    set "FILE_!IDX!=%%~nxF"
    echo   !IDX!. %%~nxF
)
if %IDX%==0 (
    echo No .xlsx files found in !CFG!
    pause
    exit /b
)
echo.
set /p PICK=Enter number (1-%IDX%):
if "%PICK%"=="" (
    echo No selection. Exiting.
    pause
    exit /b
)
call set "MAPFILE=%%FILE_%PICK%%%"
if "%MAPFILE%"=="" (
    echo Invalid selection.
    pause
    exit /b
)

echo.
echo ============================================
echo  Choose import mode:
echo    1. Import ALL rows in mapping file
echo    2. Import by KEYWORD (filter by RPT_FileName, loop)
echo ============================================
set /p MODECHOICE=Enter 1 or 2:

if "%MODECHOICE%"=="1" goto IMPORT_ALL
if "%MODECHOICE%"=="2" goto IMPORT_KEYWORD
echo Invalid selection.
pause
endlocal
exit /b

:IMPORT_ALL
echo.
echo ============================================
echo  SAP B1 Layout Import (ALL)
echo  Server   : %SERVER%
echo  Database : %COMPANYDB%
echo  MapFile  : !MAPFILE!
echo  RptRoot  : !RPT!
echo  Mode     : %MODE% (empty=real run)
echo  OnDup    : %ONDUP%
echo ============================================
echo.
pause

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Import_SQL_Direct.ps1" ^
    -Server "%SERVER%" ^
    -CompanyDB "%COMPANYDB%" ^
    -DBUser "%DBUSER%" ^
    -DBPassword "%DBPASSWORD%" ^
    -Author "%AUTHOR%" ^
    -MapFile "!CFG!\!MAPFILE!" ^
    -RptRoot "!RPT!" ^
    -UseFileNameAsDocName ^
    -OnDuplicate %ONDUP% ^
    %MODE%

goto END

:IMPORT_KEYWORD
echo.
echo ============================================
echo  SAP B1 Layout Import (BY KEYWORD)
echo  Server   : %SERVER%
echo  Database : %COMPANYDB%
echo  MapFile  : !MAPFILE!
echo  RptRoot  : !RPT!
echo  Mode     : %MODE% (empty=real run)
echo  OnDup    : %ONDUP%
echo ============================================

:KEYWORD_LOOP
echo.
set "FILTER="
set /p FILTER=Type keyword (e.g. Sale Order) -- empty Enter to quit:

if "%FILTER%"=="" goto END

echo.
echo Importing rows matching "%FILTER%" ...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Import_SQL_Direct.ps1" ^
    -Server "%SERVER%" ^
    -CompanyDB "%COMPANYDB%" ^
    -DBUser "%DBUSER%" ^
    -DBPassword "%DBPASSWORD%" ^
    -Author "%AUTHOR%" ^
    -MapFile "!CFG!\!MAPFILE!" ^
    -RptRoot "!RPT!" ^
    -FilterFileName "%FILTER%" ^
    -UseFileNameAsDocName ^
    -OnDuplicate %ONDUP% ^
    %MODE%

echo.
echo ============================================
echo  Done. Type next keyword, or empty Enter to quit.
echo ============================================
goto KEYWORD_LOOP

:END
endlocal

echo.
echo ============================================
echo  Done. Check log: %~dp0Import_SQL_Log.txt
echo ============================================
pause
