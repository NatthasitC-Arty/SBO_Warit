@echo off
chcp 65001 >nul
REM ============================================================
REM  Convert between annotated XLSX and plain CSV.
REM  Requires Microsoft Excel installed.
REM
REM  Modes:
REM    1 = CSV -> XLSX (add column comments)
REM    2 = XLSX -> CSV (flatten back to CSV for Import)
REM ============================================================
setlocal enabledelayedexpansion
set "CFG=%~dp0Config"
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

echo ============================================
echo  Convert UDV files
echo  Config: !CFG!
echo ============================================
echo   1. CSV  -^> XLSX  (add column comments)
echo   2. XLSX -^> CSV   (flatten for Import)
echo.
set /p MODE=Enter mode (1 or 2):

if "%MODE%"=="1" goto :MODE_CSV2XLSX
if "%MODE%"=="2" goto :MODE_XLSX2CSV
echo Invalid mode.
pause
exit /b

:MODE_CSV2XLSX
echo.
echo === CSV -^> XLSX ===
set IDX=0
for %%F in ("!CFG!\UDV_*.csv") do (
    set /a IDX+=1
    set "FILE_!IDX!=%%~nxF"
    echo   !IDX!. %%~nxF
)
if %IDX%==0 ( echo No UDV_*.csv files in !CFG! & pause & exit /b )
echo.
set /p PICK=Enter number (1-%IDX%):
if "%PICK%"=="" ( echo No selection. & pause & exit /b )
call set "FILE=%%FILE_%PICK%%%"
if "%FILE%"=="" ( echo Invalid selection. & pause & exit /b )

echo Converting !FILE! -^> .xlsx
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Convert-CsvToAnnotatedXlsx.ps1" -InputCsv "!CFG!\!FILE!"
goto :END

:MODE_XLSX2CSV
echo.
echo === XLSX -^> CSV ===
set IDX=0
for %%F in ("!CFG!\UDV_*.xlsx") do (
    set /a IDX+=1
    set "FILE_!IDX!=%%~nxF"
    echo   !IDX!. %%~nxF
)
if %IDX%==0 ( echo No UDV_*.xlsx files in !CFG! & pause & exit /b )
echo.
set /p PICK=Enter number (1-%IDX%):
if "%PICK%"=="" ( echo No selection. & pause & exit /b )
call set "FILE=%%FILE_%PICK%%%"
if "%FILE%"=="" ( echo Invalid selection. & pause & exit /b )

echo Converting !FILE! -^> .csv
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Convert-XlsxToCsv.ps1" -InputXlsx "!CFG!\!FILE!"
goto :END

:END
echo.
pause
