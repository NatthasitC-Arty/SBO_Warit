@echo off
chcp 65001 >nul
REM ============================================================
REM  Mirror FMS rows from one FormID to another via direct SQL.
REM  Use when DI API write path doesn't persist (v10 limitation).
REM
REM  Edit SOURCE / TARGET / ITEM below before running.
REM ============================================================
if not exist "%~dp0_settings.bat" (
    echo ERROR: _settings.bat not found.
    pause
    exit /b 1
)
call "%~dp0_settings.bat"

REM ----------- EDIT THESE -----------
set SOURCE=149
set TARGET=139
set ITEM=38
REM Leave ITEM empty to mirror every item on the form.
REM Set MODE=-DryRun to preview without writing.
set MODE=
REM ----------------------------------

REM Use 64-bit PS (no DI API required for this script)
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

echo ============================================
echo  Mirror FMS via direct SQL
echo  Server    : %SERVER%
echo  Database  : %COMPANYDB%
echo  Source    : FormID = %SOURCE%
echo  Target    : FormID = %TARGET%
echo  ItemID    : %ITEM%  (blank = all items)
echo  Mode      : %MODE% (empty = REAL run)
echo ============================================
echo.
echo This will DELETE existing rows on target %TARGET% (scoped to ItemID
echo if set) and re-INSERT them from source %SOURCE%. Press Ctrl+C to abort.
echo.
pause

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Mirror-FormFMS.ps1" ^
    -Server "%SERVER%" ^
    -CompanyDB "%COMPANYDB%" ^
    -DBUser "%DBUSER%" ^
    -DBPassword "%DBPASSWORD%" ^
    -SourceFormID "%SOURCE%" ^
    -TargetFormID "%TARGET%" ^
    -ItemID "%ITEM%" ^
    %MODE%

echo.
echo ============================================
echo  Done. Check log: %~dp0Mirror_FMS_Log.txt
echo ============================================
pause
