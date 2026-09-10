@echo off
chcp 65001 >nul
REM ============================================================
REM  Mirror CPRF rows: source (UserSign+TPLId) -> target (UserSign+TPLId)
REM  for a specific FormID. Transaction-wrapped.
REM
REM  Default: push manager's local edit (UserSign=1, TPLId=34) to
REM  the UI ALL master (UserSign=26, TPLId=3) for FormID=139.
REM ============================================================
if not exist "%~dp0_settings.bat" ( echo ERROR: _settings.bat not found. & pause & exit /b 1 )
call "%~dp0_settings.bat"

set FORMID=139
set SRC_USER=1
set SRC_TPL=34
set TGT_USER=1
set TGT_TPL=3

REM Toggle: -DryRun = preview only, blank = real run
set MODE=
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Mirror-FormUI.ps1" ^
    -Server "%SERVER%" -CompanyDB "%COMPANYDB%" -DBUser "%DBUSER%" -DBPassword "%DBPASSWORD%" ^
    -FormID "%FORMID%" ^
    -SourceUserSign %SRC_USER% -SourceTPLId %SRC_TPL% ^
    -TargetUserSign %TGT_USER% -TargetTPLId %TGT_TPL% ^
    %MODE%

echo.
echo Log: %~dp0Mirror_FormUI_Log.txt
pause
