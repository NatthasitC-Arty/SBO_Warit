@echo off
chcp 65001 >nul
if not exist "%~dp0_settings.bat" ( echo ERROR: _settings.bat not found. & pause & exit /b 1 )
call "%~dp0_settings.bat"

REM --- Compare TPLId=34 (manager's local override of UI ALL) vs TPLId=3 (UI ALL master) ---
set FORMID=139
set USERSIGN=1
set LEFT=34
set RIGHT=3

set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Diff-Templates.ps1" ^
    -Server "%SERVER%" -CompanyDB "%COMPANYDB%" -DBUser "%DBUSER%" -DBPassword "%DBPASSWORD%" ^
    -FormID "%FORMID%" -UserSign %USERSIGN% -LeftTPL %LEFT% -RightTPL %RIGHT%

echo.
pause
