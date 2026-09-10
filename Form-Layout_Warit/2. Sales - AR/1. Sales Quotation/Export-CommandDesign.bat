@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

set "RPT_FOLDER=%~dp0"
set "OUTPUT_ROOT=%~dp0"
if "%RPT_FOLDER:~-1%"=="\" set "RPT_FOLDER=%RPT_FOLDER:~0,-1%"
if "%OUTPUT_ROOT:~-1%"=="\" set "OUTPUT_ROOT=%OUTPUT_ROOT:~0,-1%"

set "SCRIPT="

rem 1) Climb parent folders from this .bat looking for the script (finds the repo copy)
set "DIR=%~dp0"
:findloop
if "%DIR:~-1%"=="\" set "DIR=%DIR:~0,-1%"
if exist "%DIR%\Extract-CrystalReport.ps1" set "SCRIPT=%DIR%\Extract-CrystalReport.ps1" & goto found
for %%P in ("%DIR%") do set "PARENT=%%~dpP"
if "%PARENT:~-1%"=="\" set "PARENT=%PARENT:~0,-1%"
if /I not "%PARENT%"=="%DIR%" ( set "DIR=%PARENT%" & goto findloop )

rem 2) Fallbacks for other machines
if exist "%USERPROFILE%\Desktop\Extract-CrystalReport.ps1" set "SCRIPT=%USERPROFILE%\Desktop\Extract-CrystalReport.ps1" & goto found
if exist "C:\SDA\Extract-CrystalReport.ps1" set "SCRIPT=C:\SDA\Extract-CrystalReport.ps1" & goto found

:found
if not defined SCRIPT (
    echo [ERROR] Extract-CrystalReport.ps1 not found.
    echo         Put it in the Form-Layout folder ^(next to this repo^) or on the Desktop.
    pause
    exit /b 1
)

echo [INFO] Using script: %SCRIPT%
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -RptFolder "%RPT_FOLDER%" -OutputRoot "%OUTPUT_ROOT%"

echo.
pause
