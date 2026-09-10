@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul

rem ---------------------------------------------------------------------------
rem  Fix-ReportSql.bat -- menu wrapper for Fix-ReportSql.ps1
rem
rem  Pass-through mode: any argument given here goes straight to the .ps1, e.g.
rem      Fix-ReportSql.bat -Path "C:\GitHub\SDA\Form-Layout_SDA" -Approver -WhatIf
rem  With no arguments it asks the questions below.
rem
rem  ORDER OF WORK -- do NOT swap these two steps:
rem    1. run every SQL fix here first
rem    2. THEN do one Crystal Designer pass:
rem       Database > Set Datasource Location > Update > Ctrl+S
rem  Each SQL write re-injects 'Database DLL' into the connection, which is what
rem  stops SAP B1 opening a layout, and no offline edit can remove it. Doing the
rem  Designer pass first just throws that work away.
rem ---------------------------------------------------------------------------

set "PS1=%~dp0Fix-ReportSql.ps1"
if not exist "%PS1%" (
    echo [ERROR] Not found: %PS1%
    pause
    exit /b 4
)

set "PSEXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

rem pause must not sit between the call and the exit -- it overwrites ERRORLEVEL
if not "%~1"=="" (
    "%PSEXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %*
    set "RC=!ERRORLEVEL!"
    echo.
    if "!RC!"=="0" (echo Done -- exit code 0.) else (echo [WARN] exit code !RC!)
    exit /b !RC!
)

echo ===========================================================
echo  Fix-ReportSql -- repair SQL Commands inside .rpt layouts
echo ===========================================================
echo.
echo Default folder: %~dp0
set "TARGET=%~dp0"
set /p "TARGET=Folder or single .rpt [%TARGET%]: "

echo.
echo Which fixes?
echo   1 = approver box shows the wrong round      (TOP 1 + ORDER BY)
echo   2 = draft documents print duplicated lines  (pj join -^> OUTER APPLY)
echo   3 = reports show the project CODE only      (add a ProjectName column)
echo   4 = 1 + 2 + 3                               (recommended)
echo   5 = all of the above, plus AR Down Payment OWDD.ObjType '23' -^> '203'
set "PICK=4"
set /p "PICK=Choice [4]: "

set "SW="
if "%PICK%"=="1" set "SW=-Approver"
if "%PICK%"=="2" set "SW=-ProjectJoin"
if "%PICK%"=="3" set "SW=-ProjectName"
if "%PICK%"=="4" set "SW=-Approver -ProjectJoin -ProjectName"
if "%PICK%"=="5" set "SW=-Approver -ProjectJoin -ProjectName -DownPaymentObjType"
if "!SW!"=="" (
    echo [ERROR] Choice must be 1-5.
    pause
    exit /b 4
)

echo.
echo Mode:
echo   1 = WhatIf -- report only, write nothing  (recommended first)
echo   2 = Apply  -- write the files
set "MODE=1"
set /p "MODE=Choice [1]: "
set "WI="
if "%MODE%"=="1" set "WI=-WhatIf"

echo.
echo   Path       : %TARGET%
echo   Transforms : !SW!
if defined WI (echo   Mode       : WHATIF -- nothing will be written) else (echo   Mode       : APPLY -- files WILL be rewritten, backups are kept)
echo.
set "GO=n"
set /p "GO=Proceed? [y/N]: "
if /i not "!GO!"=="y" (
    echo Cancelled.
    pause
    exit /b 0
)

echo.
"%PSEXE%" -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Path "%TARGET%" !SW! !WI!
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
    if not defined WI (
        echo Done. Now open each changed report in Crystal Designer and do
        echo   Database ^> Set Datasource Location ^> Update ^> Ctrl+S
        echo Backups sit next to each file with a .bak suffix.
    )
) else (
    echo [WARN] exit code %RC% -- check the log for rolled-back files.
)
echo.
pause
exit /b %RC%
