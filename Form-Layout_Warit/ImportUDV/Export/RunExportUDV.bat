@echo off
chcp 65001 >nul
REM ============================================================
REM  Export User-Defined Values (FMS) from SAP B1 to CSV.
REM  Uses .NET SqlClient direct (no DI API required).
REM  Output: Config\UDV_Export_<timestamp>.csv (Import-ready).
REM  Connection settings in _settings.bat (shared, gitignored).
REM ============================================================
if not exist "%~dp0..\_settings.bat" (
    echo ERROR: _settings.bat not found.
    echo Copy _settings.bat.example to _settings.bat and edit it.
    pause
    exit /b 1
)
call "%~dp0..\_settings.bat"

echo ============================================
echo  Export UDV/FMS to CSV  (SQL Direct)
echo  Server   : %SERVER%
echo  Database : %COMPANYDB%
echo ============================================
echo.
echo  FormIDs (verified - sap_b1_form_ids.json):
echo  --- Sales / AR ---       --- Purchasing / AP ---
echo   149  Sales Quotation     142  Purchase Order
echo   139  Sales Order         143  Goods Receipt PO
echo   140  Delivery            141  A/P Invoice
echo   133  A/R Invoice         181  A/P Credit Memo
echo   179  A/R Credit Memo
echo  --- Master / Finance / Inventory ---
echo   134  Business Partner    392  Journal Entry
echo   150  Item Master         170  Incoming Payment
echo   720  Goods Issue         426  Outgoing Payment
echo   721  Goods Receipt       940  Inventory Transfer
echo.
echo  (full list: Config\FormID_Reference.md)
echo  (or use Ctrl+Shift+I on any form to find its FormID)
echo.
echo  Leave FormID blank to export ALL forms.
echo ============================================
echo.

set "FORMID="
set "ITEMID="
set "COLID="
set /p FORMID=FormID filter [blank = all]:
if defined FORMID goto :ask_more
goto :run

:ask_more
set /p ITEMID=ItemID filter [blank = all items]:
set /p COLID=ColumnID filter [blank = all columns]:

:run
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

echo.
echo Exporting... FormID=[%FORMID%]  ItemID=[%ITEMID%]  ColumnID=[%COLID%]
echo.

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0Scripts\Export-UDV.ps1" ^
    -Server "%SERVER%" ^
    -CompanyDB "%COMPANYDB%" ^
    -DBUser "%DBUSER%" ^
    -DBPassword "%DBPASSWORD%" ^
    -FormID "%FORMID%" ^
    -ItemID "%ITEMID%" ^
    -ColumnID "%COLID%"

echo.
echo ============================================
echo  Done. Check Config\UDV_Export_*.csv
echo  Log: %~dp0Export_UDV_Log.txt
echo ============================================
pause
