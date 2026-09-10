@echo off
REM ============================================================
REM  CONNECTION SETTINGS -- Manage UI (Form Settings)
REM ------------------------------------------------------------
REM  Copy this file to _settings.bat (without .example) and fill
REM  in your environment values. _settings.bat is gitignored.
REM ============================================================

REM --- Database ---
set SERVER=10.10.10.115
set COMPANYDB=SBO_Update_UI
set DBUSER=sa
set DBPASSWORD=1q2w3e4r@

REM --- DB engine: MSSQL or HANA ---
set DBTYPE=MSSQL
