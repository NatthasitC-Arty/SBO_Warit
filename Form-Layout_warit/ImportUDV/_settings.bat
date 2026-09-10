@echo off
REM ============================================================
REM  SHARED CONNECTION SETTINGS  (template) -- ImportUDV
REM ------------------------------------------------------------
REM  Copy this file to _settings.bat (without .example) and edit
REM  the values for your environment. _settings.bat is gitignored.
REM ============================================================

REM --- Database (also used as DI API credentials) ---
set SERVER=10.0.70.61
set COMPANYDB=SBO_SDA_REAL
set DBUSER=sa
set DBPASSWORD=Sa#SemTest2026

REM --- SAP B1 application login (required by DI API) ---
REM  Leave SAPPASSWORD blank to reuse DBPASSWORD.
set SAPUSER=manager
set SAPPASSWORD=1111

REM --- DB engine: MSSQL or HANA ---
set DBTYPE=MSSQL
