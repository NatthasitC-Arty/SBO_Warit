# SAP B1 Crystal Layout Batch Import

เครื่องมือสำหรับ **import ไฟล์ Crystal Reports (.rpt) จำนวนมาก** เข้า SAP Business One แบบ batch ผ่าน **SQL Direct** (INSERT/UPDATE ตรงเข้า table `RDOC`)

---

## สารบัญ

1. [ทำอะไรได้บ้าง](#ทำอะไรได้บ้าง)
2. [Requirements](#requirements)
3. [โครงสร้างไฟล์](#โครงสร้างไฟล์)
4. [Setup ครั้งแรก](#setup-ครั้งแรก)
5. [วิธีใช้งาน](#วิธีใช้งาน)
6. [การแก้ไข Mapping (Excel)](#การแก้ไข-mapping-excel)
7. [เปลี่ยน Server / ย้ายไป Client ใหม่](#เปลี่ยน-server--ย้ายไป-client-ใหม่)
8. [Troubleshooting](#troubleshooting)
9. [Technical Details](#technical-details)
10. [ข้อควรระวัง](#ข้อควรระวัง)

---

## ทำอะไรได้บ้าง

- ✅ Import Crystal Report layouts หลายไฟล์พร้อมกันจาก Excel mapping
- ✅ UPDATE layout ที่มีอยู่แล้ว (overwrite) หรือ INSERT ใหม่
- ✅ Filter เฉพาะ keyword (เช่น import เฉพาะ "Journal Entry")
- ✅ DryRun preview ก่อน run จริง
- ✅ Backup + Rollback อัตโนมัติ
- ✅ Log ทุกการกระทำ
- ✅ ใช้ได้กับ SAP B1 v10 บน MSSQL (2017/2019/2022)

---

## Requirements

| รายการ | Version |
|--------|---------|
| Windows | 10 / 11 / Server 2016+ |
| PowerShell | 5.1 (built-in) |
| SQL Server | 2017/2019/2022 |
| SAP Business One | v10.0 |
| Microsoft Excel | Installed (ใช้ COM อ่าน `.xlsx`) |
| SQL User | ต้องมีสิทธิ์ INSERT/UPDATE/DELETE ที่ table `RDOC` |

> **Note:** ไม่ต้องลง SAP DI API, Interop, Service Layer — ใช้ `System.Data.SqlClient` ใน .NET Framework

### 🖥️ รันได้ทั้งบน Server และ Client

เครื่องมือนี้ไม่ผูกกับเครื่อง SAP Server — **รันจากเครื่องไหนก็ได้** ที่เข้าเงื่อนไขนี้

| ตัวเลือก | ต้องมี | เหมาะกับใคร |
|---------|--------|-------------|
| **รันบน SQL Server** | — | admin/DBA ที่ remote ถึง DB host |
| **รันบน SAP B1 Client** ⭐ | Network ถึง SQL Server port 1433 | consultant/implementer ทั่วไป (แนะนำ) |
| **รันบนเครื่อง PC/Laptop อื่น** | Network ถึง SQL Server + ติด Excel | ทำงาน remote |

**ไม่ต้องติดตั้ง SAP บนเครื่องที่รัน script** — เพราะ script คุยกับ SQL ตรงๆ ไม่คุยกับ SAP Server/Service Layer

**สิ่งเดียวที่ต้องมีในเครื่องรัน:**
1. ✅ PowerShell 5.1 (Windows built-in)
2. ✅ Microsoft Excel (สำหรับอ่าน `RPT_Import_Map.xlsx` ผ่าน COM)
3. ✅ Network access → SQL Server (port 1433)
4. ✅ Credentials ของ SQL login ที่มีสิทธิ์ write ที่ table `RDOC`

**ไฟล์ `.rpt` ทั้งหมด** วางไว้ที่ไหนก็ได้ — แค่ `RPT_Import_Map.xlsx` column `RPT_Folder` ชี้ path ให้ถูก (หรือส่ง parameter `-RptRoot`)

---

## โครงสร้างไฟล์

```
ImportLayouts/
│
├── 📁 Backups/                     ← backup DB (สร้างอัตโนมัติ)
│
├── 🔑 _settings.bat                ← ⭐ shared connection settings (gitignored)
├── 🔑 _settings.bat.example        ← template for new installs (committed)
│
├── 🔧 .bat files (double-click ได้)
│   ├── TestConnect.bat             ← ทดสอบต่อ DB
│   ├── RunImport.bat               ← import จาก Excel (เลือก: ทั้งหมด / ตาม keyword)
│   ├── RunRollback.bat             ← ลบที่ import จาก Excel
│   └── RunDeleteNonSystem.bat      ← ลบทุก non-system layout
│
├── ⚙️ PowerShell scripts
│   ├── Test-SQLConnect.ps1         ← logic ทดสอบ DB
│   ├── Backup-RDOC.ps1             ← backup table RDOC
│   ├── Import_SQL_Direct.ps1       ← ⭐ script หลัก
│   ├── Rollback-BySelection.ps1    ← logic rollback (list + เลือก)
│   ├── Delete-NonSystemLayouts.ps1 ← logic delete non-system
│   └── Restore-SystemLayouts.ps1   ← restore system จาก DB อื่น
│
├── 🔌 DB Plugin
│   └── DB-MSSQL.ps1                ← abstraction layer สำหรับ MSSQL
│
├── 📊 Data
│   └── Config/RPT_Import_Map.xlsx  ← mapping table
│
├── 📋 Log
│   └── Import_SQL_Log.txt          ← log ทุก import (สร้างอัตโนมัติ, gitignored)
│
└── 🛡️ .gitignore                   ← exclude _settings.bat + log + backups
```

---

## Setup ครั้งแรก

### 1. Copy ทั้งโฟลเดอร์ไปเครื่องที่จะรัน

```
C:\SDA\SDA\Form-Layout\ImportLayouts\
```

**รันบนเครื่องไหนก็ได้** — SAP Server, SQL Server, เครื่อง Client ที่ใช้ SAP B1, หรือ Laptop ของ implementer ก็ได้ ขอแค่:
- Network ถึง SQL Server (ping + TCP 1433)
- ติดตั้ง Excel
- ไฟล์ `.rpt` ทั้งหมดเข้าถึงได้จากเครื่องนั้น (local path หรือ shared folder)

**แนะนำรันบนเครื่อง Client** เพราะสะดวก + ทดสอบ preview ใน SAP UI ต่อได้ทันที

### 2. สร้าง `_settings.bat` (ตั้งครั้งเดียว ใช้กับทุก .bat)

```cmd
copy _settings.bat.example _settings.bat
```

แล้วเปิด `_settings.bat` ด้วย **Notepad** แก้ 5 บรรทัดให้ตรงกับ environment ของคุณ:

```bat
set SERVER=10.10.10.115            ← IP หรือชื่อ SQL Server
set COMPANYDB=SBO_SDA_MARK1        ← ชื่อ Company DB
set DBUSER=sa                      ← SQL user (ต้องมีสิทธิ์ INSERT ใน RDOC)
set DBPASSWORD=YourPassword        ← password
set RPTROOT=C:\SDA\SDA\Form-Layout ← root folder ที่มีไฟล์ .rpt
```

> 📌 **ทุก `Run*.bat` `call _settings.bat` อัตโนมัติ** — เปลี่ยน server/DB/password ที่ไฟล์เดียว ไม่ต้องไปแก้ทุก .bat
>
> 🔒 **`_settings.bat` ถูก gitignored** — ไม่ขึ้น git ไม่หลุดไปเครื่องอื่น (share เฉพาะ `_settings.bat.example` ที่เป็น template)

ใน `Run*.bat` แต่ละไฟล์มีตั้งค่าเฉพาะของตัวเองเพิ่มเติม (บนสุด หลัง `call _settings.bat`):
```bat
set AUTHOR=manager       ← RunImport / RunRollback
set MODE=                ← -DryRun / ว่าง / -Force
set ONDUP=Update         ← RunImport เท่านั้น
```

### 3. ทดสอบ connection

**ดับเบิลคลิก `TestConnect.bat`**

ต้องเห็น:
```
[1/4] Pinging 10.10.10.115 ...     Ping: OK
[2/4] Testing SQL port 1433 ...    TCP 1433: OPEN
[3/4] Testing SQL connection ...   SQL Login: OK (server 15.00.2165)
[4/4] Counting layouts in RDOC ... Total: 662  Crystal: 197
READY TO IMPORT
```

ถ้า FAIL → ดู [Troubleshooting](#troubleshooting)

---

## วิธีใช้งาน

### Workflow มาตรฐาน (import ทั้งหมดจาก Excel)

```
1. ตรวจ TestConnect.bat                  ← ต้องเห็น "READY TO IMPORT"
2. .\Backup-RDOC.ps1 -Server ... -CompanyDB ... -DBPassword ...
3. แก้ MODE=-DryRun ใน RunImport.bat → ดับเบิลคลิก (ดู preview)
4. แก้ MODE=     → ดับเบิลคลิก RunImport.bat (run จริง)
5. Verify ใน SAP B1 Client
```

---

### 🔹 Workflow A: Import ทั้งหมด (batch 57+ layouts)

**เปิด `RunImport.bat` ด้วย Notepad ตั้งค่า:**
```bat
set MODE=-DryRun       REM ครั้งแรก preview ก่อน
set ONDUP=Update       REM overwrite ของเดิมที่ชื่อซ้ำ
```

**ดับเบิลคลิก** → เลือกไฟล์ Excel → ตอบ `1` ที่หน้าเมนู (Import ALL) → ดูผล

ถ้า OK แก้เป็น:
```bat
set MODE=               REM ว่าง = run จริง
```

**ดับเบิลคลิกอีกครั้ง** → ตอบ `1` → รอ 10-15 วินาที

**ผลลัพธ์ที่คาดหวัง:**
```
UPDATE [  1] 1.Journal Entry...       -> DocCode=JDT20002
UPDATE [  2] 2.Sale Quotation_(Bom)... -> DocCode=QUT20003
INSERT [  3] 2.Sale Quotation BOM EN...-> DocCode=QUT20008
...
=== Summary: OK=57 FAIL=0 SKIP=6 ===
```

**ความหมาย:**
- `UPDATE` = มี layout ชื่อเดียวกันอยู่แล้ว → overwrite content
- `INSERT` = ชื่อใหม่ → สร้าง row ใหม่
- `SKIP` = rows ที่ไม่มี ObjectType หรือ unmapped

---

### 🔹 Workflow B: Import ทีละไฟล์ (ตาม keyword)

**ดับเบิลคลิก `RunImport.bat`** → เลือกไฟล์ Excel → ตอบ `2` ที่หน้าเมนู (Import by KEYWORD) → พิมพ์ keyword แล้ว Enter

```
Type keyword (e.g. Sale Order) -- empty Enter to quit: Journal Entry
```

จะ import **เฉพาะ row ที่ RPT_FileName มี keyword นั้น** แล้วถาม keyword ถัดไปวนไปเรื่อยๆ — กด Enter ว่างๆ เพื่อจบ

**ตัวอย่าง keywords:**
| พิมพ์ | Import กี่ไฟล์ |
|------|----------------|
| `Journal Entry` | 1 |
| `Sale Order` | 2 (BOM + Dis) |
| `Sale Quotation` | 4 |
| `AR Invoice` | 8 |
| `Sale Order_ใบสั่งขาย_(Bom)` | 1 (เจาะจง) |

---

### 🔹 Workflow C: Rollback (เลือก layout ที่จะลบ)

**ดับเบิลคลิก `RunRollback.bat`** → script จะ list layout ทั้งหมดที่ไม่ใช่ system (`Author<>'System'`) ให้เลือก

```
Filter by DocName/Author keyword (empty=show all): [Enter ว่าง = แสดงทั้งหมด, หรือพิมพ์ keyword กรองก่อน]

=== 42 layout(s) found ===
 #  DocCode  TypeCode Author  DocName
 -  -------  -------- ------  -------
 1  INV20007 INV2     manager AR Invoice - SDA
 2  INV20008 INV2     manager AR Invoice - SDA EN
 3  RDR20009 RDR2     SDA     Sale Order_ใบสั่งขาย_(Bom)
 ...

Enter selection: 1,3,5
   หรือ  1-10
   หรือ  1-5,8,12-15
   หรือ  all
   หรือ  Enter ว่าง = ยกเลิก
```

ตอบยืนยัน `yes` ก่อน script จะลบ (พร้อม children ใน RITM/RDC1/RCON และ DFLT_PRNTING orphans, ทั้งหมดอยู่ใน transaction เดียวกัน)

**Options ใน `RunRollback.bat`:**
```bat
set MODE=-DryRun        REM preview การเลือกอย่างเดียว ไม่ลบจริง
set MODE=               REM ลบจริง ถาม "yes" ก่อน
set MODE=-Force         REM ลบทันที ไม่ถาม
set SYSTEMAUTHOR=System REM Author ที่ป้องกันไว้ (เช็คด้วย SELECT DISTINCT Author FROM RDOC)
```

**ความปลอดภัย:** layout ที่ `Author='System'` (ของ SAP) จะไม่ถูกแสดงในลิสต์เลย ลบไม่ได้ผ่าน flow นี้

---

### 🔹 Workflow D: Backup

```powershell
cd C:\SDA\SDA\Form-Layout\ImportLayouts
.\Backup-RDOC.ps1 -Server "10.10.10.115" -CompanyDB "SBO_SDA_MARK1" -DBPassword "YourPassword"
```

**ที่เก็บ:** `Backups/RDOC_Backup_<timestamp>.bak` (binary) + `.csv` (index)

---

## การแก้ไข Mapping (Excel)

### ไฟล์: `RPT_Import_Map.xlsx` Sheet: `RPT_MAP`

| Column | ชื่อ | ตัวอย่าง | สำคัญไหม? |
|--------|-----|---------|-----------|
| A | No | 1 | แค่ลำดับ |
| B | Module | Financials | reference |
| C | **RPT_FileName** | `1.AR Invoice.rpt` | ⭐ ต้องตรง filename จริง |
| D | **RPT_Folder** | `4. Sales/1.AR Invoice` | ⭐ path relative จาก RptRoot |
| E | SAP_Document | ใบแจ้งหนี้ | reference |
| F | HeaderTable | OINV | reference |
| G | LineTable | INV1 | reference |
| H | **ObjectType** | `13` | ⭐ script ใช้ map เป็น TypeCode |
| I | FormMenuUID | `133` | ⭐ reference (script ไม่ใช้ตอนนี้) |
| J | **LayoutName** | `AR Invoice - SDA` | ⭐ ใช้เป็น DocName (ถ้าไม่ใช้ `-UseFileNameAsDocName`) |
| K | Note | (free text) | comment |

**คอลัมน์ที่ script ใช้จริง:** C, D, H, J

### ตาราง ObjectType → TypeCode (รู้จักใน script)

| ObjectType | Form | TypeCode |
|-----------|------|----------|
| 13 | AR Invoice | INV2 |
| 14 | AR Credit Memo | RIN2 |
| 15 | Delivery | DLN2 |
| 16 | Returns | RDN2 |
| 17 | Sales Order | RDR2 |
| 18 | AP Invoice | PCH2 |
| 19 | AP Credit Memo | RPC2 |
| 20 | GRPO | PDN2 |
| 21 | Goods Return | RPD2 |
| 22 | Purchase Order | POR2 |
| 23 | Sales Quotation | QUT2 |
| 24 | Incoming Payment | RCT1 |
| 30 | Journal Entry | JDT2 |
| 46 | Outgoing Payment | VPM1 |
| 59 | Goods Receipt | IGN1 |
| 60 | Goods Issue | IGE1 |
| 67 | Inv Transfer | WTR1 |
| 69 | Landed Costs | IPF1 |
| 202 | Production Order | WOR1 |
| 203 | AR Down Payment | DPI2 |
| 204 | AP Down Payment | DPO2 |
| 540000405 | Purchase Quotation | PQT2 |
| 1250000001 | Inv Transfer Request | WTQ1 |
| 1470000065 | Inv Counting | INC1 |
| 1470000113 | Purchase Request | PRQ2 |
| 162 | Inv Revaluation | ❌ unmapped (skip) |

**ถ้าต้องการเพิ่ม ObjectType:** แก้ hash map `$TypeCodeMap` ใน `Import_SQL_Direct.ps1` (Rollback ใช้การเลือกจาก DocCode ตรงๆ ไม่ต้องแก้)

---

## เปลี่ยน Server / ย้ายไป Client ใหม่

### Step 1: Copy ทั้งโฟลเดอร์

```
C:\SDA\SDA\Form-Layout\ImportLayouts\
```

### Step 2: สร้าง / แก้ `_settings.bat` (ที่เดียว)

```cmd
copy _settings.bat.example _settings.bat
notepad _settings.bat
```

แก้ 5 บรรทัด:
```bat
set SERVER=...
set COMPANYDB=...
set DBUSER=...
set DBPASSWORD=...
set RPTROOT=...
```

ทุก `Run*.bat` ใช้ค่าจาก `_settings.bat` อัตโนมัติ

### Step 3: แก้ Excel (ถ้า path .rpt เปลี่ยน)

แก้ `RPTROOT` ใน `_settings.bat` หรือคอลัมน์ `RPT_Folder` ใน `RPT_Import_Map.xlsx`

### Step 4: Test → Backup → Import

```
1. TestConnect.bat → READY TO IMPORT
2. .\Backup-RDOC.ps1 -Server ... -CompanyDB ... -DBPassword ...
3. RunImport.bat (DryRun → Run จริง)
```

---

## Troubleshooting

### `Login failed for user 'sa'`
- ตรวจ `DBPASSWORD` ถูกต้อง
- ลอง login SSMS ด้วย user เดียวกันยืนยัน
- Special chars ใน password → ใส่ในเครื่องหมาย `"..."` ใน `.bat`

### `Cannot open database "XXX"`
- ตรวจชื่อ `COMPANYDB` ตรง (case-sensitive บาง config)
- Run query: `SELECT name FROM sys.databases` ดู list DBs

### `A network-related error occurred`
- SQL Server ไม่ทำงาน / firewall block
- เช็ค: `Test-NetConnection -ComputerName <server> -Port 1433`

### Import ผ่าน แต่ SAP B1 เปิด layout ไม่ได้
- **Crystal Runtime version ต่างกัน** — .rpt สร้างด้วย CR Designer ใหม่กว่า runtime ของ Client
- **วิธีแก้:** Upgrade Crystal Reports for SAP B1 ที่เครื่อง Client หรือ re-save .rpt ด้วย CR Designer version เก่ากว่า

### `running scripts is disabled on this system`
- ใช้ `.bat` wrapper ที่มีอยู่ (ใช้ `-ExecutionPolicy Bypass`)
- หรือ: `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` (รันเป็น admin)

### Import แล้วมี layout ซ้ำ 2 ตัว
- เกิดเมื่อ `DocName` ต่างกันแต่เป็นไฟล์เดียวกัน
- ตรวจ: ก่อนหน้าเคย import ด้วย Author อื่น/ชื่ออื่นหรือไม่
- แก้: ใช้ `RunRollback.bat` เลือก row เก่า ลบทิ้ง แล้ว import ใหม่

### SKIP "unmapped ObjectType=162"
- Inventory Revaluation ไม่มี RTYP.CODE (ของ SAP จัดการแบบ custom)
- ต้อง import manual ผ่าน SAP B1 Client UI

---

## Technical Details

### การเก็บ Crystal Layout ใน SAP B1

Layout ถูกเก็บใน table **`RDOC`** มี column สำคัญ:

| Column | Type | ความหมาย |
|--------|------|---------|
| `DocCode` | nvarchar(8) | PK, format = `<TypeCode><4-digit>` เช่น `JDT20003` |
| `DocName` | nvarchar(120) | ชื่อ layout ที่แสดงใน UI |
| `Author` | nvarchar(155) | คนสร้าง (เช่น `manager`) |
| `TypeCode` | nvarchar(4) | ผูกกับ RTYP.CODE → ผูกกับ form |
| `Category` | char(1) | `C` = Crystal, `P` = PLD |
| `Status` | char(1) | `A` = Active |
| `Template` | image | ⭐ raw .rpt binary bytes |
| `RptHash` | nvarchar(254) | MD5 hex ของ Template |
| `CreateDate` | datetime | |
| `UpdateDate` | datetime | |

### DocCode Generation

Script จะ query max sequence ปัจจุบันต่อ TypeCode แล้วบวก 1

```sql
SELECT TypeCode, MAX(CAST(SUBSTRING(DocCode, LEN(TypeCode)+1, 4) AS INT))
FROM RDOC
GROUP BY TypeCode
```

เช่น ถ้ามี `QUT20008` อยู่แล้ว → ตัวถัดไปคือ `QUT20009`

### Duplicate Detection

ตรวจด้วย 2 ฟิลด์: `DocName + TypeCode` (ไม่สนใจ `Author` — re-import จะ UPDATE ข้ามเจ้าของได้)

ถ้าตรงทั้ง 2 → ถือเป็น duplicate → ทำตาม `-OnDuplicate`:
- `Update` → UPDATE `Template`, `RptHash`, `UpdateDate` (DocCode เดิม)
- `Skip` → ไม่ทำอะไร
- `Insert` → สร้าง row ใหม่ (DocCode ใหม่ → ในระบบจะมี 2 ตัว)

### Crystal .rpt File Format

- Magic bytes: `D0 CF 11 E0 A1 B1 1A E1` (OLE Compound Document)
- `RDOC.Template` = raw .rpt bytes โดยตรง (ไม่ encode เพิ่ม)
- Hash = MD5 hex (32 chars) ของ bytes เดียวกัน

### DB Plugin Architecture

`DB-MSSQL.ps1` provides:
```powershell
$DB_PARAM = "@"              # parameter prefix
$DB_NOW   = "GETDATE()"      # current timestamp
$DB_ISNUM = "ISNUMERIC"      # numeric check

function New-DBConnection { ... }  # returns SqlConnection
function Add-BlobParam     { ... }  # bind as SqlDbType.Image
function Add-DBParam       { ... }  # bind regular value
function Convert-DBSql     { ... }  # SQL dialect translation (no-op for MSSQL)
```

Main script dot-sources plugin:
```powershell
. "$PSScriptRoot\DB-MSSQL.ps1"
```

หากต้องการรองรับ HANA ในอนาคต สร้าง `DB-HANA.ps1` ตาม contract นี้และแก้ load statement

---

## ข้อควรระวัง

### ⚠️ SAP ไม่ Support การ INSERT ตรงเข้า RDOC
- วิธีนี้เป็น **undocumented** / unsupported by SAP
- ถ้าเกิดปัญหา SAP support อาจไม่ช่วย
- ⭐ **Backup DB ทุกครั้งก่อน import**

### ⚠️ Crystal Version Mismatch
- .rpt ที่สร้างด้วย Crystal Designer version **ใหม่** กว่า Crystal Runtime ของ SAP B1 Client → เปิดไม่ได้
- ทดสอบกับ 1 layout ก่อนเสมอ

### ⚠️ Password ใน `_settings.bat`
- `_settings.bat` เก็บ password เป็น **plain text** (ตามข้อจำกัดของ batch script)
- `.gitignore` exclude `_settings.bat` แล้ว — ไม่ขึ้น git, ไม่หลุดไปเครื่องอื่น
- Share repo ได้ — คนใหม่จะเห็นแต่ `_settings.bat.example` (ค่า placeholder) ต้องตั้งของตัวเอง
- **อย่า** เก็บไฟล์ `_settings.bat` ไว้ใน cloud public หรือส่งทาง email

### ⚠️ Layouts ที่ skip (import ไม่ได้ผ่าน script)
- **Inventory Revaluation** (ObjectType=162) — ไม่มี RTYP.CODE
- **Fixed Asset reports** (4 ตัว) — ไม่มี ObjectType ใน Excel
- → ต้อง import manual ผ่าน SAP UI (Tools → Crystal Reports → Import)

### ⚠️ Overwrite Behavior
- `OnDuplicate=Update` → overwrite content **ทันที** ไม่มี undo
- ของเดิมหายถาวร (เว้นมี backup)

---

## Parameter Reference

### `Import_SQL_Direct.ps1`

| Parameter | Default | ความหมาย |
|-----------|---------|---------|
| `-Server` | `SLD-C072` | SQL Server host |
| `-CompanyDB` | `SBO_SDA` | Company DB name |
| `-DBUser` | `sa` | SQL login |
| `-DBPassword` | `1q2w3e4r` | SQL password |
| `-MapFile` | Excel path | mapping file |
| `-RptRoot` | `C:\SDA\SDA\Form-Layout` | base path ของ .rpt |
| `-LogFile` | `Import_SQL_Log.txt` | log output |
| `-Author` | `SDA` | owner ที่ใส่ใน RDOC.Author |
| `-OnDuplicate` | `Update` | Update / Skip / Insert |
| `-FilterFileName` | `""` | keyword filter |
| `-UseFileNameAsDocName` | off | ใช้ filename แทน LayoutName |
| `-DryRun` | off | preview only |

---

## ตัวอย่างการใช้งานแบบ command line

```powershell
# Full import, overwrite duplicates
.\Import_SQL_Direct.ps1 -Server 10.10.10.115 -CompanyDB SBO_PROD -DBPassword "xxx"

# DryRun preview
.\Import_SQL_Direct.ps1 -DryRun

# Import เฉพาะ Journal Entry, ใช้ filename เป็น DocName
.\Import_SQL_Direct.ps1 -FilterFileName "Journal Entry" -UseFileNameAsDocName

# Skip ถ้ามีแล้ว (ไม่ overwrite)
.\Import_SQL_Direct.ps1 -OnDuplicate Skip

# เปลี่ยน Author
.\Import_SQL_Direct.ps1 -Author "manager"

# Backup DB เฉพาะตอนนี้
.\Backup-RDOC.ps1 -Server 10.10.10.115 -CompanyDB SBO_PROD -DBPassword "xxx"

# Rollback แบบเลือก (list + pick)
.\Rollback-BySelection.ps1 -Server 10.10.10.115 -CompanyDB SBO_PROD -DBPassword "xxx" -DryRun
.\Rollback-BySelection.ps1 -Server 10.10.10.115 -CompanyDB SBO_PROD -DBPassword "xxx" -Filter "Sale Order"
```

---

## License / Support

Internal tool — SDA Consult Team
Contact: consult@sala-daeng.com
