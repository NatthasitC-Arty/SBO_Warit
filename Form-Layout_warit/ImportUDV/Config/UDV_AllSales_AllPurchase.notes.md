# UDV_AllSales_AllPurchase.csv — Coverage & Notes

Deploy SDA discount FMS pattern ครบทั้ง **Sales (7 forms)** และ **Purchase (6 forms)**

**Total: 13 forms × 10 CSV rows = 130 rows**

---

## 📊 Forms Covered

### 🟢 Sales side (7 forms)

| FormID | Name | ObjectType | TypeCode |
|---|---|---|---|
| **143** | Sales Quotation (standard) | 23 | QUT2 |
| **139** | Sales Order ⭐ existing | 17 | RDR2 |
| **141** | Delivery | 15 | DLN2 |
| **178** | Returns | 16 | RDN2 |
| **133** | A/R Invoice | 13 | INV2 |
| **179** | A/R Credit Memo | 14 | RIN2 |
| **149** | (user's "SQ" / AR Credit Memo) ⭐ existing | — | — |

### 🟣 Purchase side (6 forms)

| FormID | Name | ObjectType | TypeCode |
|---|---|---|---|
| **540000405** | Purchase Quotation | 540000405 | PQT2 |
| **22** | Purchase Order | 22 | POR2 |
| **20** | Goods Receipt PO | 20 | PDN2 |
| **21** | Goods Return | 21 | RPD2 |
| **18** | A/P Invoice | 18 | PCH2 |
| **19** | A/P Credit Memo | 19 | RPC2 |

---

## 📋 FMS Pattern (ต่อ form)

ทุก form จะได้ FMS 6 ตัวเหมือนกัน:

| ColumnID | FMSAction | Query | Refresh | Triggers (in SHS1) |
|---|---|---|---|---|
| `15` | Q | FMS_DisAmount | Y | U_SLD_T_BeDis, U_SLD_Dis_Sum |
| `20` | F | — | N | — (ByField=C) |
| `21` | F | — | N | — |
| `U_SLD_Dis_Amount` | Q | FMS_ToUnit | Y | U_SLD_Dis_Sum, 11 |
| `U_SLD_Dis_Sum` | Q | FMS_ToSum | Y | U_SLD_Dis_Amount, 11 |
| `U_SLD_T_BeDis` | Q | FMS_Price_Total | Y | 11, 14 |

ItemID = 38 (matrix row table) ทุก form

---

## ⚠️ Verify FormIDs Before Import

FormIDs ในไฟล์นี้ใช้ **มาตรฐาน SAP B1 v10** — บางเครื่องอาจ map ต่าง (โดยเฉพาะ custom/localization)

**วิธี verify:**
1. เปิด form ที่ต้องการ → กด `Ctrl + Shift + I`
2. ดู `Form=NNN` ที่ status bar
3. ถ้าไม่ตรง — แก้ในคอลัมน์ FormID ของ CSV

**FormIDs ที่อาจต้องระวัง:**
- `143` Sales Quotation — บางระบบใช้ `149` แทน (เช่นของคุณตอนนี้)
- `133, 178, 179` AR docs — บางระบบใช้ template ใหม่ `60090`, `60091`
- `22` Purchase Order — บางระบบ v10 ใหม่ใช้ `142` (newer template)
- `20, 21` Goods Receipt/Return PO — บางระบบใช้ `146`, `144`
- `18, 19` AP Invoice/Credit Memo — บางระบบใช้ `142`, `159`

---

## 🚦 Pre-Import Checklist

- [ ] Verify FormIDs match your B1 install (Ctrl+Shift+I per form)
- [ ] ItemID 38 = matrix row table on the form
- [ ] UDF columns `U_SLD_T_BeDis`, `U_SLD_Dis_Amount`, `U_SLD_Dis_Sum` มีอยู่ใน table ของ form (`RDR1`, `INV1`, `POR1`, `PDN1`, `PCH1` etc.)
- [ ] DryRun ก่อน real run

---

## 🏃 วิธีรัน

### 1. DryRun ดู preview ก่อน
แก้ `RunImportUDV.bat`:
```bat
set MODE=-DryRun
```
รัน → จะแสดง 13 unique CSHS keys × form = 78 [NEW/EXISTS] preview

### 2. Real run
ตั้ง `set MODE=` (เอา -DryRun ออก) → รัน

### 3. Verify
```sql
-- ตรวจว่ามี FMS บนทุก form หรือยัง
SELECT FormID, COUNT(*) AS Cnt
FROM CSHS
WHERE FormID IN ('143','139','141','178','133','179','149','540000405','22','20','21','18','19')
GROUP BY FormID
ORDER BY FormID;
-- ควรเห็น 6 ต่อ form
```

```sql
-- ตรวจ SHS1 triggers
SELECT c.FormID, c.ColID, COUNT(s.FieldID) AS TriggerCount
FROM CSHS c
LEFT JOIN SHS1 s ON s.IndexID = c.IndexID
WHERE c.FormID IN ('143','139','141','178','133','179','149','540000405','22','20','21','18','19')
GROUP BY c.FormID, c.ColID
ORDER BY c.FormID, c.ColID;
-- col 15, U_SLD_Dis_Amount, U_SLD_Dis_Sum, U_SLD_T_BeDis ควรมี 2 triggers
```

### 4. Restart B1 client → เช็คใน UI

---

## 🧪 ทดสอบหลัง Import

แต่ละ form เปิดดู:
1. เพิ่ม item row
2. กรอก Qty + Unit Price
3. ดูว่า `U_SLD_T_BeDis` คำนวณ qty × price อัตโนมัติ
4. กรอก Discount Amount/Sum → ดูว่าค่าอื่นๆ refresh ตาม

---

## 📐 Pattern Logic Recap

```
Quantity × Unit Price ─────────────► U_SLD_T_BeDis (Total Before Discount)
                                            │
                                            ▼
              U_SLD_Dis_Sum ◄────────► U_SLD_Dis_Amount × Qty
                                            │
                                            ▼
                                     col 15 (% = Dis_Sum / T_BeDis × 100)
```

ทุก field interconnect — เปลี่ยน 1 ตัว → refresh ตัวอื่นๆ ที่ depend

---

## 🔄 Round-trip Workflow

```
Edit CSV ─► Import ─► Verify in B1 ─► (optional) Re-export to confirm
```

ถ้าจะ migrate ไปอีก Company DB:
1. Edit `_settings.bat` ชี้ DB อื่น
2. รัน `RunImportUDV.bat` เลือกไฟล์นี้
3. Same 130 rows → 78 CSHS rows + ~150+ SHS1 trigger rows
