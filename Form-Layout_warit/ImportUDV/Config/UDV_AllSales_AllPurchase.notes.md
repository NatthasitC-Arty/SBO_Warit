# UDV_AllSales_AllPurchase.csv — Coverage & Notes

Deploy SDA discount FMS pattern ครบทั้ง **Sales (5 forms)** และ **Purchase (4 forms)**

**FormIDs verified จาก `sap_b1_form_ids.json` (ระบบจริง)**

**Total: 9 forms × 11 CSV rows = 99 rows**

---

## 📊 Forms Covered (verified FormIDs)

### 🟢 Sales side (5 forms)

| FormID | Name |
|---|---|
| **149** | Sales Quotation |
| **139** | Sales Order |
| **140** | Delivery |
| **133** | A/R Invoice |
| **179** | A/R Credit Memo |

### 🟣 Purchase side (4 forms)

| FormID | Name |
|---|---|
| **142** | Purchase Order |
| **143** | Goods Receipt PO |
| **141** | A/P Invoice |
| **181** | A/P Credit Memo |

---

## 📋 FMS Pattern (ต่อ form — 11 CSV rows / 6 CSHS keys)

อ้างอิงจาก export จริงของ FormID 139 (Sales Order)

| ColumnID | FMSAction | Query | Refresh | Triggers (SHS1) |
|---|---|---|---|---|
| `15` | Q | FMS_DisAmount | Y | U_SLD_Dis_Amount, U_SLD_Dis_Sum, U_SLD_T_BeDis (3) |
| `20` | F | — | N | — (ByField=C) |
| `21` | F | — | N | — (ByField=N) |
| `U_SLD_Dis_Amount` | Q | **FMS_ToSum** | Y | U_SLD_Dis_Sum (1) |
| `U_SLD_Dis_Sum` | Q | **FMS_ToUnit** | Y | 11, U_SLD_Dis_Amount (2) |
| `U_SLD_T_BeDis` | Q | FMS_Price_Total | Y | 11, 14, 20 (3) |

ItemID = 38 (matrix row table) ทุก form

**Query bodies** (เก็บใน OUQR, shared ทุก form):
- `FMS_DisAmount` — Discount % = `Dis_Sum / col14 × 100`
- `FMS_ToSum` — `Dis_Sum × Qty` (on col U_SLD_Dis_Amount)
- `FMS_ToUnit` — `Dis_Amount / Qty` (on col U_SLD_Dis_Sum)
- `FMS_Price_Total` — `Qty × col14` (on col U_SLD_T_BeDis)

⚠️ **Note:** ชื่อ query กับ column ดู "สลับกัน" (col Dis_Amount ใช้ FMS_ToSum, col Dis_Sum ใช้ FMS_ToUnit) — นี่คือ pattern จริงจากระบบ ไม่ใช่ความผิดพลาด

---

## 🚦 Pre-Import Checklist

- [x] FormIDs verified จาก `sap_b1_form_ids.json`
- [ ] ItemID 38 = matrix row table บนแต่ละ form (ส่วนใหญ่ใช่ — verify Ctrl+Shift+I ถ้าไม่แน่ใจ)
- [ ] UDF columns `U_SLD_T_BeDis`, `U_SLD_Dis_Amount`, `U_SLD_Dis_Sum` มีอยู่ใน table ของแต่ละ form:
      - Sales: `RDR1` (SO), `QUT1` (Quotation), `DLN1` (Delivery), `INV1` (AR Invoice), `RIN1` (AR Credit Memo)
      - Purchase: `POR1` (PO), `PDN1` (GRPO), `PCH1` (AP Invoice), `RPC1` (AP Credit Memo)
- [ ] DryRun ก่อน real run

⚠️ **ถ้า UDF ยังไม่มีในบาง table** — ต้องสร้าง UDF ก่อน (ผ่าน Tools > Customization Tools > User-Defined Fields) ไม่งั้น FMS จะ attach ไม่ได้

---

## 🏃 วิธีรัน

### 1. DryRun
แก้ `RunImportUDV.bat`: `set MODE=-DryRun`
รัน → preview 54 unique CSHS keys (9 forms × 6 keys)

### 2. Real run
`set MODE=` (เอา -DryRun ออก) → รัน

### 3. Verify
```sql
-- จำนวน FMS ต่อ form (ควรเห็น 6)
SELECT FormID, COUNT(*) AS Cnt
FROM CSHS
WHERE FormID IN ('149','139','140','133','179','142','143','141','181')
GROUP BY FormID
ORDER BY FormID;
```

```sql
-- Triggers ต่อ column
SELECT c.FormID, c.ColID, COUNT(s.FieldID) AS TriggerCount
FROM CSHS c
LEFT JOIN SHS1 s ON s.IndexID = c.IndexID
WHERE c.FormID IN ('149','139','140','133','179','142','143','141','181')
GROUP BY c.FormID, c.ColID
ORDER BY c.FormID, c.ColID;
-- col 15, U_SLD_Dis_Amount, U_SLD_Dis_Sum = 2 triggers
-- col U_SLD_T_BeDis = 3 triggers (11, 14, 20)
```

### 4. Restart B1 client → เช็คใน UI

---

## 📐 Pattern Logic

```
Quantity x Unit Price ──────────► U_SLD_T_BeDis (Total Before Discount)
                                         │
                                         ▼
           U_SLD_Dis_Sum ◄─────────► U_SLD_Dis_Amount x Qty
                                         │
                                         ▼
                                  col 15 (% = Dis_Sum / T_BeDis x 100)
```

---

## 🔄 Migrate ไป Company DB อื่น

1. แก้ `_settings.bat` ชี้ DB ปลายทาง
2. รัน `RunImportUDV.bat` → เลือก `UDV_AllSales_AllPurchase.csv`
3. Same 99 rows → 54 CSHS rows + ~99 SHS1 trigger rows
4. Restart B1 client บนเครื่องที่ใช้ DB นั้น

---

## 📎 Related Files

| File | หน้าที่ |
|---|---|
| `sap_b1_form_ids.json` | Authoritative FormID mapping (source of truth) |
| `FormID_Reference.md` | FormID lookup table (human-readable) |
| `UDV_AllSales_AllPurchase.csv` | ไฟล์ data นี้ |
