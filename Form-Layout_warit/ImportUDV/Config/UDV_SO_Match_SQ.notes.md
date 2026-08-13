# UDV_SO_Match_SQ.csv — Column Notes

CSV นี้ deploy 6 FMS ลง **Sales Order (FormID=139)** มี total 10 rows
(หลาย rows ต่อ FMS เพราะ multi-trigger)

---

## Schema (14 columns)

| # | Column | ค่าใน CSV นี้ | คำอธิบาย |
|---|---|---|---|
| 1 | `Action` | `UPSERT` ทุก row | คำสั่ง — `UPSERT` = สร้างถ้าไม่มี / แก้ถ้ามี (ปลอดภัยสุด) |
| 2 | `FormID` | `139` | Sales Order (จาก System Info Ctrl+Shift+I) |
| 3 | `ItemID` | `38` | Matrix item ของ Sales Order (row table) |
| 4 | `ColumnID` | `15` / `20` / `21` / `U_SLD_*` | ระบุ column ใน matrix — ตัวเลข = system col, `U_SLD_*` = UDF col |
| 5 | `FMSAction` | `Q` / `F` | `Q` = Saved Query, `F` = Fixed Value |
| 6 | `QueryName` | `FMS_*` (เมื่อ Q) | ชื่อ query ใน `OUQR.QName` — ถ้ามีชื่อนี้แล้วจะ reuse |
| 7 | `QueryCategory` | `-1` | `-1` = General category (ใช้มาตรฐาน) |
| 8 | `QueryBody` | SQL string (เมื่อ Q) | SQL จริงของ query มี token `$[$ItemID.Col.Type]` |
| 9 | `FixedValue` | ว่าง | ใช้เฉพาะเมื่อ FMSAction = F (ค่าคงที่) |
| 10 | `Refresh` | `Y` / `N` | `Y` = Auto-Refresh on field change |
| 11 | `ByField` | `N` / `C` | DB flag — `N` ปกติ, `C` กรณีพิเศษ (col 20) |
| 12 | `TriggerID` | UDF code หรือ col number | ⭐ Trigger field — เขียนลง **SHS1** ไม่ใช่ CSHS |
| 13 | `TriggerColumn` | ว่าง | Legacy column — ไม่ใช้ใน B1 v10 |
| 14 | `ForceRefresh` | `N` | `Y` = "Display Saved Values" ทันทีเมื่อเปิด form |

---

## Row-by-Row breakdown

### `139/38/15` — Discount % (col 15) — มี 2 triggers

```csv
UPSERT,139,38,15,Q,FMS_DisAmount,-1,"SELECT CASE ... END",,Y,N,U_SLD_T_BeDis,,N
UPSERT,139,38,15,Q,FMS_DisAmount,-1,"SELECT CASE ... END",,Y,N,U_SLD_Dis_Sum,,N
```

**ทำอะไร:** คำนวณ Discount % จาก `Dis_Sum / T_BeDis × 100`

**Multi-row:** 2 rows = 2 triggers ที่จะใส่ใน SHS1:
- Trigger 1: `U_SLD_T_BeDis` (Total Before Discount)
- Trigger 2: `U_SLD_Dis_Sum` (Discount Sum)

→ เมื่อค่าใน 2 field นี้เปลี่ยน → FMS auto-refresh

---

### `139/38/20` และ `139/38/21` — Fixed-Value placeholders

```csv
UPSERT,139,38,20,F,,,,,N,C,,,N
UPSERT,139,38,21,F,,,,,N,N,,,N
```

**ทำอะไร:** Fixed-Value FMS (F type) ที่ไม่มี QueryBody หรือ FixedValue
- น่าจะเป็น placeholder/marker row ที่ B1 ใช้สื่อสารกับ matrix state
- col 20 มี `ByField=C` (พิเศษ) — น่าจะเป็น checkbox state

**Note:** Row เดียวต่อ FMS เพราะ F-type ไม่มี trigger

---

### `139/38/U_SLD_Dis_Amount` — Discount Per Unit — 2 triggers

```csv
UPSERT,139,38,U_SLD_Dis_Amount,Q,FMS_ToUnit,-1,"SELECT CASE ... END",,Y,N,U_SLD_Dis_Sum,,N
UPSERT,139,38,U_SLD_Dis_Amount,Q,FMS_ToUnit,-1,"SELECT CASE ... END",,Y,N,11,,N
```

**ทำอะไร:** คำนวณ Discount Per Unit จาก `Dis_Sum / Qty`

**Triggers:**
- `U_SLD_Dis_Sum` (UDF column)
- `11` (= Quantity, system column)

---

### `139/38/U_SLD_Dis_Sum` — Discount Sum — 2 triggers

```csv
UPSERT,139,38,U_SLD_Dis_Sum,Q,FMS_ToSum,-1,"SELECT ... * ...",,Y,N,U_SLD_Dis_Amount,,N
UPSERT,139,38,U_SLD_Dis_Sum,Q,FMS_ToSum,-1,"SELECT ... * ...",,Y,N,11,,N
```

**ทำอะไร:** คำนวณ Discount Sum = `Dis_Amount × Qty`

**Triggers:** `U_SLD_Dis_Amount` + `11` (Quantity)

---

### `139/38/U_SLD_T_BeDis` — Total Before Discount — 2 triggers

```csv
UPSERT,139,38,U_SLD_T_BeDis,Q,FMS_Price_Total,-1,"SELECT CAST(...) * (CAST(...))",,Y,N,11,,N
UPSERT,139,38,U_SLD_T_BeDis,Q,FMS_Price_Total,-1,"SELECT CAST(...) * (CAST(...))",,Y,N,14,,N
```

**ทำอะไร:** คำนวณ Total Before Discount = `Qty × Unit Price` (col 11 × col 14)

**Triggers:** `11` (Quantity) + `14` (Unit Price)

---

## ⚠️ จุดสำคัญที่ต้องเข้าใจ

### 1. Multi-Row = Multi-Trigger

Rows ที่มี `(FormID, ItemID, ColumnID)` เดียวกัน → 1 FMS หลาย triggers

Import script จะ:
1. Group rows ตาม key
2. ใช้ row แรกของ group กำหนด CSHS + OUQR
3. **DELETE** ทุก SHS1 row ของ IndexID นั้น
4. **INSERT** SHS1 row ละ TriggerID จากทุก row ใน group

### 2. ทุก row ใน group ต้องมีข้อมูล identical (ยกเว้น TriggerID)

QueryBody, QueryName, Refresh, ByField, ForceRefresh ฯลฯ **ต้องเหมือนกันทุก row** ของ group เดียวกัน — script ใช้ row แรกเป็นต้นแบบ

### 3. CSHS.FieldID จะถูกตั้งเป็นว่าง

Script ปล่อยให้ SHS1 จัดการ trigger list ทั้งหมด — CSHS.FieldID = empty เสมอ

### 4. Shared QueryName = แก้ที่เดียวกระทบทุกฟอร์ม

`FMS_Price_Total` ในไฟล์นี้ใช้ที่ 139/U_SLD_T_BeDis และอาจ shared กับ 149, 65300 (form อื่นที่ใช้ query name เดียวกัน)

→ แก้ `QueryBody` ใน CSV นี้ จะ UPDATE `OUQR.QString` → **กระทบทุกฟอร์มที่ใช้ query นี้**

ถ้าไม่อยากกระทบ → เปลี่ยน `QueryName` ใน CSV เป็นชื่อใหม่ เช่น `FMS_Price_Total_SO`

### 5. Token Syntax ใน QueryBody

| Pattern | ตัวอย่างในไฟล์ | ความหมาย |
|---|---|---|
| `$[$38.11.NUMBER]` | col U_SLD_T_BeDis | Matrix col 11 (Quantity) as number |
| `$[$38.14.NUMBER]` | col U_SLD_T_BeDis | Matrix col 14 (Unit Price) as number |
| `$[$38.U_SLD_T_BeDis.NUMBER]` | col 15 | Matrix UDF col |
| `$[$38.U_SLD_Dis_Sum.NUMBER]` | หลายที่ | Matrix UDF col |

---

## 🔄 Database write summary หลัง Import

| Table | Rows ที่จะถูกแก้ |
|---|---|
| `CSHS` | 6 rows (1 ต่อ unique key) — UPSERT FormID/ItemID/ColID/ActionT/QueryId/Refresh/ByField/FrceRfrsh |
| `SHS1` | 10 rows (DELETE old + INSERT 1 ต่อ trigger) |
| `OUQR` | 4 rows (FMS_DisAmount, FMS_ToUnit, FMS_ToSum, FMS_Price_Total) — UPDATE QString ถ้ามีอยู่ |
