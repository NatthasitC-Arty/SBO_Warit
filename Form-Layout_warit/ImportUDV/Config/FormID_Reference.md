# SAP B1 FormID Reference

**Authoritative mapping** จากระบบจริง — source: `sap_b1_form_ids.json` (Config โฟลเดอร์เดียวกัน)

สำหรับใส่ค่าในคอลัมน์ `FormID` ของ CSV

---

## ✅ FormID ที่ verify แล้ว (ระบบจริง)

### 🟢 Sales / A/R

| FormID | Form Name |
|---|---|
| **149** | Sales Quotation |
| **139** | Sales Order |
| **140** | Delivery |
| **133** | A/R Invoice |
| **179** | A/R Credit Memo |

### 🟣 Purchasing / A/P

| FormID | Form Name |
|---|---|
| **142** | Purchase Order |
| **143** | Goods Receipt PO |
| **141** | A/P Invoice |
| **181** | A/P Credit Memo |

### 🔵 Master Data

| FormID | Form Name |
|---|---|
| **134** | Business Partner Master Data |
| **150** | Item Master Data |

### 🟡 Finance & Inventory

| FormID | Form Name |
|---|---|
| **392** | Journal Entry |
| **170** | Incoming Payments |
| **426** | Outgoing Payments |
| **940** | Inventory Transfer |
| **721** | Goods Receipt |
| **720** | Goods Issue |

---

## ⚠️ ระวัง — FormID ที่ตัวเลขใกล้กันแต่คนละ form

| FormID | Form | หมายเหตุ |
|---|---|---|
| 139 | Sales Order | |
| 140 | **Delivery** | ไม่ใช่ PO! |
| 141 | **A/P Invoice** | ไม่ใช่ Delivery! |
| 142 | **Purchase Order** | |
| 143 | **Goods Receipt PO** | ไม่ใช่ Sales Quotation! |
| 149 | **Sales Quotation** | (ที่เรียก "SQ") |
| 179 | A/R Credit Memo | |
| 181 | A/P Credit Memo | |

→ ตัวเลข 139-149 กระจายทั้ง Sales/Purchase — **ต้องดูตารางเสมอ ห้ามเดา**

---

## 🔎 วิธีหา FormID เอง (form ที่ไม่อยู่ในตาราง)

1. เปิด form ใน SAP B1 client
2. กด `Ctrl + Shift + I` (View → System Information)
3. ดู status bar — มองหา `Form=NNN`

```
Form=139   Item=38   Pane=1   Column=11
```
→ FormID = `139`

---

## 🧭 Common Item IDs (Matrix)

| ItemID | บริบท |
|---|---|
| **38** | Matrix row table — Sales/Purchase docs (SO, Invoice, PO, GRPO, etc.) |
| **76** | Matrix row table — Journal Entry / Voucher |
| **13** | Matrix row table — Inventory docs (Goods Issue/Receipt/Transfer) |

---

## 🧰 Common Matrix ColumnIDs (Sales/Purchase docs, ItemID=38)

| ColID | Field |
|---|---|
| 1 | Item No. |
| 2 | Description |
| 11 | Quantity |
| 14 | Unit Price |
| 15 | Discount % |
| 20 | (varies — Total/Tax related) |
| 21 | (varies) |
| 31 | Tax Code |

### UDF Columns (SDA setup)

| UDF Code | บริบท |
|---|---|
| `U_SLD_T_BeDis` | Total Before Discount |
| `U_SLD_Dis_Amount` | Discount Amount (per unit) |
| `U_SLD_Dis_Sum` | Discount Sum (total) |
| `U_SLD_SuppCode` | Supplier Code (JE/JV) |
| `U_SLD_LPBranch` | Legal/BP Branch |
| `U_SLD_FullName` | BP Full Name |
| `U_SLD_Title` | BP Title (บริษัท/นาย/นาง) |

---

## 📦 Forms ที่ deploy SDA discount FMS (UDV_AllSales_AllPurchase.csv)

| FormID | Form | Side |
|---|---|---|
| 149 | Sales Quotation | Sales |
| 139 | Sales Order | Sales |
| 140 | Delivery | Sales |
| 133 | A/R Invoice | Sales |
| 179 | A/R Credit Memo | Sales |
| 142 | Purchase Order | Purchase |
| 143 | Goods Receipt PO | Purchase |
| 141 | A/P Invoice | Purchase |
| 181 | A/P Credit Memo | Purchase |

รวม 9 forms — ดูรายละเอียดใน `UDV_AllSales_AllPurchase.notes.md`

---

## 🔢 FormID Patterns

| Pattern | ความหมาย |
|---|---|
| **2-3 digits** (e.g. `139`, `142`) | B1 standard form |
| **3-4 digits** (e.g. `426`, `940`) | B1 standard form (อีกกลุ่ม) |
| **10 digits** (e.g. `1470000200`) | Add-On / Customization form |

---

## 💡 Tips

1. **ใช้ตารางนี้เป็นหลัก** — มาจากระบบจริง verified แล้ว
2. **อย่าเดา FormID** — 139-149 กระจายทั้ง Sales และ Purchase
3. **Form ที่ไม่อยู่ในตาราง** → Ctrl+Shift+I เอาเอง
4. ถ้า FormID เปลี่ยน (เช่นหลัง upgrade B1) → อัปเดต `sap_b1_form_ids.json` แล้ว regenerate
