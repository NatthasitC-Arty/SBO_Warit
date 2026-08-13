# SAP B1 FormID Reference

Quick lookup ของ FormID ที่ใช้บ่อย + form ในระบบของ user
สำหรับใส่ค่าในคอลัมน์ `FormID` ของ CSV

---

## 🔎 วิธีหา FormID เอง (เมื่อไม่อยู่ในตารางนี้)

1. เปิด form ที่ต้องการใน SAP B1 client
2. กด `Ctrl + Shift + I` (View → System Information)
3. ดู status bar ที่ด้านล่างหน้าจอ
4. มองหา `Form=NNN`

ตัวอย่าง status bar:
```
Form=139   Item=38   Pane=1   Column=11
```

→ FormID = `139`

---

## ⭐ ในระบบของคุณ (จาก Export ที่ผ่านมา)

| FormID | ชื่อที่คุณใช้ | ประเภท | จำนวน FMS |
|---|---|---|---|
| **139** | Sales Order (SO) | Sales | 2 → 6 (หลังจาก mirror จาก 149) |
| **149** | Sales Quotation (SQ) — ในระบบคุณ | Sales | 6 ⭐ template หลัก |
| **134** | BP-related (UDF Series/Account/FullName) | BP | 3 |
| **150** | Business Partner Master | BP | 1 |
| **185** | Series / Setup-related | Admin | 2 |
| **392** | Journal Entry (JE) | Finance | 6 |
| **393** | Journal Voucher (JV) | Finance | 6 |
| **65300** | Custom form | Custom | 1 |
| **720** | Goods Issue | Inventory | 1 |
| **721** | Goods Receipt | Inventory | 1 |
| **1470000200** | Add-On / Custom form | Custom | 3 |

---

## 📋 SAP B1 v10 Standard FormIDs

### Sales / Marketing
| FormID | Form Name |
|---|---|
| 143 | Sales Quotation (standard B1) |
| **139** | **Sales Order** |
| 140 | (varies) |
| 141 | Delivery |
| 142 | Returns |
| 133 | A/R Invoice |
| 60090 | A/R Invoice (newer template) |
| 149 | A/R Credit Memo (in some versions) |
| 203 | A/R Down Payment |

### Purchasing
| FormID | Form Name |
|---|---|
| 540000405 | Purchase Quotation |
| 22 / 1470000113 | Purchase Order / Request |
| 140 | Purchase Order (v10) |
| 146 | Goods Receipt PO |
| 21 | Goods Return |
| 142 | A/P Invoice |
| 159 | A/P Credit Memo |
| 204 | A/P Down Payment |

### Inventory
| FormID | Form Name |
|---|---|
| **720** | **Goods Issue** |
| **721** | **Goods Receipt** |
| 940 | Inventory Transfer |
| 941 | Inventory Counting |
| 162 | Inventory Revaluation |

### Banking
| FormID | Form Name |
|---|---|
| 170 | Incoming Payment |
| 180 | Outgoing Payment |
| 69 | Landed Costs |

### Production
| FormID | Form Name |
|---|---|
| 202 | Production Order |
| 60 | Goods Issue (Production) |
| 59 | Goods Receipt (Production) |

### Master Data
| FormID | Form Name |
|---|---|
| **150** | **Business Partner Master Data** |
| 156 | Item Master Data |
| 157 | Service Call |

### Financials
| FormID | Form Name |
|---|---|
| **392** | **Journal Entry** |
| **393** | **Journal Voucher** |
| 720 | (overlap — depends on version) |

### Setup / Admin
| FormID | Form Name |
|---|---|
| **185** | (varies — likely Document Numbering / Series) |
| 100 | General Settings |
| 5000 | Choose Company |
| 130 | Document Settings |

---

## 🔢 FormID Patterns

| Pattern | ความหมาย |
|---|---|
| **2-3 digits** (e.g. `22`, `139`) | Legacy B1 standard form (มาตั้งแต่ v9.x) |
| **4-5 digits** (e.g. `60090`, `65300`) | v10 standard form (template ใหม่) |
| **10 digits starting 14**` (e.g. `1470000200`) | Add-On / Customization form |
| **10 digits starting other** (e.g. `540000405`) | Specific version's reissue |

---

## 🧭 Common Item IDs (Matrix)

หลังจากรู้ FormID แล้ว ItemID ที่ใช้บ่อย:

| ItemID | บริบท |
|---|---|
| **38** | Matrix row table ของ Sales/Purchase docs (Sales Order, Invoice, etc.) |
| **76** | Matrix row table ของ Journal Entry / Voucher |
| **13** | Matrix row table ของ Inventory docs (Goods Issue/Receipt) |
| **4** | Header field (varies by form) |
| **7, 16, 17, 38** | Header UDF fields (varies) |

---

## 🧰 Common Matrix ColumnIDs (Sales/Purchase docs, ItemID=38)

| ColID | Field |
|---|---|
| 1 | Item No. |
| 2 | Description |
| 11 | Quantity |
| 14 | Unit Price |
| 15 | Discount % |
| 20 | (varies — Total/Tax) |
| 21 | (varies — Total After Discount/Tax) |
| 31 | Tax Code |

### Common UDF Columns (SDA setup specific)
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

## 💡 Tips

1. **ถ้าไม่แน่ใจ FormID** — เปิด form นั้นจริงๆ แล้วใช้ `Ctrl+Shift+I` เสมอ (เร็วและถูกต้องที่สุด)
2. **FormID แตกต่างกันได้** ตามเวอร์ชั่นและ localization (Thai vs English)
3. **FormID 10 หลัก** = custom/add-on form ระวัง — อาจไม่มีในระบบอื่นถ้า migrate
4. **Item Master Data** อาจเป็น 156 หรือ 157 ขึ้นกับ version
5. **Sales Quotation** มาตรฐาน B1 = 143, แต่ระบบของคุณใช้ 149 (custom mapping)

---

## 🔗 Cross-Reference กับ CSV

ใน `UDV_*.csv` column `FormID`:
- ใส่เลขตามตารางนี้
- หรือ Ctrl+Shift+I บน form จริงเอามาใส่
- เปิด `UDV_SO_Match_SQ.notes.md` เพื่อดูตัวอย่าง mapping
