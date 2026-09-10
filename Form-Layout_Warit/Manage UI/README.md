# Manage UI — SAP B1 Form Settings (CPRF) Toolkit

Tools to **export, diff, and mirror** SAP B1 matrix Form Settings
(column visibility, order, width per `FormID/ItemID/ColID`) — the
data behind **Tools → Form Settings** dialog.

Bypasses DI API (which has known gaps for UI templates) by writing
directly to `CPRF` table via .NET SqlClient in a single transaction.

---

## Schema cheat sheet

### Tables involved

| Table | Purpose |
|-------|---------|
| **CPRF** | Column profile per `(FormID, ItemID, ColID, UserSign, TPLId)` — the row-level state |
| **UICU** | UI Template master: `TPLId, TPLName, TPLDesc, UserID, Parent, IsTemplate` |
| **UIC3** | Template ↔ User assignment (which users receive which template) |
| **UIC4** | (related — per-user per-form override flag) |
| **OUSR** | Users: `USERID` ↔ `USER_CODE` |

### CPRF columns (17)

`FormID, ItemID, ColID, Width, VisInForm, VisualIndx, EditInForm,
VisInExpnd, ExpandIndx, EditInEXP, Folded, UserSign, ExtDisable,
ExtInvsbl, TPLId, TableName, ItemUID`

| Column | Meaning |
|--------|---------|
| `VisInForm` | `Y`/`N` — column visible in normal view |
| `VisualIndx` | Column position (lower = leftmost) |
| `Width` | Pixel width (0 if hidden) |
| `EditInForm` | `Y`/`N` — column editable |
| `VisInExpnd / ExpandIndx / EditInEXP` | Same flags for "Expanded" matrix view |
| `Folded` | Column collapsed/folded |
| `TPLId` | UI Template ID (FK → `UICU.TPLId`) |
| `UserSign` | User who owns these rows (not always = `UICU.UserID`) |

### Template pattern (observed)

Each user gets:
- **`<UserID>_local_template`** (Parent=0) — personal default
- **`<UserID>_local_template_<ParentTPLId>`** (Parent=N) — local override of a shared template

Shared/named templates (`IsTemplate=Y`) like `"UI ALL"` or `"คลังสินค้า"`
are assigned to users via **UIC3**. When a user edits Form Settings
under a shared template, B1 creates `<UserID>_local_template_<masterTPLId>`
as their personal override.

### UserSign quirk

`CPRF.UserSign` ≠ `UICU.UserID` for **shared master templates**:
- TPLId=3 ("UI ALL", owner SLD01 in UICU) → all 919 CPRF rows are
  stored under `UserSign=1 (manager)`, not 26 (SLD01).
- TPLId=66 ("คลังสินค้า", owner manager) → CPRF rows under manager.

Investigate with `FindTPLOwner.bat` before mirroring.

---

## Tools

All `.bat` files read shared connection settings from `_settings.bat`
(copy `_settings.bat.example`, edit values, gitignored).

### Discovery / read-only

| Tool | Purpose |
|------|---------|
| `ProbeFormSettings.bat` | Find which table holds matrix settings (run once) |
| `ListUsers.bat` | Dump OUSR + per-user CPRF row counts for `$FORMID` |
| `ProbeTemplates.bat` | TPLId breakdown for a user/form |
| `ProbeUICU.bat` | List all UI Templates + UIC3 assignments |
| `FindTPLOwner.bat` | Map TPLId → CPRF.UserSign actual owner |
| `ExportCPRF.bat` | Dump CPRF to `Config\CPRF_Export_<ts>.csv` (filter by FORMID/ITEMID/COLID) |
| `DiffTemplates.bat` | Compare two `(UserSign, TPLId)` profiles |

### Mutation (transaction-wrapped)

| Tool | Purpose |
|------|---------|
| `MirrorFormUI.bat` | Copy CPRF rows from `(srcUser, srcTPL)` → `(tgtUser, tgtTPL)` for one FormID. Atomic DELETE+INSERT. `MODE=-DryRun` for preview. |

---

## Workflow: push manager's edits to all users

### Phase 1 — Discover

```
ProbeFormSettings.bat   → confirm CPRF is the right table
ListUsers.bat           → find manager's UserSign (= 1)
ProbeUICU.bat           → list templates; pick the shared master to push to
```

### Phase 2 — Inspect current state

```
ExportCPRF.bat          → set FORMID=139, export to CSV
ProbeTemplates.bat      → see which TPLIds the editing user has
FindTPLOwner.bat        → critical: which UserSign holds the target template's CPRF rows
```

### Phase 3 — Decide strategy

| Strategy | What it does | Reach |
|----------|--------------|-------|
| **A. Push to master** | Mirror source → shared TPLId (e.g. TPLId=3 "UI ALL") | Only users **without** local overrides |
| **B. Push to all local overrides** | Mirror source → every `<X>_local_template_<masterTPLId>` | All users currently assigned the master |
| **C. A + DELETE local overrides** | Push master + drop all local copies so B1 regenerates from master | Everyone (destructive) |

### Phase 4 — Apply

```
DiffTemplates.bat       → preview what will change
MirrorFormUI.bat        → DryRun first, review log, then real run
```

### Phase 5 — Verify

Login B1 as a target user → open the form → Tools → Form Settings →
confirm columns match.

> ⚠️ Users with **existing local overrides** read from their local TPLId,
> not the master — they won't see Strategy A changes until their local
> copy is deleted or also updated.

---

## Example: push manager's SO edit to UI ALL master

History (already run on `SBO_Update_UI`, FormID=139):

1. Manager logged in to B1, opened Sales Order, customized Form Settings.
   B1 created TPLId=34 (`1_local_template_3`, Parent=3) with 828 rows.
2. UI ALL master = TPLId=3, 919 rows under `UserSign=1` (not 26 despite UICU owner=SLD01).
3. Ran:
   ```
   MirrorFormUI.bat
     FORMID=139
     SRC_USER=1  SRC_TPL=34   (manager's edit)
     TGT_USER=1  TGT_TPL=3    (UI ALL master)
   ```
4. Result: DELETE 919 + INSERT 828, committed.

**Outcome:** master template now matches manager's edit, but verification
pending — users with local overrides (TPLId 34–54, 56, 58, 60, 64) still
see their own copy, not the new master.

---

## Files in this folder

```
Manage UI/
├── README.md                      (this file)
├── _settings.bat                  (gitignored — has password)
├── _settings.bat.example          (template)
│
├── ProbeFormSettings.bat          → Scripts/Probe-FormSettings.ps1
├── ListUsers.bat                  → Scripts/List-Users.ps1
├── ProbeTemplates.bat             → Scripts/Probe-Templates.ps1
├── ProbeUICU.bat                  → Scripts/Probe-UICU.ps1
├── FindTPLOwner.bat               → Scripts/Find-TPLOwner.ps1
├── ExportCPRF.bat                 → Scripts/Export-CPRF.ps1
├── DiffTemplates.bat              → Scripts/Diff-Templates.ps1
├── MirrorFormUI.bat               → Scripts/Mirror-FormUI.ps1
│
├── Config/                        (CSV exports + future mapping files)
├── Scripts/                       (all .ps1)
├── Mirror_FormUI_Log.txt          (rolling log of Mirror runs)
└── Export_CPRF_Log.txt            (rolling log of Export runs)
```

---

## Caveats

- **No automatic rollback after commit.** Take a SQL backup of CPRF
  (or `SELECT INTO CPRF_backup_<date>`) before real `MirrorFormUI.bat` runs.
- **B1 client caches Form Settings.** Restart B1 or log out/in to see
  changes from another user's session.
- **CPRF row count varies per TPLId.** Shared masters tend to have
  more rows (919) than user-specific overrides (815–828) because
  overrides only include columns the user actually customized.
- **B1 version-specific schema.** Tooling discovers CPRF columns
  dynamically via `INFORMATION_SCHEMA` — safe across versions, but
  validate `ItemUID`/`TableName`/`ExtDisable` exist on your B1 build.
- **`Manage UI` folder name has a space.** Always quote paths in
  shell commands.
