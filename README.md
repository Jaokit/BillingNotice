# BillingNotice – Excel/VBA Dormitory Billing

This project uses **Excel + VBA** to generate **dormitory rent & utilities invoices**.  
All calculations run automatically on the **Bill** sheet; with one click you can print and log invoices.

## Update 1.2

- ✅ **Meter inputs with labels (by panel):**
  - Use **Column F** for your **label text** (e.g., “Water (meter)”, “Electric (meter)”).
  - Enter **Prev** reading in **Column G** and **Current** in **Column H**.
  - The workbook computes **Usage = Current − Prev** and writes it to the **Units** cells used by billing:
    - **Water Units → C5, C16, C27**
    - **Electric Units → C6, C17, C28**
- ✅ **Live calculation & validation:**
  - Editing any G/H cell triggers the calc instantly.
  - If **Current < Prev**, the usage is cleared and G/H are **soft-red highlighted**.
- ✅ **Non-printing helper area:**
  - The bill’s print area remains **A–E**. Columns **F–H** are for labels and meter entry only (not printed).
- ✅ **Clear & Print flow updated:**
  - `SaveAllPanelsToHistorAndPrint` now **updates usage from G/H** before recalculating amounts, logging to **Histor**, printing, and clearing.
  - `ClearAllPanels` also **clears G/H** and removes any highlight color.

## Update 1.1

- ✅ **Automatic Month/Year display** — derived from the **Billing Date** you enter on the Bill sheet.
- ✅ **Automatic Room Owner name** — pulled from the **`Name`** sheet:
  - Column **A**: Room No.
  - Column **B**: Owner Name  
    Type the room number on the Bill sheet and the owner name will fill in automatically.

---

## How This Workbook Works

The workbook creates invoices for rent, water, and electricity directly inside Excel.  
Everything is auto-calculated on the **Bill** sheet; a single button prints the page and appends each filled panel to history.

### Initial Setup

1. Create a sheet named **`Name`**.
2. Enter your mapping table:
   - **Column A** → Room number (e.g., A101, B205, 301)
   - **Column B** → Owner name (e.g., Somchai K., Room 301 Owner)
3. Make sure the **Room No.** you type on the Bill sheet matches **Column A** in `Name` exactly (same text/spacing).

### Bill Sheet – What to Fill (per panel: Top / Middle / Bottom)

- **Room No.** (E2 / E13 / E24) → _Owner auto-fills to B3 / B14 / B25_
- **Billing Date** (E3 / E14 / E25) → _Month/Year auto-formats to mm/yyyy_
- **Meter Readings (with labels):**
  - Put any **label text** in **F5/F6, F16/F17, F27/F28** (optional helper text).
  - Enter **Prev** in **G5/G6, G16/G17, G27/G28** and **Current** in **H5/H6, H16/H17, H27/H28**.
  - The workbook computes **Usage = Current − Prev** and writes it to the **Units** cells:
    - Water Units → **C5, C16, C27**
    - Electric Units → **C6, C17, C28**
      > Tip: You can still type units directly in C5/C6/etc., but using G/H is recommended for accuracy.
- **Room fee** (C8 / C19 / C30):
  - Auto-priced by rules; if the room is in the special “A1–A12” range, the workbook will prompt for manual entry.
- **Fine** (C9 / C20 / C31) — optional

### Auto-calculation (Instant)

As you type (Room, Date, G/H readings, etc.), the workbook:

- Calculates **Water** = units × **28**
- Calculates **Electricity** = units × **10**
- Sets **Garbage** = **฿20** (fixed)
- Determines **Room fee** (auto or manual as required)
- Computes **Grand total** = sum of all amounts
- Applies **Thai baht format (฿#,##0)** to money fields
- Validates **Current ≥ Prev** for meter entries; otherwise highlights G/H and skips usage

### Print & Save (one click)

Click the button bound to **`SaveAllPanelsToHistorAndPrint`**. It will:

1. **Update usage** from G/H, then **recalculate** all amounts,
2. **Append** one row per filled panel to the **Histor** sheet (starting at row 2),
3. **Print** one page (3 bills),
4. **Clear** the inputs (including G/H and highlight) for next use.

### Histor Sheet (what gets saved)

For each panel with a Room No., the macro saves:  
**Month/Year, Room, Water Units, Water Amount, Electric Units, Electric Amount, Garbage, Room Fee, Fine, Grand Total**.

> Note: The **owner name** is used for display on the invoice (auto-filled from `Name`). Keep the `Name` sheet up to date.

---

## Field Map (Quick Reference)

**Panel 1 (Top):**

- **Water Units** C5 ← from **G5/H5** (Prev/Current)
- **Electric Units** C6 ← from **G6/H6**
- **Amounts**: Water E5, Electric E6, Garbage E7, Room Fee E8, Fine E9, **Total E11**
- **Room** E2, **Owner** B3, **Billing Date** E3 (mm/yyyy)

**Panel 2 (Middle):**

- **Water Units** C16 ← from **G16/H16**
- **Electric Units** C17 ← from **G17/H17**
- **Amounts**: Water E16, Electric E17, Garbage E18, Room Fee E19, Fine E20, **Total E22**
- **Room** E13, **Owner** B14, **Billing Date** E14

**Panel 3 (Bottom):**

- **Water Units** C27 ← from **G27/H27**
- **Electric Units** C28 ← from **G28/H28**
- **Amounts**: Water E27, Electric E28, Garbage E29, Room Fee E30, Fine E31, **Total E33**
- **Room** E24, **Owner** B25, **Billing Date** E25

> **Print area**: columns **A–E**. Columns **F–H** are helper/notes & meter entry (not printed).

---

## Configuration (Rates & Fees)

- Edit the constants in `Module1.bas` if needed:
  - `WATER_RATE = 28`
  - `ELEC_RATE  = 10`
  - `GARBAGE_FEE = 20`
- Room fee rules are in `RoomRate(...)`. The code prompts for manual entry for special ranges (e.g., **A1–A12**).

---

## Troubleshooting & Notes

- If auto-calculation or events get disabled (e.g., after an interrupted macro), run **`ResetAppState`**.
- If the owner name does not appear:
  - Ensure **Room No.** on **Bill** matches exactly a value in **`Name`!A** (no extra spaces, same format).
  - Confirm the `Name` sheet is spelled exactly **Name**.
- If meter **Current < Prev**, usage is cleared and G/H are highlighted. Fix the readings and the usage will recalc automatically.
