# BillingNotice – Excel/VBA Dormitory Billing

This project uses **Excel + VBA** to generate **dormitory rent & utilities invoices**.  
All calculations run automatically on the **Bill** sheet; with one click you can print and log invoices.

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
1) Create a sheet named **`Name`**.  
2) Enter your mapping table:  
   - **Column A** → Room number (e.g., A101, B205, 301)  
   - **Column B** → Owner name (e.g., Somchai K., Room 301 Owner)  
3) Make sure the **Room No.** you type on the Bill sheet matches **Column A** in `Name` exactly (same text/spacing).

### Workflow

1) **Open & Enable Macros**  
   Open the file and allow macros.

2) **Fill 3 Panels on the _Bill_ sheet** (Top / Middle / Bottom). For each panel, enter:  
   - **Room No.** → *Owner name auto-fills from* `Name`  
   - **Billing Date** → *Month/Year shows automatically*  
   - **Water units**  
   - **Electric units**  
   - **Room fee** (only if the room is in A1–A12; other rooms are auto-priced)  
   - **Fine** (optional)

3) **Auto-calculation (instant)**  
   As you type, the workbook calculates and displays:  
   - **Water** = units × **28**  
   - **Electricity** = units × **10**  
   - **Garbage** = **฿20** (fixed)  
   - **Room fee** = automatic per room rules (or manual for A1–A12)  
   - **Grand total** = sum of all amounts  
   - **Month/Year** = derived automatically from **Billing Date**  
   All money fields use **Thai baht format (฿#,##0)**.

4) **Print & Save (one click)**  
   Click the button bound to `SaveAllPanelsToHistorAndPrint`. It will:  
   - Recalculate all 3 panels,  
   - **Append** one row per filled panel to the **Histor** sheet (start writing at row 2; row 1 is your header),  
   - **Print** one page (3 bills),  
   - **Clear** the inputs for next use (including Room/Date).

### Histor Sheet (what gets saved)
For each panel with a Room No., these fields are saved:  
**Month/Year, Room, Water Units, Water Amount, Electric Units, Electric Amount, Garbage, Room Fee, Fine, Grand Total**.

> Note: The **owner name** is used for display on the invoice (auto-filled from `Name`). Keep the `Name` sheet up to date.

### Troubleshooting & Notes
- Page setup, fonts, and row/column sizes are **yours to manage**. The macros do **not** modify styling.  
- If auto-calculation or events ever get disabled (e.g., after an interrupted macro), run the helper macro **`ResetAppState`** to restore them.  
- If the owner name does not appear:
  - Ensure **Room No.** on **Bill** matches exactly the value in **`Name`!A** (no extra spaces, same case/format).  
  - Confirm the `Name` sheet is spelled exactly **Name** (no trailing spaces).

---
