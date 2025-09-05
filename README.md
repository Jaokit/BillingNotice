# How This Workbook Works (Excel/VBA)

This workbook generates **dormitory rent & utilities invoices** directly in Excel.  
All calculations run automatically on the **Bill** sheet; a single click prints and logs the bills.

## Workflow

1) **Open & Enable Macros**  
   Open the file and allow macros.

2) **Fill 3 Panels on the _Bill_ sheet** (Top / Middle / Bottom)  
   Per panel, enter:
   - **Room No.**  
   - **Billing Date**  
   - **Water units**  
   - **Electric units**  
   - **Room fee** (only if the room is in A1–A12; other rooms are auto-priced)  
   - **Fine** (optional)

3) **Auto-calculation (instant)**  
   As you type, the workbook calculates and displays:
   - **Water** = units × **28**  
   - **Electricity** = units × **10**  
   - **Garbage** = **฿20** (fixed)  
   - **Room fee** = auto per room rules (or manual for A1–A12)  
   - **Grand total** = sum of all amounts  
   All money fields use **Thai baht format (฿#,##0)**.

4) **Print & Save (one click)**  
   Press the button assigned to `SaveAllPanelsToHistorAndPrint`. It will:
   - Recalculate all 3 panels,  
   - **Append** one row per filled panel to the **Histor** sheet (starts at row 2; row 1 is your own headers),  
   - **Print** one page (3 bills),  
   - **Clear** the inputs for the next use (including Room/Date).

## Histor Sheet (what gets saved)

For each panel with a Room No., the following fields are saved:
**Month/Year, Room, Water Units, Water Amount, Electric Units, Electric Amount, Garbage, Room Fee, Fine, Grand Total**.

## Notes

- Page setup, fonts, and column/row sizes are **managed by you**; macros do not change styling.  
- If auto-calculation ever stops (e.g., after an interrupted macro), run the helper macro **`ResetAppState`** to re-enable events and auto-calc.
