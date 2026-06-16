# Excel Automation & Data Entry Projects

A portfolio of **5 end-to-end data automation projects** spanning PDF data extraction, Excel data cleaning, Word document generation, interactive dashboarding, and a full VBA employee management system.

---

## Projects Overview

| # | Project | Technology | What It Does |
|---|---------|-----------|--------------|
| 1 | [PDF to Excel Conversion](./Task%201%20-%20PDF%20to%20excel%20conversion) | VBA / Excel | Extracts 50 customer records from a PDF into a structured Excel workbook |
| 2 | [Excel Cleanup](./Task%202%20%E2%80%94%20Excel%20Cleanup) | Excel (manual/formula) | Cleans 100 messy customer entries — fixes names, emails, dates, phones, removes duplicates |
| 3 | [Business Report Word Formatting](./Task%203%20-%20Business%20report%20word%20formatting) | Python (win32com) | Generates a 12-section professional Word report with OLE-embedded Excel charts, tables, and cover page |
| 4 | [Sales Dashboard](./Task%204%20%E2%80%94%20Sales%20Dashboard) | Excel + VBA | 7-sheet executive dashboard with KPI cards, charts, FORECAST.ETS, what-if analysis, and slicer filtering |
| 5 | [Employee Management VBA](./Task%205%20-%20Employee%20management%20VBA%20Task) | Excel VBA | Full CRUD application with UserForm GUI, input validation, audit logging, and auto-ID generation |

---

## Task Details

### Task 1 — PDF to Excel Conversion
- **Source:** 5-page PDF with 50 dummy customer records (Name, Email, Phone, City)
- **Output:** `customer_records.xlsx` / `.xlsm` — structured Excel table
- **Key skill:** Data extraction and tabular transformation

### Task 2 — Excel Cleanup
- **Source:** `Messy_Excel_Cleanup_Task_100_Entries.xlsx` (111 rows, intentional inconsistencies)
- **Issues fixed:** Case inconsistencies, whitespace, mixed date formats (`05.07.2024`, `July 8, 2024`, `2024/03/04`), `+91-` phone prefixes, duplicate records
- **Output:** `Cleaned_Customer_Data.xlsx` (101 rows, 100 unique records, standardized formatting)

### Task 3 — Business Report Word Formatting
- **Script:** `build_final.py` (663 lines, Python + win32com)
- **Output:** 12-section Word document with:
  - Cover page, table of contents, executive summary
  - Regional revenue analysis with OLE-embedded Excel bar/pie/line charts
  - KPI tables with conditional formatting (green/amber/red status)
  - Goal progress bars (Unicode block characters)
  - Page X of Y footers, consistent Navy/Gold theme
- **Charts are live OLE objects** — double-click to edit in Excel

### Task 4 — Sales Dashboard
- **Source:** 12 months of sales data (Revenue, Expenses, Leads, Orders)
- **7 sheets:** Source Data → Calculations → Charts → Dashboard → Forecast → What-If → Quarterly View
- **Features:**
  - 5 KPI cards with sparklines + dynamic title (shows live profit margin)
  - 4 chart types: line, column, combo, waterfall
  - `FORECAST.ETS` with 90% confidence bands (3-month projection)
  - What-if scenario modeling (3%/5%/8%/10% compounding)
  - Excel Table + slicer for interactive quarterly filtering
  - VBA macro `ExportDashboardToPDF` for one-click PDF export

### Task 5 — Employee Management VBA
- **Stack:** Excel VBA (4 modules + 1 UserForm, ~974 lines)
- **Modules:**
  - `Module_Validation.bas` — email, phone, salary, ID, duplicate validation
  - `Module_Utilities.bas` — auto-ID generation, logging, sheet formatting
  - `Module_EmployeeCRUD.bas` — Add, Search, Update, Delete, List All
  - `frmEmployeeManagement` — Professional GUI with search, list view, action buttons
- **Features:**
  - `tblEmployees` formatted table with Indian Rupee salary formatting
  - Audit log (Timestamp, Action, Employee ID) on a hidden "Logs" sheet
  - Click-to-populate from list, confirmation dialogs, keyboard Enter for search
  - 9 pre-loaded sample employees across 6 departments

---

## Technologies Used

| Technology | Usage |
|-----------|-------|
| **Python** | `win32com` for Word/Excel COM automation (Task 3) |
| **VBA** | UserForms, Excel object model, COM automation (Tasks 1, 4, 5) |
| **Excel** | Tables, PivotTables, FORECAST.ETS, conditional formatting, data validation, slicers, camera tool |
| **Word** | OLE object embedding, field codes, styles, headers/footers |
| **PDF** | Source data extraction target |

---

## Repository Structure

```
Excel-automation-data-entry-projects/
├── Task 1 - PDF to excel conversion/
│   ├── Dummy_Customer_Records_PDF_Conversion_Task.pdf
│   ├── customer_records.xlsx
│   └── customer_records.xlsm
├── Task 2 — Excel Cleanup/
│   ├── Messy_Excel_Cleanup_Task_100_Entries.xlsx
│   └── Cleaned_Customer_Data.xlsx
├── Task 3 - Business report word formatting/
│   ├── build_final.py
│   ├── Raw_Word_Formatting_Task.docx
│   ├── Professional_Business_Report_2024.docx
│   └── Professional_Business_Report_2024_v2.docx
├── Task 4 — Sales Dashboard/
│   ├── Sales_Dashboard_Source_Data.xlsx
│   ├── Sales_Dashboard.xlsx
│   ├── Sales_Dashboard.xlsm
│   ├── ExportDashboardToPDF.bas
│   └── Dashboard_Export_2026-06-15.pdf
├── Task 5 - Employee management VBA Task/
│   ├── Employee_Management_VBA_Task.xlsx
│   ├── Employee_Management_VBA_Task.xlsm
│   ├── vba-code/
│   │   ├── Module_Validation.bas
│   │   ├── Module_Utilities.bas
│   │   ├── Module_EmployeeCRUD.bas
│   │   ├── frmEmployeeManagement.frm
│   │   └── frmEmployeeManagement_Code.bas
│   └── screenshots/
└── README.md
```

---

## Quick Start

1. **Excel projects (Tasks 1, 2, 4, 5):** Open the `.xlsx` or `.xlsm` files in Excel. Enable macros where prompted (`.xlsm` files).
2. **Word report (Task 3):** Run `python build_final.py` from the Task 3 directory (requires `pywin32` on Windows with Word installed).
3. **VBA source code (Task 5):** Browse `vba-code/` — each module can be imported into the VBA editor via `File > Import File`.

---

*All projects were built as hands-on data automation exercises demonstrating real-world data entry, cleaning, reporting, and application development workflows.*
