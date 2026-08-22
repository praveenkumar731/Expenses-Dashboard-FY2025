# 📊 Expenses Dashboard — FY 2025 (Excel VBA)

An automated, fully self-building expense-tracking dashboard for Excel, powered entirely by VBA. Run a single macro and it transforms a raw expenses table into a polished, interactive dashboard with KPI cards, charts, and slicers — no manual pivot-table or chart setup required.

## ✨ Features

- **One-click build** — run `BuildExpensesDashboard` and the entire dashboard is generated automatically
- **Live progress tracking** — status bar shows step-by-step build progress as it runs
- **6 KPI cards** — Total Expenses, Budget Allocated, Budget Remaining, Transactions, Approval Rate, Avg Expense
- **5 interactive charts** — Expense by Department, Expense by Category (doughnut), Monthly Trend, Budget vs Actual, Payment Method
- **7 PivotTables** — auto-generated on a dedicated `Pivot` sheet, feeding all charts and KPIs
- **Interactive slicers** — filter the entire dashboard by **Month** or **Department**
- **Auto-formatted data table** — converts raw data into a structured Excel Table with clean styling
- **Professional design** — green-themed corporate dashboard layout, print-ready formatting

---

## 📁 Repository Contents

| File | Description |
|---|---|
| `Expenses_Dashboard.xlsm` | Macro-enabled workbook containing the VBA code, sample data, and generated dashboard |
| `README.md` | This file |

---

## 🚀 How to Use

1. **Download** `Expenses_Dashboard.xlsm` from this repository.
2. **Open** it in Excel (Windows, Office 2016 or later).
3. When prompted, click **Enable Content / Enable Macros**.
4. Make sure your data lives on a sheet named **`Data`** with these exact column headers in Row 1:

   | Expense ID | Date | Department | Category | Amount (INR) | Payment Method | Approved | Budget Allocated | Budget Remaining |
   |---|---|---|---|---|---|---|---|---|

5. Press **Alt+F8**, select **`BuildExpensesDashboard`**, and click **Run**.
6. Watch the status bar for live progress — a confirmation popup appears when the build completes.
7. The dashboard is generated on a new **`Dashboard`** sheet, with supporting pivots on a **`Pivot`** sheet.

---

## 🛠️ Built With

- **Excel VBA** (Visual Basic for Applications)
- **PivotTables & PivotCharts**
- **Excel Slicers**
- **Excel Tables (ListObjects)**

---

## 📌 Notes

- Macros must be enabled for the dashboard to build — this is a `.xlsm` file, not `.xlsx`.
- Re-running the macro safely rebuilds the `Pivot` and `Dashboard` sheets from scratch, so it can be re-run any time the underlying data changes.
- Data is sample/demo data for FY 2025, generated for demonstration purposes.

---

## 📄 License

Feel free to use, modify, and adapt this project for personal or commercial use.
