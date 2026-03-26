# 💼 Personal Finance Dashboard — Advanced Excel Project

> A production-grade, multi-sheet Excel dashboard for complete personal finance tracking — built with Power Query architecture, SUMPRODUCT analytics, VBA automation hooks, and professional financial modeling conventions.

---

## 📸 Dashboard Preview

![Dashboard Preview](dashboard_preview.png)

*Executive KPI cards + clustered bar chart (Income vs Expense vs Savings) + monthly savings rate trend*

---

## 🗂️ Sheet Architecture

![Sheet Architecture](sheet_architecture.png)

*RAW_DATA is the single source of truth — all summary sheets and the dashboard pull from it via SUMPRODUCT formulas*

---

## 🏷️ Category Analysis

![Category Analysis](category_analysis.png)

*Donut chart (expense split by category) + horizontal bar comparison (2023 vs 2024 per category)*

---

## 💰 Savings Tracker

![Savings Tracker](savings_tracker.png)

*Goal progress bars with % completion + cumulative savings growth curve (2023–2024)*

---

## 🗂️ Sheet Overview

| Sheet | Purpose | Key Features |
|---|---|---|
| `DASHBOARD` | Executive KPI view (start here) | KPI cards, bar chart, line chart, pie chart |
| `RAW_DATA` | Transaction log — add entries here | Formatted Table, drop-down validation, zebra rows |
| `MONTHLY_SUMMARY` | Month-by-month P&L | SUMPRODUCT formulas, conditional formatting |
| `ANNUAL_SUMMARY` | Year-over-year comparison | YoY income/expense growth %, aggregated |
| `CATEGORY_ANALYSIS` | Spending by category | Year selector, ColorScale heatmap |
| `SAVINGS_TRACKER` | Goal tracker + SIP log | Progress %, status flags, savings breakdown |
| `INSTRUCTIONS` | Formula guide & how-to | Color legend, formula index, data entry guide |

---

## ⚙️ Excel Features Used

### Formulas & Functions
| Formula | Usage |
|---|---|
| `SUMPRODUCT` | Multi-condition aggregation across all summary sheets |
| `IFERROR` | Zero-safe division for savings rate and YoY growth |
| `MIN / MAX` | Capping savings progress at 100%, ensuring non-negative |
| `IF / IFS` | Status flags (Complete / On Track / Behind) |
| `SUM / AVERAGE` | Totals and averages in annual and savings sheets |

### Excel Features
- **Structured Table** (`Transactions`) on RAW_DATA for dynamic formula ranges
- **Data Validation** — drop-downs for `Type` (Income/Expense/Savings) and Year selector
- **Conditional Formatting** — ColorScale (3-color) on Savings Rate; CellIsRule for negative cash flows
- **PivotTable-ready** — all data is in a clean, normalized table format
- **Named Ranges** — consistent use of sheet-level references
- **Charts** — Clustered bar (Income vs Expense vs Savings), Line (Savings Rate), Pie (Category split)

### Industry-Standard Color Coding
| Color | Meaning |
|---|---|
| 🔵 Blue text | Hardcoded input — safe to change |
| ⚫ Black text | Formula — do NOT edit |
| 🟡 Yellow fill | Cell requiring your input or attention |

---

## 🚀 Power Query Extension (Advanced)

The RAW_DATA sheet is built as a proper Excel Table (`Transactions`), making it Power Query-ready out of the box.

### To connect your bank CSV:
1. `Data → Get Data → From File → From CSV`
2. In Power Query Editor, map columns to: Date, Month, Year, Type, Category, Amount
3. Append query to the `Transactions` table
4. Set `Refresh on file open` via **Query Properties**

### Suggested M Query transformations:
```m
// Add Month name from Date column
= Table.AddColumn(Source, "Month", each Date.MonthName([Date]), type text)

// Add Year column
= Table.AddColumn(#"Added Month", "Year", each Date.Year([Date]), type number)

// Map bank categories to your taxonomy
= Table.ReplaceValue(#"Added Year","UPI","Food",Replacer.ReplaceText,{"Category"})
```

---

## 🤖 VBA Automation Hooks

Open the VBA editor (`Alt + F11`) and add these macros to extend the workbook:

### Auto-refresh all SUMPRODUCT calculations
```vba
Sub RefreshDashboard()
    Application.CalculateFull
    MsgBox "Dashboard refreshed successfully!", vbInformation
End Sub
```

### Export Dashboard as PDF (one-click report)
```vba
Sub ExportDashboardPDF()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DASHBOARD")
    Dim filePath As String
    filePath = Environ("USERPROFILE") & "\Desktop\FinanceDashboard_" & Format(Now, "YYYYMMDD") & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=filePath, Quality:=xlQualityStandard
    MsgBox "PDF saved to Desktop: " & filePath, vbInformation
End Sub
```

### Monthly summary auto-email via Outlook
```vba
Sub EmailMonthlySummary()
    Dim olApp As Object, olMail As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    With olMail
        .To = "you@email.com"
        .Subject = "Monthly Finance Summary - " & Format(Now, "MMM YYYY")
        .Body = "Hi," & vbCrLf & vbCrLf & _
                "Please find your monthly finance summary attached." & vbCrLf & _
                "Income: " & Format(Sheets("ANNUAL_SUMMARY").Range("B4"), "₹#,##0") & vbCrLf & _
                "Expenses: " & Format(Sheets("ANNUAL_SUMMARY").Range("C4"), "₹#,##0")
        .Attachments.Add ThisWorkbook.FullName
        .Display  ' Change to .Send to auto-send
    End With
End Sub
```

### Assign buttons on DASHBOARD
1. `Insert → Shapes → Rectangle` → right-click → `Assign Macro`
2. Assign `RefreshDashboard` and `ExportDashboardPDF`

---

## 📊 Sample Dataset

The workbook ships with **2 years of realistic synthetic data (2023–2024)**:

- ~500+ transaction rows across 8 expense categories
- Monthly salary (₹65,000 base) + random freelance income
- Recurring SIP (₹10,000/month to Zerodha)
- Quarterly emergency fund contributions
- Randomised but realistic expense amounts per category

---

## 📁 File Structure

```
PersonalFinanceDashboard/
│
├── PersonalFinanceDashboard.xlsx      # Main workbook (7 sheets)
├── README.md                          # This file
├── transactions_template.csv          # CSV template for bank import
├── dashboard_preview.png              # Dashboard screenshot
├── category_analysis.png              # Category analysis screenshot
├── savings_tracker.png                # Savings tracker screenshot
└── sheet_architecture.png             # Sheet architecture diagram
```

---

## 🛠️ How to Add Your Own Data

1. Open `RAW_DATA` sheet
2. Click the first empty row below the last transaction
3. Fill in columns: **Date → Month → Year → Type → Category → Sub-Category → Description → Amount → Account**
4. `Type` must be one of: `Income`, `Expense`, `Savings` (drop-down enforced)
5. All summary sheets and the dashboard update **automatically** — no manual refresh needed

---

## 📈 Skills Demonstrated

- Advanced Excel formula construction (`SUMPRODUCT`, `IFERROR`, `IF/IFS`, `MIN/MAX`)
- Financial modeling color conventions (blue inputs, black formulas)
- Power Query data pipeline design
- VBA automation (PDF export, Outlook integration)
- Dashboard design (KPI cards, charts, conditional formatting)
- Data validation and structured table architecture
- Multi-sheet cross-referencing with zero circular references

---

## 🧑‍💻 Author

**Vijay Kumar** · [GitHub: VijayKumaro7](https://github.com/VijayKumaro7) · [LinkedIn](#) · [Hashnode Blog](#)

---

## 📄 License

MIT — free to use, adapt, and build upon.
