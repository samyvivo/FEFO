
# FEFO / Batch Inventory / Expiry Monitoring

A desktop application (Python + Tkinter) that analyzes **sales** and **stock** Excel files to:

- Detect **FEFO (First Expired, First Out) violations** at batch, expiry-group and product level,
- Quantify violation volume in units,
- Monitor **expiry risk** (Expired / Critical / Warning / Watch / OK),
- Produce a **multi-sheet itemized Excel report** with pre-colored, ready-to-share views.

The UI supports **Persian / Jalali dates** (via `jdatetime`) and Persian column headers.

---

## ✨ Features

- **Graphical interface** — pick input files, output path, report date, and stock mode from a simple Tkinter window.
- **Two stock modes**
  - `Stock = current remaining AFTER sales` (typical)
  - `Stock = BEFORE sales`
- **FEFO violation detection**
  - Compares per-expiry sold quantities against the strict FEFO expectation,
  - Quantifies oversold units (violation units), not just a binary flag,
  - Distinguishes strictly earlier expiries (same-expiry batches form a single pool).
- **Expiry buckets** with per-status color coding:| Status            | Condition                  |
  | ----------------- | -------------------------- |
  | Expired           | `days_to_expiry < 0`     |
  | Critical (<=30d)  | `0 ≤ days ≤ 30`        |
  | Warning (<=90d)   | `30 < days ≤ 90`        |
  | Watch (<=180d)    | `90 < days ≤ 180`       |
  | OK                | `days > 180`             |
  | No/invalid expiry | missing or unparsable date |
- **Two violation-rate metrics** on the Summary sheet
  - `FEFO Violation Rate (%)` — violation units ÷ (sold units of *compliant* multi-batch products + violation units)
  - `Overall FEFO violation rate (%)` — violation units ÷ total sold units
- **Flexible column matching** — Persian headers are matched after stripping ZWNJ (`\u200c`) and RTL/LTR marks.
- **Persian-digit normalization** — `۰۱۲۳` and `٠١٢٣` are converted before parsing numbers / dates.
- **Handles both Gregorian and Jalali dates** — detects year magnitude and converts accordingly.
- **Excel report** with auto-fit column widths, bold headers and frozen top row.

---

## 📦 Requirements

- Python **3.9+**
- Packages:

```bash
pip install pandas numpy openpyxl jdatetime
```
