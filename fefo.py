#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FEFO / Batch Inventory / Expiry Monitoring - GUI + itemized Excel report.
"""

import threading
from pathlib import Path

import numpy as np
import pandas as pd

try:
    import jdatetime
except ImportError:
    jdatetime = None

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import font as tkfont


# ============================================================
# Utilities
# ============================================================
PERSIAN_DIGITS = str.maketrans(
    "۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩",
    "01234567890123456789",
)

ZWNJ = "\u200c"


def normalize_colname(s):
    return (
        str(s)
        .replace(ZWNJ, "")
        .replace("\u200f", "")
        .replace("\u200e", "")
        .strip()
    )


def clean_text(x):
    if pd.isna(x):
        return ""
    return str(x).translate(PERSIAN_DIGITS).replace(ZWNJ, " ").strip()


def clean_code(x):
    if pd.isna(x):
        return ""
    s = str(x).translate(PERSIAN_DIGITS).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def to_num(x):
    if pd.isna(x):
        return 0.0
    s = str(x).translate(PERSIAN_DIGITS).replace(",", "").strip()
    return pd.to_numeric(s, errors="coerce")


def parse_date(x):
    if pd.isna(x):
        return pd.NaT
    s = clean_text(x).replace("-", "/").replace(".", "/")
    if not s:
        return pd.NaT
    parts = s.split("/")
    if len(parts) != 3:
        return pd.to_datetime(s, errors="coerce")
    try:
        a, b, c = [int(p) for p in parts]
    except ValueError:
        return pd.NaT
    if a > 1700:
        try:
            return pd.Timestamp(a, b, c)
        except Exception:
            return pd.NaT
    if c > 1700:
        try:
            return pd.Timestamp(c, b, a)
        except Exception:
            return pd.NaT
    if jdatetime is not None:
        try:
            return pd.Timestamp(jdatetime.date(a, b, c).togregorian())
        except Exception:
            return pd.NaT
    return pd.NaT


def expiry_status(days):
    if pd.isna(days):
        return "No/invalid expiry"
    if days < 0:
        return "Expired"
    if days <= 30:
        return "Critical (<=30d)"
    if days <= 90:
        return "Warning (<=90d)"
    if days <= 180:
        return "Watch (<=180d)"
    return "OK"


def find_col(df, candidates):
    """Flexible column matcher ignoring ZWNJ / RTL marks."""
    norm_map = {normalize_colname(c): c for c in df.columns}
    for cand in candidates:
        key = normalize_colname(cand)
        if key in norm_map:
            return norm_map[key]
    return None


def get_required(df, spec, label):
    missing = []
    mapping = {}
    for human, candidates in spec.items():
        col = find_col(df, candidates)
        if col is None:
            missing.append(human)
        else:
            mapping[human] = col
    if missing:
        raise ValueError(
            f"{label} file is missing columns: {missing}\n"
            f"Found: {list(df.columns)}"
        )
    return mapping


# ============================================================
# Analysis
# ============================================================
def add_fefo_flags(df, tol=1e-6):
    """
    Add FEFO flags AND quantify the violation in units.
    """
    df = df.copy()
    df["expiry_sort"] = df["expiry"].fillna(pd.Timestamp("2100-01-01"))
    df = df.sort_values(["code", "expiry_sort", "batch"]).reset_index(drop=True)

    # Batch-level cumulative (for display purposes in reports)
    df["cum_opening_before"] = (
        df.groupby("code")["opening_qty"].cumsum() - df["opening_qty"]
    )
    df["cum_sold_before"] = (
        df.groupby("code")["sold_qty"].cumsum() - df["sold_qty"]
    )

    # Expiry-level aggregation: batches sharing the same expiry form one pool
    expiry_agg = df.groupby(["code", "expiry_sort"]).agg(
        opening_qty=("opening_qty", "sum"),
        sold_qty=("sold_qty", "sum"),
    ).reset_index()
    expiry_agg = expiry_agg.sort_values(["code", "expiry_sort"]).reset_index(drop=True)

    expiry_agg["cum_opening_before_exp"] = (
        expiry_agg.groupby("code")["opening_qty"].cumsum() - expiry_agg["opening_qty"]
    )
    expiry_agg["cum_sold_before_exp"] = (
        expiry_agg.groupby("code")["sold_qty"].cumsum() - expiry_agg["sold_qty"]
    )

    # Only STRICTLY earlier expiry dates contribute to "earlier unconsumed"
    expiry_agg["earlier_unconsumed_qty"] = (
        expiry_agg["cum_opening_before_exp"] - expiry_agg["cum_sold_before_exp"]
    ).clip(lower=0)

    expiry_agg["cum_opening_after_exp"] = expiry_agg.groupby("code")["opening_qty"].cumsum()
    expiry_agg["total_sold_product"] = expiry_agg.groupby("code")["sold_qty"].transform("sum")

    # FEFO expected sold quantity (per expiry group)
    expiry_agg["expected_sold_qty"] = (
        np.minimum(expiry_agg["total_sold_product"], expiry_agg["cum_opening_after_exp"])
        - np.minimum(expiry_agg["total_sold_product"], expiry_agg["cum_opening_before_exp"])
    ).clip(lower=0)

    expiry_agg["sold_vs_fefo_diff"] = expiry_agg["sold_qty"] - expiry_agg["expected_sold_qty"]

    # FEFO violation units per expiry group
    expiry_agg["group_violation_units"] = expiry_agg["sold_vs_fefo_diff"].clip(lower=0)

    expiry_agg["expiry_violation"] = (
        (expiry_agg["sold_vs_fefo_diff"] > tol)
        | ((expiry_agg["sold_qty"] > tol) & (expiry_agg["earlier_unconsumed_qty"] > tol))
    )

    expiry_merge = expiry_agg[[
        "code", "expiry_sort", "earlier_unconsumed_qty", "expected_sold_qty",
        "sold_vs_fefo_diff", "expiry_violation",
        "group_violation_units", "sold_qty",
    ]].rename(columns={"sold_qty": "group_sold_qty"})

    df = df.merge(expiry_merge, on=["code", "expiry_sort"], how="left")
    df["group_sold_qty"] = df["group_sold_qty"].fillna(0.0)
    df["group_violation_units"] = df["group_violation_units"].fillna(0.0)

    # Allocate group violation to the batches that actually sold
    df["violation_units"] = np.where(
        df["group_sold_qty"] > tol,
        df["sold_qty"] / df["group_sold_qty"] * df["group_violation_units"],
        0.0,
    )

    df = df.sort_values(["code", "expiry_sort", "batch"]).reset_index(drop=True)

    # Batch-level violation flag
    df["fefo_violation"] = (df["sold_qty"] > tol) & (
        (df["earlier_unconsumed_qty"] > tol) | (df["expiry_violation"] == True)
    )
    return df


def run_analysis(sales_path, stock_path, stock_is_current_remaining, report_date):
    sales = pd.read_excel(sales_path, dtype=str, engine="openpyxl")
    stock = pd.read_excel(stock_path, dtype=str, engine="openpyxl")
    sales.columns = [str(c).strip() for c in sales.columns]
    stock.columns = [str(c).strip() for c in stock.columns]

    smap = get_required(sales, {
        "batch": ["سری ساخت"],
        "expiry": ["تاریخ انقضاء", "تاریخ انقضا"],
        "code": ["کد کالا"],
        "name": ["نام کالا"],
        "qty": ["تعداد فروش و جایزه-ظرف"],
    }, "Sales")

    kmap = get_required(stock, {
        "code": ["کد کالا"],
        "name": ["نام کالا"],
        "qty": ["موجودی"],
        "batch": ["شماره بچ"],
        "expiry": ["تاریخ انقضا", "تاریخ انقضاء"],
    }, "Stock")

    sales_n = pd.DataFrame({
        "code": sales[smap["code"]].apply(clean_code),
        "name": sales[smap["name"]].apply(clean_text),
        "batch": sales[smap["batch"]].apply(clean_text),
        "expiry": sales[smap["expiry"]].apply(parse_date),
        "qty": sales[smap["qty"]].apply(to_num),
    })
    sales_n = sales_n[(sales_n["code"] != "") & (sales_n["qty"] != 0)].copy()

    stock_n = pd.DataFrame({
        "code": stock[kmap["code"]].apply(clean_code),
        "name": stock[kmap["name"]].apply(clean_text),
        "batch": stock[kmap["batch"]].apply(clean_text),
        "expiry": stock[kmap["expiry"]].apply(parse_date),
        "qty": stock[kmap["qty"]].apply(to_num),
    })
    stock_n = stock_n[stock_n["code"] != ""].copy()

    sales_agg = (
        sales_n.groupby(["code", "batch", "expiry"], dropna=False)
        .agg(name=("name", "first"), sold_qty=("qty", "sum"),
             sales_rows=("qty", "size"))
        .reset_index()
    )
    
    # Remove sales_rows column from sales_agg
    sales_agg = sales_agg.drop(columns=["sales_rows"])

    stock_agg = (
        stock_n.groupby(["code", "batch", "expiry"], dropna=False)
        .agg(name=("name", "first"), stock_qty=("qty", "sum"),
             stock_rows=("qty", "size"))
        .reset_index()
    )

    df = pd.merge(
        stock_agg, sales_agg,
        on=["code", "batch", "expiry"], how="outer",
        suffixes=("_stock", "_sales"),
    )
    df["name_stock"] = df["name_stock"].fillna("").astype(str)
    df["name_sales"] = df["name_sales"].fillna("").astype(str)
    df["name"] = df["name_stock"].where(df["name_stock"].ne(""), df["name_sales"])
    df["stock_qty"] = df["stock_qty"].fillna(0.0)
    df["sold_qty"] = df["sold_qty"].fillna(0.0)
    df["stock_rows"] = df["stock_rows"].fillna(0).astype(int)

    if stock_is_current_remaining:
        df["opening_qty"] = df["stock_qty"] + df["sold_qty"]
        df["remaining_qty"] = df["stock_qty"]
    else:
        df["opening_qty"] = df["stock_qty"]
        df["remaining_qty"] = df["stock_qty"] - df["sold_qty"]

    df["data_quality_note"] = ""
    df.loc[df["batch"].eq(""), "data_quality_note"] += "missing batch; "
    df.loc[df["expiry"].isna(), "data_quality_note"] += "missing/invalid expiry; "

    df = add_fefo_flags(df)

    df["days_to_expiry"] = (df["expiry"] - report_date).dt.days
    df["expiry_status"] = df["days_to_expiry"].apply(expiry_status)

    # Define standard sorting for expiry statuses
    status_sort_map = {
        "Expired": 1,
        "Critical (<=30d)": 2,
        "Warning (<=90d)": 3,
        "Watch (<=180d)": 4,
        "OK": 5,
        "No/invalid expiry": 6
    }

    # ----- Batch inventory -----
    batch_inventory = df.rename(columns={
        "stock_qty": "stock_qty_current",
        "opening_qty": "opening_qty_for_fefo",
    })[[
        "code", "name", "batch", "expiry",
        "stock_qty_current", "sold_qty", "opening_qty_for_fefo", "remaining_qty",
        "days_to_expiry", "expiry_status", "fefo_violation",
        "expected_sold_qty", "sold_vs_fefo_diff", "earlier_unconsumed_qty",
        "data_quality_note",
    ]].copy()
    
    # Sort batch_inventory
    batch_inventory["_sort"] = batch_inventory["expiry_status"].map(status_sort_map).fillna(99)
    batch_inventory = batch_inventory.sort_values(["_sort", "days_to_expiry", "code", "expiry"]).drop(columns=["_sort"]).reset_index(drop=True)

    # ----- Expiry monitoring (full) -----
    expiry_monitoring = df[
        (df["stock_qty"] > 0) | (df["sold_qty"] > 0)
    ][[
        "code", "name", "batch", "expiry",
        "stock_qty", "sold_qty", "days_to_expiry", "expiry_status",
        "data_quality_note",
    ]].copy()

    # Sort expiry_monitoring
    expiry_monitoring["_sort"] = expiry_monitoring["expiry_status"].map(status_sort_map).fillna(99)
    expiry_monitoring = expiry_monitoring.sort_values(["_sort", "days_to_expiry", "code", "expiry"]).drop(columns=["_sort"]).reset_index(drop=True)

    # ----- FEFO compliance (product level) -----
    sales_codes = set(sales_agg["code"].unique())
    fefo_compliance = (
        df.groupby("code")
        .agg(
            name=("name", "first"),
            total_stock_qty=("stock_qty", "sum"),
            total_sold_qty=("sold_qty", "sum"),
            batch_count=("batch", "nunique"),
            unique_expiries=("expiry", "nunique"), 
            violation_rows=("fefo_violation", "sum"),
            total_violation_units=("violation_units", "sum"),
            first_expiry=("expiry", "min"),
            last_expiry=("expiry", "max"),
        )
        .reset_index()
    )
    fefo_compliance = fefo_compliance[fefo_compliance["code"].isin(sales_codes)].copy()
    fefo_compliance["fefo_status"] = np.where(
        fefo_compliance["violation_rows"] > 0, "VIOLATION", "COMPLIANT"
    )
    fefo_compliance["violation_rate_pct"] = np.where(
        fefo_compliance["total_sold_qty"] > 0,
        (fefo_compliance["total_violation_units"]
         / fefo_compliance["total_sold_qty"] * 100).round(2),
        0.0,
    )

    fefo_compliance = fefo_compliance.drop(columns=["violation_rows"])

    fefo_violation_products = fefo_compliance[
        fefo_compliance["fefo_status"] == "VIOLATION"
    ].reset_index(drop=True)
    
    fefo_compliant_products = fefo_compliance[
        fefo_compliance["fefo_status"] == "COMPLIANT"
    ].drop(columns=["total_violation_units", "violation_rate_pct"]).reset_index(drop=True)

    # ----- Category slices -----
    def slice_batches(mask):
        return df[mask][[
            "code", "name", "batch", "expiry", "days_to_expiry",
            "stock_qty", "sold_qty", "expiry_status"
        ]].sort_values(["code", "expiry"]).reset_index(drop=True)

    expired = slice_batches((df["stock_qty"] > 0) & (df["days_to_expiry"] < 0))
    critical = slice_batches(
        (df["stock_qty"] > 0)
        & (df["days_to_expiry"] >= 0)
        & (df["days_to_expiry"] <= 30)
    )
    warning = slice_batches(
        (df["stock_qty"] > 0)
        & (df["days_to_expiry"] > 30)
        & (df["days_to_expiry"] <= 90)
    )
    watch = slice_batches(
        (df["stock_qty"] > 0)
        & (df["days_to_expiry"] > 90)
        & (df["days_to_expiry"] <= 180)
    )

    # ----- Products by Status -----
    s = df[df["stock_qty"] > 0].copy()
    rows = []
    for code, g in s.groupby("code"):
        name = g["name"].iloc[0]
        expired_qty = g.loc[g["days_to_expiry"] < 0, "stock_qty"].sum()
        crit_qty = g.loc[(g["days_to_expiry"] >= 0) & (g["days_to_expiry"] <= 30),
                         "stock_qty"].sum()
        warn_qty = g.loc[(g["days_to_expiry"] > 30) & (g["days_to_expiry"] <= 90),
                         "stock_qty"].sum()
        watch_qty = g.loc[(g["days_to_expiry"] > 90) & (g["days_to_expiry"] <= 180),
                          "stock_qty"].sum()
        ok_qty = g.loc[g["days_to_expiry"] > 180, "stock_qty"].sum()
        min_days = g["days_to_expiry"].min()
        total = g["stock_qty"].sum()
        
        if expired_qty > 0:
            status = "Expired"
        elif crit_qty > 0:
            status = "Critical (<=30d)"
        elif warn_qty > 0:
            status = "Warning (<=90d)"
        elif watch_qty > 0:
            status = "Watch (<=180d)"
        elif ok_qty > 0:
            status = "OK"
        else:
            status = "No exp date or batch"
            
        rows.append({
            "code": code, "name": name, "status": status,
            "total_stock": total,
            "min_days_to_expiry": min_days,
            "expired_qty": expired_qty,
            "critical_qty_30d": crit_qty,
            "warning_qty_90d": warn_qty,
            "watch_qty_180d": watch_qty,
            "ok_qty": ok_qty,
            "batch_count": g["batch"].nunique(),
        })
        
    if rows:
        df_ps = pd.DataFrame(rows)
        status_order = {
            "Expired": 1,
            "Critical (<=30d)": 2,
            "Warning (<=90d)": 3,
            "Watch (<=180d)": 4,
            "OK": 5,
            "No exp date or batch": 6,
        }
        df_ps["status_order"] = df_ps["status"].map(status_order).fillna(99)
        products_status = df_ps.sort_values(["status_order", "min_days_to_expiry"]).drop(columns=["status_order"]).reset_index(drop=True)
    else:
        products_status = pd.DataFrame(columns=[
            "code", "name", "status", "total_stock", "min_days_to_expiry",
            "expired_qty", "critical_qty_30d", "warning_qty_90d",
            "watch_qty_180d", "ok_qty", "batch_count",
        ])

    # Attach FEFO status AND violation units to products_status
    if not products_status.empty:
        products_status = products_status.merge(
            fefo_compliance[["code", "fefo_status",
                             "total_violation_units", "violation_rate_pct"]],
            on="code", how="left",
        )
        products_status["fefo_status"] = products_status["fefo_status"].fillna(
            "No sales"
        )
        products_status["total_violation_units"] = (
            products_status["total_violation_units"].fillna(0.0)
        )
        products_status["violation_rate_pct"] = (
            products_status["violation_rate_pct"].fillna(0.0)
        )

    # ============================================================
    # METRICS CALCULATION
    # ============================================================
    total_violation_units = float(df["violation_units"].sum())
    
    compliant_gt2 = fefo_compliance[
        (fefo_compliance["fefo_status"] == "COMPLIANT") &
        (fefo_compliance["batch_count"] > 2) &
        (fefo_compliance["unique_expiries"] > 1) 
    ].copy()
    
    denominator_units = float(compliant_gt2["total_sold_qty"].sum())
    denominator_products_count = int(len(compliant_gt2))
    
    new_fefo_rate = (
        (total_violation_units / (denominator_units + total_violation_units)  * 100)
        if denominator_units > 0 else 0.0
    )

    total_sold_all = float(sales_agg["sold_qty"].sum())
    old_overall_rate = (
        total_violation_units / total_sold_all * 100
        if total_sold_all > 0 else 0.0
    )

    # ----- Summary (Strictly Numbers in the Value column) -----
    summary_rows = [
        ("Report date (YYYYMMDD)", report_date.strftime("%Y/%m/%d")),
        ("Stock mode (1=after, 0=before)", 1 if stock_is_current_remaining else 0),
        ("Total sales quantity (all products)", float(sales_agg["sold_qty"].sum())),
        ("Unique product codes in sales", int(sales_agg["code"].nunique())),
        ("Unique batches in stock with qty > 0", int((df["stock_qty"] > 0).sum())),

        # === FEFO METRICS ===
        ("Total FEFO violation units", round(total_violation_units, 2)),
        ("Total sold of compliant products (>2 batches & not 1 pool)", round(denominator_units, 2)),
        ("FEFO Violation Rate (%)", round(new_fefo_rate, 2)),
        ("Count of compliant products matching criteria", denominator_products_count),
        
        # === OLD METRICS ===
        ("Overall FEFO violation rate (%)", round(old_overall_rate, 2)),

        ("FEFO compliant products (count)", int((fefo_compliance["fefo_status"] == "COMPLIANT").sum())),
        ("FEFO violation products (count)", int((fefo_compliance["fefo_status"] == "VIOLATION").sum())),

        ("Expired stock batch rows", int(expired.shape[0])),
        ("Expired stock quantity", float(expired["stock_qty"].sum())),

        ("Critical (<=30d) batch rows", int(critical.shape[0])),
        ("Critical (<=30d) quantity", float(critical["stock_qty"].sum())),

        ("Warning (31-90d) batch rows", int(warning.shape[0])),
        ("Warning (31-90d) quantity", float(warning["stock_qty"].sum())),

        ("Watch (91-180d) batch rows", int(watch.shape[0])),
        ("Watch (91-180d) quantity", float(watch["stock_qty"].sum())),
    ]
    summary = pd.DataFrame(summary_rows, columns=["Metric", "Value"])

    return {
        "summary": summary,
        "products_status": products_status,
        "expired": expired,
        "critical": critical,
        "warning": warning,
        "watch": watch,
        "fefo_violation_products": fefo_violation_products,
        "fefo_compliant": fefo_compliant_products,
        "batch_inventory": batch_inventory,
        "expiry_monitoring": expiry_monitoring,
        "sales_agg": sales_agg,
    }


# ============================================================
# Excel writing
# ============================================================
SHEET_ORDER = [
    ("summary", "Summary"),
    ("products_status", "Products_Status"),
    ("expired", "Expired_Stock"),
    ("critical", "Critical_30d"),
    ("warning", "Warning_31_90d"),
    ("watch", "Watch_91_180d"),
    ("fefo_violation_products", "FEFO_Violation_Products"),
    ("fefo_compliant", "FEFO_Compliant_Products"),
    ("batch_inventory", "Batch_Inventory"),
    ("expiry_monitoring", "Expiry_Monitoring"),
    ("sales_agg", "Sales_Aggregated"),
]


def write_excel(data, path):
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for key, sheet_name in SHEET_ORDER:
            if key in data:
                data[key].to_excel(
                    writer, sheet_name=sheet_name[:31], index=False
                )

        for ws in writer.book.worksheets:
            for col_idx in range(1, ws.max_column + 1):
                letter = get_column_letter(col_idx)
                max_len = 8
                for row_idx in range(1, ws.max_row + 1):
                    v = ws.cell(row=row_idx, column=col_idx).value
                    if v is not None:
                        l = len(str(v))
                        if l > max_len:
                            max_len = l
                ws.column_dimensions[letter].width = min(max_len + 2, 70)
            for cell in ws[1]:
                cell.font = Font(bold=True)
            ws.freeze_panes = "A2"


# ============================================================
# GUI
# ============================================================
class FEFOApp:
    def __init__(self, root):
        self.root = root
        root.title("FEFO / Batch Inventory / Expiry Monitoring")
        root.geometry("1320x880")

        try:
            default_font = tkfont.nametofont("TkDefaultFont")
            default_font.configure(family="Tahoma", size=10)
        except Exception:
            pass

        self.data = None
        self._build()

    def _build(self):
        pad = {"padx": 6, "pady": 4}

        form = ttk.LabelFrame(self.root, text="Inputs")
        form.pack(fill="x", **pad)

        ttk.Label(form, text="Sales file:").grid(row=0, column=0, sticky="e", **pad)
        self.sales_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.sales_var, width=80).grid(
            row=0, column=1, sticky="we", **pad
        )
        ttk.Button(form, text="Browse...", command=self._browse_sales).grid(
            row=0, column=2, **pad
        )

        ttk.Label(form, text="Stock file:").grid(row=1, column=0, sticky="e", **pad)
        self.stock_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.stock_var, width=80).grid(
            row=1, column=1, sticky="we", **pad
        )
        ttk.Button(form, text="Browse...", command=self._browse_stock).grid(
            row=1, column=2, **pad
        )

        ttk.Label(form, text="Output file:").grid(row=2, column=0, sticky="e", **pad)
        self.out_var = tk.StringVar(value="FEFO_Report.xlsx")
        ttk.Entry(form, textvariable=self.out_var, width=80).grid(
            row=2, column=1, sticky="we", **pad
        )
        ttk.Button(form, text="Browse...", command=self._browse_out).grid(
            row=2, column=2, **pad
        )

        ttk.Label(form, text="Report date (YYYY-MM-DD):").grid(
            row=3, column=0, sticky="e", **pad
        )
        self.date_var = tk.StringVar(
            value=pd.Timestamp.today().strftime("%Y-%m-%d")
        )
        ttk.Entry(form, textvariable=self.date_var, width=20).grid(
            row=3, column=1, sticky="w", **pad
        )

        mode = ttk.Frame(form)
        mode.grid(row=4, column=0, columnspan=3, sticky="w", **pad)
        ttk.Label(mode, text="Stock mode:").pack(side="left", padx=4)
        self.stock_mode_var = tk.StringVar(value="current")
        ttk.Radiobutton(
            mode,
            text="Stock = current remaining AFTER sales (typical)",
            variable=self.stock_mode_var,
            value="current",
        ).pack(side="left", padx=6)
        ttk.Radiobutton(
            mode,
            text="Stock = BEFORE sales",
            variable=self.stock_mode_var,
            value="before",
        ).pack(side="left", padx=6)

        form.columnconfigure(1, weight=1)

        bar = ttk.Frame(self.root)
        bar.pack(fill="x", **pad)
        self.run_btn = ttk.Button(
            bar, text="Run Analysis & Save Report",
            command=self._run,
        )
        self.run_btn.pack(side="left", padx=4)

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(
            bar, textvariable=self.status_var, foreground="#444"
        ).pack(side="left", padx=12)

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, **pad)

    # ---------------- file dialogs ----------------
    def _browse_sales(self):
        p = filedialog.askopenfilename(
            title="Select Sales.xlsx",
            filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")],
        )
        if p:
            self.sales_var.set(p)

    def _browse_stock(self):
        p = filedialog.askopenfilename(
            title="Select Stock.xlsx",
            filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")],
        )
        if p:
            self.stock_var.set(p)

    def _browse_out(self):
        p = filedialog.asksaveasfilename(
            title="Save report as",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if p:
            self.out_var.set(p)

    # ---------------- run ----------------
    def _run(self):
        sales = self.sales_var.get().strip()
        stock = self.stock_var.get().strip()
        out = self.out_var.get().strip()

        if not sales or not Path(sales).exists():
            messagebox.showerror("Missing input", "Choose the Sales Excel file.")
            return
        if not stock or not Path(stock).exists():
            messagebox.showerror("Missing input", "Choose the Stock Excel file.")
            return
        if not out:
            out = str(Path(sales).with_name("FEFO_Report.xlsx"))
            self.out_var.set(out)

        try:
            report_date = pd.Timestamp(self.date_var.get().strip()).normalize()
        except Exception:
            messagebox.showerror(
                "Invalid date", "Use YYYY-MM-DD, e.g. 2026-11-09"
            )
            return

        stock_is_current = self.stock_mode_var.get() == "current"

        self.run_btn.configure(state="disabled")
        self.status_var.set("Running analysis...")

        def worker():
            try:
                data = run_analysis(
                    sales, stock, stock_is_current, report_date
                )
                write_excel(data, out)
                self.root.after(0, lambda: self._done(data, out, None))
            except Exception as e:
                self.root.after(0, lambda e=e: self._done(None, None, e))

        threading.Thread(target=worker, daemon=True).start()

    def _done(self, data, out, err):
        self.run_btn.configure(state="normal")
        if err is not None:
            self.status_var.set("Error.")
            messagebox.showerror("Analysis failed", str(err))
            return
        self.data = data
        self.status_var.set(f"Done. Report saved to: {out}")
        self._populate_tabs(data)

    # ---------------- tabs ----------------
    def _populate_tabs(self, data):
        for tab in self.notebook.tabs():
            self.notebook.forget(tab)

        tab_order = [
            ("summary", "Summary"),
            ("products_status", "Products by Status"),
            ("expired", "Expired Stock"),
            ("critical", "Critical (<=30d)"),
            ("warning", "Warning (31-90d)"),
            ("watch", "Watch (91-180d)"),
            ("fefo_violation_products", "FEFO Violation Products"),
            ("fefo_compliant", "FEFO Compliant Products"),
            ("batch_inventory", "Batch Inventory"),
            ("expiry_monitoring", "Expiry Monitoring"),
            ("sales_agg", "Sales Aggregated"),
        ]
        for key, title in tab_order:
            if key in data:
                self._add_tab(self.notebook, title, data[key])

    @staticmethod
    def _fmt(v):
        if v is None:
            return ""
        if isinstance(v, pd.Timestamp):
            return v.strftime("%Y-%m-%d")
        try:
            if pd.isna(v):
                return ""
        except Exception:
            pass
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v)

    @staticmethod
    def _row_tag(row, title):
        # Force colors based on the Tab Title
        if title == "Expired Stock":
            return "expired"
        if title == "Critical (<=30d)":
            return "critical"
        if title == "Warning (31-90d)":
            return "warning"
        if title == "Watch (91-180d)":
            return "watch"
        if title == "FEFO Violation Products":
            return "expired"  # Red
        if title == "FEFO Compliant Products":
            return "ok"       # Green
        # force color bas
        # For mixed tabs (Products by Status, Batch Inventory, etc.), check row data
        status = ""
        for k in ("expiry_status", "status", "fefo_status"):
            if k in row and not pd.isna(row[k]):
                status = str(row[k])
                break
                
        if "Expired" in status: return "expired"
        if "Critical" in status: return "critical"
        if "Warning" in status: return "warning"
        if "Watch" in status: return "watch"
        if "OK" in status: return "ok"
        
        # Conditional colour for FEFO / Overall FEFO violation rates
        metric = str(row.get("Metric", ""))
        if "Rate (%)" in metric:
            try:
                val = float(row.get("Value"))
            except (TypeError, ValueError):
                return "rate_highlight"          # fallback
            if pd.isna(val):
                return "rate_highlight"          # fallback

            if val >= 100:            # exactly 100 % or more  → GREEN
                return "rate_ok"
            elif val >= 95:           # 95 ≤ x < 100           → YELLOW
                return "rate_warning"
            else:                     # x < 95                 → RED
                return "rate_expired"

        return ""

    def _add_tab(self, parent, title, df):
        frame = ttk.Frame(parent)
        parent.add(frame, text=title)

        if df is None or df.empty:
            ttk.Label(frame, text="No rows in this category.",
                      padding=20).pack()
            return

        LIMIT = 1000
        total = len(df)
        display_df = df.head(LIMIT)
        cols = [str(c) for c in df.columns]

        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill="both", expand=True)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        tree = ttk.Treeview(
            tree_frame, columns=cols, show="headings",
            yscrollcommand=vsb.set, xscrollcommand=hsb.set,
        )
        vsb.config(command=tree.yview)
        hsb.config(command=tree.xview)

        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=150, minwidth=60, stretch=False, anchor="w")

        # Colors
        tree.tag_configure("expired", background="#ffcccc") 
        tree.tag_configure("critical", background="#ffdead")  
        tree.tag_configure("warning", background="#fff4c2")    
        tree.tag_configure("watch", background="#E6F700")     
        tree.tag_configure("ok", background="#d9f39e")         
        tree.tag_configure("rate_highlight", background="#ffeb3b", font=("Tahoma", 10, "bold"))  
        tree.tag_configure("rate_ok",      background="#d9f39e", font=("Tahoma", 10, "bold"))
        tree.tag_configure("rate_warning", background="#fff4c2", font=("Tahoma", 10, "bold"))
        tree.tag_configure("rate_expired", background="#fd1515", font=("Tahoma", 10, "bold"))

        for _, row in display_df.iterrows():
            vals = [self._fmt(v) for v in row.values]
            tag = self._row_tag(row, title) 
            tree.insert("", "end", values=vals, tags=(tag,) if tag else ())

        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        note = f"{total} row(s)."
        if total > LIMIT:
            note += f" Showing first {LIMIT} in GUI — full data in Excel."
        ttk.Label(frame, text=note, foreground="#555").pack(
            anchor="w", padx=6, pady=2
        )


def main():
    root = tk.Tk()
    FEFOApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()