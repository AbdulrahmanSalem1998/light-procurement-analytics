"""Data-cleaning pipeline for ERP purchase-order exports.

Pure pandas functions with no Streamlit dependency, so they can be unit-tested
and reused from notebooks or scheduled jobs. The Streamlit UI (``app.py``)
wraps the I/O-heavy ones with ``st.cache_data``.
"""
from __future__ import annotations

import math
import re
from functools import lru_cache
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd

from .config import (
    CATEGORY_FILE,
    CATEGORY_SHEET,
    FX_TO_SAR,
    PAYMENT_TERM_MAP,
    SECTOR_MAP,
    WORK_WEEKMASK,
)

# ── Data loaders ───────────────────────────────────────────────
def load_data(file_bytes, file_name, header_marker="entity"):
    """header_marker: the lowercase text of a header cell used to auto-detect which
    row is the true header row (skipping title/metadata rows above it).
    Main PO report -> 'entity'. Recall / Change History report -> 'sr#'."""
    if file_name.endswith(".csv"):
        return pd.read_csv(BytesIO(file_bytes))
    raw = pd.read_excel(BytesIO(file_bytes), header=None, nrows=20)
    header_row = 0
    for i, row in raw.iterrows():
        if any(str(v).strip().lower() == header_marker for v in row.dropna()):
            header_row = i
            break
    return pd.read_excel(BytesIO(file_bytes), skiprows=header_row)

# ── Cycle time parsers ─────────────────────────────────────────
def _parse_hours(val):
    if pd.isna(val): return None
    v = str(val)
    d = re.search(r"(\d+)\s*Day",    v, re.I)
    h = re.search(r"(\d+)\s*Hour",   v, re.I)
    m = re.search(r"(\d+)\s*Minute", v, re.I)
    if not any([d, h, m]): return None
    total = 0
    if d: total += int(d.group(1)) * 24
    if h: total += int(h.group(1))
    if m: total += int(m.group(1)) / 60
    return total

def parse_cycle_time_to_days(val):
    hours = _parse_hours(val)
    if hours is None: return None
    return max(1, math.floor(hours / 24))

# ── Payment terms ──────────────────────────────────────────────
def standardize_payment_term(raw):
    if pd.isna(raw):
        return None
    raw_stripped = str(raw).strip()
    return PAYMENT_TERM_MAP.get(raw_stripped.lower(), raw_stripped)


# ── Vendor category master ─────────────────────────────────────
@lru_cache(maxsize=4)
def load_category_maps(path: str | Path = CATEGORY_FILE, sheet: str = CATEGORY_SHEET):
    """Return three dicts keyed by Buyer Group code: main category, segment, subcategory.
    Missing or unreadable file -> empty dicts (the pipeline still runs)."""
    code_to_main, code_to_seg, code_to_sub = {}, {}, {}
    try:
        raw = pd.read_excel(path, sheet_name=sheet, header=None)
        hr = next((i for i, row in raw.iterrows()
                   if any(str(v).strip().lower() == "main category" for v in row.dropna())), 0)
        df = pd.read_excel(path, sheet_name=sheet, skiprows=hr)
        df.columns = df.columns.str.strip()
        df = df.dropna(subset=["Code"])
        seg_col = next((c for c in df.columns if "segment" in c.lower()), None)
        sub_col = next((c for c in df.columns if c.lower() == "subcategory"), None)
        for _, r in df.iterrows():
            code = str(r["Code"]).strip()
            code_to_main[code] = str(r["Main Category"]).strip()
            if seg_col:
                code_to_seg[code] = str(r[seg_col]).strip()
            if sub_col:
                code_to_sub[code] = str(r[sub_col]).strip()
    except (FileNotFoundError, ValueError, KeyError):
        pass
    return code_to_main, code_to_seg, code_to_sub


def _lookup(mapping, code):
    return mapping.get(str(code).strip()) if pd.notna(code) else None


# ── Column auto-detection helpers ────────────────────────────────
PR_COL_NAMES = [
    "pr number", "pr no", "pr no.", "pr", "pr#", "pr #",
    "requisition number", "requisition no", "requisition no.",
]
PO_NUMBER_COL_NAMES = [
    "po number", "po no", "po no.", "po#", "po #", "purchase order number",
    "purchase order no", "po id", "poid", "purchase id",
]
CONFIRMATION_DATE_COL_NAMES = ["confirmation date", "confirmation dt"]
CREATE_DATE_COL_NAMES = ["create date", "created date", "po create date", "creation date"]

def find_column(df, exact_names, contains=None):
    """Find a column by exact (case-insensitive) name match, falling back to a substring match."""
    for c in df.columns:
        if str(c).strip().lower() in exact_names:
            return c
    if contains:
        for c in df.columns:
            if contains in str(c).strip().lower():
                return c
    return None

# ── Clean pipeline ─────────────────────────────────────────────
def clean_dataframe(df):
    df = df.copy()
    df.columns = df.columns.str.strip()
    df = df.dropna(axis=1, how="all")
    df = df.loc[:, ~df.columns.str.startswith("Column")]

    if "Entity" in df.columns:
        df["Entity"] = df["Entity"].str.upper().str.strip()
        df.insert(df.columns.get_loc("Entity") + 1, "Sector", df["Entity"].map(SECTOR_MAP))


    # Fix long numeric ID columns — convert to string to prevent scientific notation

    for col in ["VAT number", "CR number"]:
        if col in df.columns:
            def safe_to_str(x):
                if pd.isna(x): return x
                s = str(x).strip()
                if s in ["", "nan"]: return x
                try:
                    return str(int(float(s)))
                except Exception:
                    return s
            df[col] = df[col].apply(safe_to_str)



    # Step 4: Remove canceled POs

    if "PO status" in df.columns:
        df = df[df["PO status"].str.strip() != "Canceled"].copy()

    if "Total amount" in df.columns and "Currency" in df.columns:
        df["Total Amount (SAR)"] = df.apply(
            lambda r: round(r["Total amount"] * FX_TO_SAR.get(str(r["Currency"]).strip().upper(), 1), 2)
            if pd.notna(r["Total amount"]) else None, axis=1
        )

    if "PO Cycle Time" in df.columns:
        df["PO Cycle Time"] = df["PO Cycle Time"].apply(parse_cycle_time_to_days)

    if "Buyer Group" in df.columns:
        idx = df.columns.get_loc("Buyer Group")
        code_to_main, code_to_seg, code_to_sub = load_category_maps()
        df.insert(idx + 1, "Main Category", df["Buyer Group"].apply(lambda c: _lookup(code_to_main, c)))
        df.insert(idx + 2, "Segment",       df["Buyer Group"].apply(lambda c: _lookup(code_to_seg, c)))
        df.insert(idx + 3, "Subcategory",   df["Buyer Group"].apply(lambda c: _lookup(code_to_sub, c)))

    for col in df.columns:
        if col.strip().lower() == "buyer group name":
            df.rename(columns={col: "Buyer Group Desc"}, inplace=True)
            break

    # Standardized Payment Term: maps messy raw text to a fixed set of categories
    pay_col = find_column(df, ["payment terms", "payment term"], contains="payment term")
    if pay_col:
        idx = df.columns.get_loc(pay_col)
        df.insert(idx + 1, "Standardized Payment Term", df[pay_col].apply(standardize_payment_term))

    # PO Category: Standard PO if a PR reference exists, otherwise Direct PO
    pr_col = find_column(df, PR_COL_NAMES, contains="requisition")
    if pr_col:
        def _po_category(x):
            if pd.isna(x): return "Direct PO"
            s = str(x).strip().lower()
            return "Direct PO" if s in ["", "nan", "none", "0"] else "Standard PO"
        df["PO Category"] = df[pr_col].apply(_po_category)

    if "Create Date" in df.columns:
        df["Create Date"] = pd.to_datetime(df["Create Date"], errors="coerce")
        idx = df.columns.get_loc("Create Date")
        df.insert(idx + 1, "Day",   df["Create Date"].dt.day.astype("Int64"))
        df.insert(idx + 2, "Month", df["Create Date"].dt.strftime("%b"))
        df.insert(idx + 3, "Year",  df["Create Date"].dt.year.astype("Int64"))

    # Push PO Category to the very last column, after Total Amount (SAR)
    if "PO Category" in df.columns:
        other_cols = [c for c in df.columns if c != "PO Category"]
        df = df[other_cols + ["PO Category"]]

    return df

# ── Recall / PO Change History cleaning pipeline ─────────────────
def clean_recall_df(df):
    df = df.copy()
    df.columns = df.columns.str.strip()
    df = df.dropna(axis=1, how="all")
    df = df.loc[:, ~df.columns.str.startswith("Unnamed")]

    if "Company" in df.columns:
        df["Company"] = df["Company"].astype(str).str.upper().str.strip()
        df.insert(df.columns.get_loc("Company") + 1, "Sector", df["Company"].map(SECTOR_MAP))

    # Standardize User Name per User Id (e.g. "John Smith" vs "John A. Smith")
    if "User Id" in df.columns and "User Name" in df.columns:
        canonical = df.groupby("User Id")["User Name"].agg(lambda s: s.value_counts().idxmax())
        df["User Name"] = df["User Id"].map(canonical).fillna(df["User Name"])

    if "Action Type" in df.columns:
        df["Action Type"] = df["Action Type"].astype(str).str.strip()
    if "Operation" in df.columns:
        df["Operation"] = df["Operation"].astype(str).str.strip()

    # Cross-check PO Number prefix against Company (e.g. PO-BLD000438 -> BLD)
    if "PO Number" in df.columns and "Company" in df.columns:
        def _po_entity_match(row):
            m = re.match(r"PO-([A-Za-z]+)", str(row["PO Number"]))
            if not m:
                return None
            return m.group(1).upper() == str(row["Company"]).strip().upper()
        df["PO-Entity Match"] = df.apply(_po_entity_match, axis=1)

    if "Action Date" in df.columns:
        df["Action Date"] = pd.to_datetime(df["Action Date"], errors="coerce")

    return df


def _business_days(start, end, weekmask=WORK_WEEKMASK):
    """Whole business days between start and end (Fri/Sat excluded as weekend).
    Positive when end is after start, negative when end is before start."""
    if pd.isna(start) or pd.isna(end):
        return None
    s = pd.Timestamp(start).date()
    e = pd.Timestamp(end).date()
    if e >= s:
        return int(np.busday_count(s, e, weekmask=weekmask))
    return -int(np.busday_count(e, s, weekmask=weekmask))

def add_recall_status(main_df, recall_df):
    """Adds 'Recall Status', 'Action Date', and 'Days to Action Recall' columns to
    main_df, based on the most recent action per PO Number in the cleaned recall log.
    - Recall Status:         Request change / Cancel / Not recalled
    - Action Date:           date/time of that latest action (blank if not recalled)
    - Days to Action Recall: business days from Create Date to Action Date, treating
                       Friday and Saturday as weekend (blank if not recalled or Create
                       Date is missing). Positive = action happened after the PO was
                       created; negative = action happened before creation (shouldn't
                       normally occur, but flagged as-is if the source data shows it).
    """
    main_df = main_df.copy()

    po_col_main = find_column(main_df, PO_NUMBER_COL_NAMES)
    po_col_recall = "PO Number" if recall_df is not None and "PO Number" in recall_df.columns else None

    if not po_col_main or recall_df is None or not po_col_recall:
        main_df["Recall Status"]         = "Not recalled"
        main_df["Action Date"]           = pd.NaT
        main_df["Days to Action Recall"] = None
        return main_df

    recall_sorted = recall_df.sort_values("Action Date") if "Action Date" in recall_df.columns else recall_df
    recall_sorted = recall_sorted.copy()
    recall_sorted["_po_key"] = recall_sorted[po_col_recall].astype(str).str.strip().str.upper()
    latest = recall_sorted.groupby("_po_key").last()  # most recent row per normalized PO key

    latest_action_type = latest["Action Type"] if "Action Type" in latest.columns else pd.Series(dtype=object)
    latest_action_date = latest["Action Date"] if "Action Date" in latest.columns else pd.Series(dtype="datetime64[ns]")

    def _map_status(at):
        if at is None or pd.isna(at):
            return "Not recalled"
        at = str(at).strip().lower()
        if at == "cancel":
            return "Cancel"
        if at == "request change":
            return "Request change"
        return "Not recalled"

    main_po_key = main_df[po_col_main].astype(str).str.strip().str.upper()
    mapped_action_type = main_po_key.map(latest_action_type)
    mapped_action_date = main_po_key.map(latest_action_date)

    main_df["Recall Status"] = mapped_action_type.apply(_map_status)
    main_df["Action Date"]   = mapped_action_date

    # Days to Action Recall: Create Date -> Action Date, in BUSINESS days
    # (Saudi work week: Sun-Thu working, Fri & Sat = weekend), only for recalled POs
    create_col = find_column(main_df, CREATE_DATE_COL_NAMES, contains="create date")
    if create_col:
        create_dates = pd.to_datetime(main_df[create_col], errors="coerce")
        main_df["Days to Action Recall"] = [
            _business_days(s, e) for s, e in zip(create_dates, main_df["Action Date"])
        ]
    else:
        main_df["Days to Action Recall"] = None

    return main_df

# ── PR-to-PO KPI cleaning pipeline ────────────────────────────────
KPI_DURATION_COLS = [
    "PR Approval - PO Creation  Time",
    "PO Creation  - PO Submit Time",
    "PR Approval - PO Submit Time",
]
KPI_DATE_COLS = ["PR Assignment Date", "PR Approval Date", "PO Creation Date", "PO Submission Date"]


# (start column, end column) that each duration column is measured between
KPI_DURATION_DATE_PAIRS = {
    "PR Approval - PO Creation  Time": ("PR Approval Date", "PO Creation Date"),
    "PO Creation  - PO Submit Time":   ("PO Creation Date", "PO Submission Date"),
    "PR Approval - PO Submit Time":    ("PR Approval Date", "PO Submission Date"),
}

def _weekdays_between(start, end):
    """Whole weekdays elapsed between start and end, using the Sun-Thu work week
    (Friday/Saturday are weekend and excluded). If the gap is under a full day
    (or falls entirely within a weekend), returns 1."""
    if pd.isna(start) or pd.isna(end):
        return None
    start_d = pd.Timestamp(start).normalize()
    end_d = pd.Timestamp(end).normalize()
    if end_d < start_d:
        return None
    days = int(np.busday_count(start_d.date(), end_d.date(), weekmask=WORK_WEEKMASK))
    return days if days >= 1 else 1

def clean_kpi_df(df):
    df = df.copy()
    df.columns = df.columns.str.strip()
    df = df.dropna(axis=1, how="all")
    df = df.loc[:, ~df.columns.str.startswith("Unnamed")]

    for col in KPI_DATE_COLS:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Recompute the 3 duration columns as whole working days (Sun–Thu), from the real date pairs
    for dur_col, (start_col, end_col) in KPI_DURATION_DATE_PAIRS.items():
        if start_col in df.columns and end_col in df.columns:
            df[dur_col] = [
                _weekdays_between(s, e) for s, e in zip(df[start_col], df[end_col])
            ]

    return df

def add_kpi_metrics(main_df, kpi_df):
    """Adds the 3 PR-to-PO cycle-time columns (whole working days, Sun–Thu) to main_df.
    Matches first by PR number (Purchase requisition), falling back to PO number
    (Purchase id) when no PR match is found. When the KPI file has duplicate PR or
    PO numbers, the first occurrence in the file wins."""
    main_df = main_df.copy()

    if kpi_df is None:
        for col in KPI_DURATION_COLS:
            main_df[col] = None
        return main_df

    pr_col_main = find_column(main_df, PR_COL_NAMES, contains="requisition")
    po_col_main = find_column(main_df, PO_NUMBER_COL_NAMES)

    kpi = kpi_df.copy()
    has_pr = "PR number" in kpi.columns
    has_po = "Purchase order" in kpi.columns
    if has_pr:
        kpi["_pr_key"] = kpi["PR number"].astype(str).str.strip().str.upper()
    if has_po:
        kpi["_po_key"] = kpi["Purchase order"].astype(str).str.strip().str.upper()

    pr_lookup = (kpi.dropna(subset=["_pr_key"]).drop_duplicates(subset="_pr_key", keep="first").set_index("_pr_key")
                 if has_pr else None)
    po_lookup = (kpi.dropna(subset=["_po_key"]).drop_duplicates(subset="_po_key", keep="first").set_index("_po_key")
                 if has_po else None)

    empty_series = pd.Series([None] * len(main_df), index=main_df.index)
    main_pr_key = main_df[pr_col_main].astype(str).str.strip().str.upper() if pr_col_main else empty_series
    main_po_key = main_df[po_col_main].astype(str).str.strip().str.upper() if po_col_main else empty_series

    for col in KPI_DURATION_COLS:
        if col not in kpi.columns:
            main_df[col] = None
            continue
        pr_vals = main_pr_key.map(pr_lookup[col]) if pr_lookup is not None else empty_series
        po_vals = main_po_key.map(po_lookup[col]) if po_lookup is not None else empty_series
        main_df[col] = pr_vals.combine_first(po_vals)  # PR match wins, PO match fills the rest

    return main_df
