from io import BytesIO
from pathlib import Path

import pandas as pd
import pytest

from light import pipeline as p

SAMPLE = Path(__file__).resolve().parent.parent / "sample_data"


# ── Parsers ───────────────────────────────────────────────────────
@pytest.mark.parametrize("raw, days", [
    ("2 Days 5 Hours", 2),
    ("23 Hours 59 Minutes", 1),      # under a day rounds up to 1
    ("10 Minutes", 1),
    ("3 days", 3),
    ("n/a", None),
    (None, None),
])
def test_parse_cycle_time_to_days(raw, days):
    assert p.parse_cycle_time_to_days(raw) == days


@pytest.mark.parametrize("raw, std", [
    ("Due 30 Days", "Credit 30"),
    ("  30 Days After Invoice Date ", "Credit 30"),
    ("CASH", "Cash"),
    ("Letter of credit", "LC"),
    ("Net 120 special", "Net 120 special"),   # unknown text is kept, not lost
])
def test_standardize_payment_term(raw, std):
    assert p.standardize_payment_term(raw) == std


def test_business_days_skip_friday_saturday():
    # Thu 2026-11-05 -> Sun 2026-11-08 spans a Fri/Sat weekend: 1 working day
    assert p._business_days("2026-11-05", "2026-11-08") == 1
    assert p._business_days("2026-11-08", "2026-11-05") == -1


# ── Header auto-detection ─────────────────────────────────────────
def test_load_data_skips_title_rows():
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame({"Entity": ["bld"], "Total amount": [10]}).to_excel(w, index=False, startrow=3)
    df = p.load_data(buf.getvalue(), "report.xlsx")
    assert list(df.columns) == ["Entity", "Total amount"]


# ── Main cleaning pipeline ────────────────────────────────────────
@pytest.fixture
def raw_po():
    return pd.DataFrame({
        "Entity": ["bld", "FAB", "NWH", "BLD"],
        "Purchase id": ["PO-BLD000101", "PO-FAB000102", "PO-NWH000103", "PO-BLD000104"],
        "Purchase requisition": ["PR-1", None, "PR-3", ""],
        "Create Date": ["2026-01-15", "2026-02-01", "2026-03-10", "2026-03-11"],
        "PO status": ["Open order", "Received", "Canceled", "Invoiced"],
        "Buyer Group": ["MT-CV-CM", "EQ-HV-RN", "IT-SW-LC", "UNKNOWN"],
        "Payment terms": ["Due 30 Days", "cash", "Advance", "As agreed"],
        "Currency": ["SAR", "usd", "EUR", "SAR"],
        "Total amount": [1000.0, 100.0, 10.0, None],
        "PO Cycle Time": ["2 Days", "5 Hours", "1 Days", "3 Days 2 Hours"],
    })


def test_clean_dataframe(raw_po):
    out = p.clean_dataframe(raw_po)
    assert len(out) == 3                                     # canceled PO removed
    assert out["Entity"].tolist() == ["BLD", "FAB", "BLD"]   # upper-cased
    assert out["Sector"].tolist() == ["E&C Sector", "Industrial Sector", "E&C Sector"]
    assert out["Total Amount (SAR)"].tolist()[:2] == [1000.0, 375.0]   # USD -> SAR, case-insensitive
    assert pd.isna(out["Total Amount (SAR)"].iloc[2])
    assert out["PO Cycle Time"].tolist() == [2, 1, 3]
    assert out["Main Category"].tolist()[:2] == ["Materials", "Equipment"]
    assert pd.isna(out["Main Category"].iloc[2])             # unknown code stays blank
    assert out["PO Category"].tolist() == ["Standard PO", "Direct PO", "Direct PO"]
    assert out.columns[-1] == "PO Category"
    assert {"Day", "Month", "Year"} <= set(out.columns)


def test_recall_status_latest_action_wins(raw_po):
    main = p.clean_dataframe(raw_po)
    recall = p.clean_recall_df(pd.DataFrame({
        "PO Number": ["po-bld000101", "PO-BLD000101", "PO-FAB000102"],
        "Company": ["bld", "BLD", "FAB"],
        "Action Type": ["Request change", "Cancel", "Request change"],
        "Action Date": ["2026-01-18", "2026-01-20", "2026-02-03"],
    }))
    out = p.add_recall_status(main, recall)
    assert out["Recall Status"].tolist() == ["Cancel", "Request change", "Not recalled"]
    assert recall["PO-Entity Match"].all()


# ── End-to-end on the shipped sample files ────────────────────────
def test_sample_files_end_to_end():
    read = lambda name, marker="entity": p.load_data((SAMPLE / name).read_bytes(), name, header_marker=marker)
    raw = read("po_report_sample.xlsx")
    df = p.clean_dataframe(raw)
    df = p.add_recall_status(df, p.clean_recall_df(read("recall_log_sample.xlsx", "sr#")))
    df = p.add_kpi_metrics(df, p.clean_kpi_df(read("pr_po_kpi_sample.xlsx", "pr number")))

    assert len(df) == len(raw) - (raw["PO status"] == "Canceled").sum()
    assert df["Sector"].notna().all()
    assert df["Main Category"].notna().all()
    assert df["Recall Status"].isin(["Request change", "Cancel", "Not recalled"]).all()
    assert df["PR Approval - PO Submit Time"].dropna().ge(1).all()
