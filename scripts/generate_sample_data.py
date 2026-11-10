"""Generate the synthetic sample files shipped with the repo.

All companies, vendors, people and numbers are fictional. The files mimic the
shape of real ERP exports (title rows above the header, mixed-case entity
codes, messy payment-term text, several currencies, duration strings) so every
step of the cleaning pipeline has something to do.

Usage:  python scripts/generate_sample_data.py
"""
from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
SAMPLE_DIR = ROOT / "sample_data"
DATA_DIR = ROOT / "data"
rng = np.random.default_rng(42)

# ── Reference data ────────────────────────────────────────────────
ENTITIES = {  # code: relative PO volume
    "NWH": 6, "BLD": 18, "INF": 12, "MEP": 10, "FAB": 14, "PKG": 8,
    "DEV": 9, "PRM": 5, "FAC": 8, "FLT": 5, "ITS": 5,
}

CATEGORIES = [  # code, main, segment, subcategory
    ("MT-CV-CM", "Materials", "Civil Materials", "Cement & Concrete"),
    ("MT-CV-ST", "Materials", "Civil Materials", "Reinforcement Steel"),
    ("MT-CV-AG", "Materials", "Civil Materials", "Aggregates & Sand"),
    ("MT-EL-CB", "Materials", "Electrical", "Cables & Wires"),
    ("MT-EL-LT", "Materials", "Electrical", "Lighting Fixtures"),
    ("MT-ME-PP", "Materials", "Mechanical", "Pipes & Fittings"),
    ("MT-ME-HV", "Materials", "Mechanical", "HVAC Equipment"),
    ("MT-FN-PT", "Materials", "Finishing", "Paints & Coatings"),
    ("MT-FN-TL", "Materials", "Finishing", "Tiles & Flooring"),
    ("EQ-HV-RN", "Equipment", "Heavy Equipment", "Equipment Rental"),
    ("EQ-HV-SP", "Equipment", "Heavy Equipment", "Spare Parts"),
    ("EQ-TL-PT", "Equipment", "Tools", "Power Tools"),
    ("SV-PR-CN", "Services", "Professional Services", "Consultancy"),
    ("SV-PR-LG", "Services", "Professional Services", "Legal & Audit"),
    ("SV-FM-CL", "Services", "Facility Management", "Cleaning"),
    ("SV-FM-SC", "Services", "Facility Management", "Security"),
    ("SV-LG-TR", "Services", "Logistics", "Transportation"),
    ("SV-LG-WH", "Services", "Logistics", "Warehousing"),
    ("SV-SB-CT", "Services", "Subcontracting", "Civil Works"),
    ("SV-SB-ME", "Services", "Subcontracting", "MEP Works"),
    ("IT-HW-LP", "IT", "Hardware", "Laptops & Desktops"),
    ("IT-HW-NW", "IT", "Hardware", "Network Equipment"),
    ("IT-SW-LC", "IT", "Software", "Licences & Subscriptions"),
    ("IN-OF-ST", "Indirect", "Office", "Stationery"),
    ("IN-OF-FR", "Indirect", "Office", "Furniture"),
    ("IN-HR-TR", "Indirect", "HR Services", "Training"),
    ("IN-FL-FU", "Indirect", "Fleet", "Fuel"),
    ("IN-FL-MT", "Indirect", "Fleet", "Vehicle Maintenance"),
]

# Which category families each sector tends to buy
SECTOR_PREFS = {
    "NWH": ["SV-PR", "IT-", "IN-"],
    "BLD": ["MT-CV", "MT-FN", "EQ-", "SV-SB"], "INF": ["MT-CV", "EQ-", "SV-SB", "SV-LG"],
    "MEP": ["MT-EL", "MT-ME", "SV-SB-ME", "EQ-TL"],
    "FAB": ["MT-CV-ST", "EQ-", "SV-LG", "MT-ME"], "PKG": ["SV-LG", "EQ-", "IN-"],
    "DEV": ["SV-PR", "SV-SB", "MT-FN"], "PRM": ["SV-FM", "MT-FN", "SV-PR"],
    "FAC": ["SV-FM", "IN-", "MT-EL-LT"], "FLT": ["IN-FL", "EQ-HV-SP", "SV-LG"],
    "ITS": ["IT-", "SV-PR-CN", "IN-HR"],
}

VENDOR_WORDS_A = ["Desert", "Northstar", "Crescent", "Blue Dune", "Golden Palm", "Red Sea",
                  "Summit", "Oasis", "Falcon", "Horizon", "Granite", "Cedar", "Atlas",
                  "Silverline", "Najd", "Coastal", "Pioneer", "Meridian", "Sandstone", "Orbit"]
VENDOR_WORDS_B = {
    "MT": ["Building Materials", "Trading Est.", "Supplies Co.", "Industrial Trading"],
    "EQ": ["Equipment Rental", "Machinery Co.", "Heavy Equipment"],
    "SV": ["Contracting", "Services Co.", "Consulting", "Logistics"],
    "IT": ["Technologies", "Systems", "Digital Solutions"],
    "IN": ["General Trading", "Office Solutions", "Facilities Co."],
}

PAYMENT_TERMS_RAW = [
    "30 days after invoice date", "Due 30 Days", "30 Days After Invoice Date ",
    "60 days after invoice date", "due 60 days", "45 days credit",
    "Cash on delivery", "CASH", "Payment on Delivery", "Advance",
    "50% payment in advance", "Letter of credit", "As agreed", "As Agreed in terms",
    "15 days after invoice date", "90 days after invoice date",
]
PAYMENT_WEIGHTS = np.array([14, 8, 3, 10, 5, 4, 7, 3, 2, 3, 3, 2, 4, 2, 4, 2], float)

PO_TYPES = ["Purchase order", "Service order", "Framework call-off", "Blanket order"]
FIRST = ["Ahmed", "Sara", "Omar", "Lina", "Khalid", "Noura", "Faisal", "Reem", "Yousef", "Huda"]
LAST = ["Al-Harbi", "Al-Qahtani", "Hassan", "Al-Otaibi", "Saleh", "Al-Zahrani", "Nasser", "Al-Shammari"]


def _cycle_text(hours: float) -> str:
    d, rem = divmod(hours, 24)
    h, m = int(rem), int((rem % 1) * 60)
    parts = []
    if d: parts.append(f"{int(d)} Days")
    if h: parts.append(f"{h} Hours")
    if m or not parts: parts.append(f"{m} Minutes")
    return " ".join(parts)


def build_vendors(n=70):
    vendors = []
    used = set()
    for i in range(n):
        cat = CATEGORIES[rng.integers(len(CATEGORIES))]
        fam = cat[0][:2]
        while True:
            name = f"{rng.choice(VENDOR_WORDS_A)} {rng.choice(VENDOR_WORDS_B[fam])}"
            if name not in used:
                used.add(name); break
        vendors.append({
            "Vendor account": f"V{10000 + i}",
            "Vendor name": name,
            "family": fam,
            "VAT number": int("3" + "".join(map(str, rng.integers(0, 10, 13))) + "3"),
            "CR number": int("10" + "".join(map(str, rng.integers(0, 10, 8)))),
        })
    return pd.DataFrame(vendors)


def build_po_report(n=1200):
    vendors = build_vendors()
    ent_codes = list(ENTITIES)
    ent_p = np.array(list(ENTITIES.values()), float); ent_p /= ent_p.sum()
    pay_p = PAYMENT_WEIGHTS / PAYMENT_WEIGHTS.sum()
    start, end = pd.Timestamp("2025-01-05"), pd.Timestamp("2026-10-30")
    span = (end - start).days

    counters = {e: 100 for e in ent_codes}
    rows = []
    for _ in range(n):
        ent = rng.choice(ent_codes, p=ent_p)
        prefs = SECTOR_PREFS[ent]
        cats = [c for c in CATEGORIES if any(c[0].startswith(p) for p in prefs)]
        cat = cats[rng.integers(len(cats))]
        fam = cat[0][:2]
        vpool = vendors[vendors["family"] == fam]
        v = vpool.iloc[rng.integers(len(vpool))]

        # spend: lognormal, bigger for subcontracting / equipment
        scale = {"SV": 11.2, "EQ": 11.0, "MT": 10.6, "IT": 10.2, "IN": 9.2}[fam]
        amount_sar = float(np.exp(rng.normal(scale, 1.0)))
        currency = rng.choice(["SAR", "USD", "EUR", "AED", "usd"], p=[0.82, 0.09, 0.04, 0.03, 0.02])
        fx = {"SAR": 1, "USD": 3.75, "usd": 3.75, "EUR": 4.08, "AED": 1.02}[currency]
        amount = round(amount_sar / fx, 2)

        created = start + pd.Timedelta(days=int(rng.integers(span)), hours=int(rng.integers(7, 17)),
                                       minutes=int(rng.integers(60)))
        age_days = (end - created).days
        # older POs are more likely to be closed
        p_open = 0.65 if age_days < 60 else 0.30 if age_days < 240 else 0.10
        status = rng.choice(["Open order", "Received", "Invoiced", "Canceled"],
                            p=[p_open, (1 - p_open) * 0.35, (1 - p_open) * 0.57, (1 - p_open) * 0.08])
        approval = ("Confirmed" if status in ("Received", "Invoiced") else
                    rng.choice(["Confirmed", "Approved", "In review", "Draft", "Pending", "Rejected"],
                               p=[0.55, 0.15, 0.10, 0.08, 0.08, 0.04]))
        counters[ent] += int(rng.integers(1, 4))
        has_pr = rng.random() > 0.22
        cycle_h = float(rng.gamma(2.2, 30)) if has_pr else float(rng.gamma(1.4, 10))

        rows.append({
            "Entity": ent.lower() if rng.random() < 0.15 else ent,
            "Purchase id": f"PO-{ent}{counters[ent]:06d}",
            "Purchase requisition": f"PR-{ent}{counters[ent] + 5000:06d}" if has_pr else None,
            "Create Date": created,
            "PO type": rng.choice(PO_TYPES, p=[0.62, 0.25, 0.08, 0.05]),
            "PO status": status,
            "Approval status": approval,
            "Vendor account": v["Vendor account"],
            "Vendor name": v["Vendor name"],
            "VAT number": v["VAT number"],
            "CR number": v["CR number"],
            "Buyer Group": cat[0],
            "Buyer Group Name": cat[3],
            "Payment terms": rng.choice(PAYMENT_TERMS_RAW, p=pay_p),
            "Currency": currency,
            "Total amount": amount,
            "PO Cycle Time": _cycle_text(cycle_h),
            "Bid exemption": int(rng.random() < (0.35 if amount_sar < 20000 else 0.08)),
        })
    df = pd.DataFrame(rows).sort_values("Create Date").reset_index(drop=True)
    return df


def build_recall_log(po: pd.DataFrame, n=140):
    sample = po.sample(n, random_state=7)
    users = [(f"U{100 + i}", f"{rng.choice(FIRST)} {rng.choice(LAST)}") for i in range(8)]
    rows = []
    for k, (_, r) in enumerate(sample.iterrows(), start=1):
        uid, uname = users[rng.integers(len(users))]
        if rng.random() < 0.1:  # same user, inconsistent spelling -> pipeline normalises it
            uname = uname.split()[0] + " A. " + uname.split()[-1]
        rows.append({
            "Sr#": k,
            "PO Number": r["Purchase id"],
            "Company": str(r["Entity"]).upper(),
            "User Id": uid,
            "User Name": uname,
            "Action Type": rng.choice(["Request change", "Cancel"], p=[0.8, 0.2]),
            "Operation": rng.choice(["Line quantity", "Unit price", "Delivery date", "Vendor change", "Header"]),
            "Action Date": r["Create Date"] + pd.Timedelta(days=int(rng.integers(0, 25)), hours=int(rng.integers(0, 8))),
        })
    return pd.DataFrame(rows)


def build_kpi_report(po: pd.DataFrame):
    with_pr = po[po["Purchase requisition"].notna()].sample(frac=0.85, random_state=3)
    rows = []
    for _, r in with_pr.iterrows():
        po_created = r["Create Date"]
        approval = po_created - pd.Timedelta(days=int(rng.gamma(2.0, 3)), hours=int(rng.integers(0, 8)))
        rows.append({
            "PR number": r["Purchase requisition"],
            "Purchase order": r["Purchase id"],
            "PR Assignment Date": approval - pd.Timedelta(days=int(rng.integers(0, 4))),
            "PR Approval Date": approval,
            "PO Creation Date": po_created,
            "PO Submission Date": po_created + pd.Timedelta(days=int(rng.gamma(1.5, 2)), hours=int(rng.integers(1, 9))),
            # raw ERP durations (calendar days); the pipeline recomputes them in working days
            "PR Approval - PO Creation  Time": None,
            "PO Creation  - PO Submit Time": None,
            "PR Approval - PO Submit Time": None,
        })
    return pd.DataFrame(rows)


def write_with_title(df: pd.DataFrame, path: Path, title: str, sheet: str = "Report"):
    """Write like an ERP export: 3 metadata rows, a blank row, then the table."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=sheet, startrow=4)
    wb = load_workbook(path)
    ws = wb[sheet]
    ws["A1"] = title
    ws["A2"] = "Company: Northwind Holding (fictional sample data)"
    ws["A3"] = f"Generated by: scripts/generate_sample_data.py"
    for col in ws.columns:
        width = max(len(str(c.value)) if c.value is not None else 0 for c in col[4:60])
        ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 10), 40)
    wb.save(path)


def write_category_file(path: Path):
    df = pd.DataFrame(CATEGORIES, columns=["Code", "Main Category", "Segment", "Subcategory"])
    df.insert(0, "Created Date", pd.Timestamp("2026-01-01"))
    df = df[["Created Date", "Main Category", "Segment", "Subcategory", "Code"]]
    df["Desc."] = df["Main Category"] + " / " + df["Segment"] + " / " + df["Subcategory"]
    write_with_title(df, path, "Standard Vendor Categorization (sample)", sheet="Categories")


def main():
    po = build_po_report()
    write_with_title(po, SAMPLE_DIR / "po_report_sample.xlsx", "Purchase Order Report")
    write_with_title(build_recall_log(po), SAMPLE_DIR / "recall_log_sample.xlsx", "PO Change History")
    write_with_title(build_kpi_report(po), SAMPLE_DIR / "pr_po_kpi_sample.xlsx", "PR to PO KPI Report")
    write_category_file(DATA_DIR / "vendor_categorization_sample.xlsx")
    print(f"PO rows: {len(po):,}  ·  open: {(po['PO status'] == 'Open order').sum():,}  ·  "
          f"canceled: {(po['PO status'] == 'Canceled').sum():,}")


if __name__ == "__main__":
    main()
