"""Business configuration for Light.

Everything an organisation needs to adapt lives here: which legal entity belongs
to which sector, FX rates, how raw payment-term text is standardised, and where
the vendor category master file sits.

The values below are **illustrative sample values** for a fictional holding
company ("Northwind Holding"). Replace them with your own before using the app
on real data.
"""
from __future__ import annotations

from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parent.parent

# ── Entity → Sector ──────────────────────────────────────────────
# Keys are the entity codes as they appear in the ERP "Entity" column
# (upper-cased by the pipeline). Unknown entities get an empty Sector.
SECTOR_MAP: dict[str, str] = {
    "NWH": "Corporate",
    "BLD": "E&C Sector",
    "INF": "E&C Sector",
    "MEP": "E&C Sector",
    "FAB": "Industrial Sector",
    "PKG": "Industrial Sector",
    "DEV": "Real Estate Sector",
    "PRM": "Real Estate Sector",
    "FAC": "Supporting Services Sector",
    "FLT": "Supporting Services Sector",
    "ITS": "Supporting Services Sector",
}

# ── FX rates to SAR ──────────────────────────────────────────────
# Static reference rates used to normalise spend into one currency.
# Update periodically, or swap for a rates table/API in production.
FX_TO_SAR: dict[str, float] = {
    "SAR": 1.00, "USD": 3.75, "AED": 1.02, "EUR": 4.08, "GBP": 4.76,
    "BHD": 9.95, "KWD": 12.19, "OMR": 9.74, "QAR": 1.03, "EGP": 0.075,
}

# ── Payment terms ────────────────────────────────────────────────
# Raw ERP text (matched case-insensitively, trimmed) → standard category.
# Anything not listed is kept as-is so new variants stay visible.
PAYMENT_TERM_MAP: dict[str, str] = {
    "cash": "Cash",
    "cash on delivery": "Cash",
    "payment on delivery": "Cash",
    "cheque after delivery": "Cash",
    "advance": "Advance",
    "in advance": "Advance",
    "50% payment in advance": "Advance 50%",
    "15 days after invoice date": "Credit 15",
    "due 15 days": "Credit 15",
    "30 days after invoice date": "Credit 30",
    "due 30 days": "Credit 30",
    "45 days after invoice date": "Credit 45",
    "45 days credit": "Credit 45",
    "60 days after invoice date": "Credit 60",
    "due 60 days": "Credit 60",
    "90 days after invoice date": "Credit 90",
    "30 days lc": "LC",
    "letter of credit": "LC",
    "cheque": "Cheque",
    "credit": "Credit",
    "as agreed": "As Agreed",
    "as agreed in terms": "As Agreed",
    "as agreed below": "As Agreed",
}

# ── Vendor category master ───────────────────────────────────────
# Excel file mapping each Buyer Group code to Main Category / Segment /
# Subcategory. The header row is auto-detected (looks for "Main Category").
CATEGORY_FILE = ROOT_DIR / "data" / "vendor_categorization_sample.xlsx"
CATEGORY_SHEET = "Categories"

# ── Calendar ─────────────────────────────────────────────────────
# Saudi work week: Sunday–Thursday working, Friday & Saturday weekend.
# numpy weekmask order is Mon..Sun.
WORK_WEEKMASK = "1111001"

# ── Branding ─────────────────────────────────────────────────────
LOGO_PATH = ROOT_DIR / "assets" / "logo.png"
