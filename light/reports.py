"""Excel report builders (openpyxl).

* ``to_excel_bytes``           – cleaned dataset as a formatted Excel table
* ``to_open_po_excel``         – Open PO tracker: one sheet per entity, with
                                 dropdowns for reason / expected closure date
* ``to_summary_report_excel``  – executive summary workbook: KPI cards, sector
                                 and status breakdowns, Entity × Year sheets
"""
from __future__ import annotations

from io import BytesIO

import pandas as pd

from .config import LOGO_PATH


def closure_month_options(n_months: int = 18, start=None) -> list[str]:
    """Rolling list of "Mon YYYY" labels for the Expected Closure Date dropdown,
    starting from the current month. Kept short because Excel caps list-type
    validation formulas at 255 characters."""
    start = pd.Timestamp.today() if start is None else pd.Timestamp(start)
    first = start.to_period("M")
    return [(first + i).strftime("%b %Y") for i in range(n_months)]

def to_excel_bytes(df):
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter

    date_cols = [c for c in df.columns if pd.api.types.is_datetime64_any_dtype(df[c])]

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Cleaned Data")
        ws = w.sheets["Cleaned Data"]

        header_fill  = PatternFill(start_color="243322", end_color="243322", fill_type="solid")
        header_font  = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
        header_align = Alignment(horizontal="left", vertical="center")

        for cell in ws[1]:
            cell.fill      = header_fill
            cell.font      = header_font
            cell.alignment = header_align

        data_align = Alignment(horizontal="left", vertical="center")
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = data_align

        for col_idx, col_cells in enumerate(ws.columns, start=1):
            max_len = 0
            for cell in col_cells:
                try:
                    cell_len = len(str(cell.value)) if cell.value is not None else 0
                    if cell_len > max_len:
                        max_len = cell_len
                except Exception:
                    pass
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 50)



        # Force text format for long ID columns

        text_cols = ["VAT number", "CR number"]
        for idx, col_cells in enumerate(ws.columns, start=1):
            header = ws.cell(row=1, column=idx).value
            if header in text_cols:
                col_letter = get_column_letter(idx)
                for row_num in range(2, ws.max_row + 1):
                    cell = ws[f"{col_letter}{row_num}"]
                    cell.number_format = "@"
                    if cell.value is not None:
                        try:
                            cell.value = str(int(float(str(cell.value))))
                        except Exception:
                            cell.value = str(cell.value)


        # Date columns: show as M/D/YYYY with no time (e.g. 5/15/2024)
        for idx, col_cells in enumerate(ws.columns, start=1):
            header = ws.cell(row=1, column=idx).value
            if header in date_cols:
                col_letter = get_column_letter(idx)
                for row_num in range(2, ws.max_row + 1):
                    ws[f"{col_letter}{row_num}"].number_format = "m/d/yyyy"


        # Excel AutoFilter (dropdown filters on header row)

        last_col = get_column_letter(ws.max_column)
        last_row = ws.max_row
        ws.auto_filter.ref = f"A1:{last_col}{last_row}"

    return out.getvalue()

# ── Open PO Excel builder ──────────────────────────────────────
def to_open_po_excel(df, logo_path=None):
    from openpyxl.styles import PatternFill, Font, Alignment, Border
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.drawing.image import Image as XLImage

    OPEN_REASONS = [
        "Pending Delivery",
        "Pending Invoice / Payment Processing",
        "Pending Documentation / Approval",
        "Partial Receipt / Quantity Adjustment",
        "Pending PO Closure",
        "Open PO Linked to Project Duration",
        "Under Internal Review",
        "Other",
    ]

    if "PO status" in df.columns:
        open_df = df[df["PO status"].str.strip() == "Open order"].copy()
    else:
        open_df = df.copy()

    if "Open PO Reason" not in open_df.columns:
        open_df["Open PO Reason"] = ""
    if "Expected Closure Date" not in open_df.columns:
        open_df["Expected Closure Date"] = ""
    if "Comment" not in open_df.columns:
        open_df["Comment"] = ""

    hdr_fill  = PatternFill(start_color="243322", end_color="243322", fill_type="solid")
    hdr_font  = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
    hdr_align = Alignment(horizontal="left", vertical="center")
    data_align = Alignment(horizontal="left", vertical="center")
    no_border = Border()

    out = BytesIO()

    if logo_path is None and LOGO_PATH.exists():
        logo_path = str(LOGO_PATH)

    with pd.ExcelWriter(out, engine="openpyxl") as w:
        entities = sorted(open_df["Entity"].dropna().unique().tolist()) if "Entity" in open_df.columns else ["All"]

        for entity in entities:
            if "Entity" in open_df.columns:
                sheet_df = open_df[open_df["Entity"] == entity].copy()
            else:
                sheet_df = open_df.copy()

            if sheet_df.empty:
                continue

            sheet_name = str(entity)[:31]
            sheet_df.to_excel(w, index=False, sheet_name=sheet_name, startrow=5)
            ws = w.sheets[sheet_name]

            total_cols = len(sheet_df.columns)
            last_merge_col = get_column_letter(max(total_cols, 10))

            ws.merge_cells(f"A1:{last_merge_col}1")
            ws.merge_cells(f"C2:{last_merge_col}2")
            ws.merge_cells(f"C3:{last_merge_col}3")
            ws.merge_cells(f"A4:{last_merge_col}4")

            sec_colors_hdr = {
                "e&c sector":                 "1A4449",
                "industrial sector":          "DC8436",
                "supporting services sector": "8E1717",
                "real estate sector":         "00303F",
                "corporate":                  "243322",
            }
            if "Sector" in sheet_df.columns and len(sheet_df) > 0:
                sec_key = str(sheet_df["Sector"].iloc[0]).strip().lower()
                hdr_color = sec_colors_hdr.get(sec_key, "243322")
            else:
                hdr_color = "243322"
            banner_fill = PatternFill(start_color=hdr_color, end_color=hdr_color, fill_type="solid")
            for row_num in range(1, 6):
                for col_num in range(1, max(total_cols, 10) + 1):
                    ws.cell(row=row_num, column=col_num).fill = banner_fill

            title_cell = ws["C2"]
            title_cell.value = "Open PO Tracker"
            title_cell.font  = Font(bold=True, size=20, color="FFFFFF", name="Calibri")
            title_cell.alignment = Alignment(horizontal="center", vertical="center")

            sub_cell = ws["C3"]
            sub_cell.value = f"Entity: {entity}   ·   Generated: {pd.Timestamp.today().strftime('%d %b %Y')}"
            sub_cell.font  = Font(size=10, color="FFFFFF", bold=True, name="Calibri")
            sub_cell.alignment = Alignment(horizontal="center", vertical="center")

            ws.row_dimensions[1].height = 8
            ws.row_dimensions[2].height = 42
            ws.row_dimensions[3].height = 22
            ws.row_dimensions[4].height = 8
            ws.row_dimensions[5].height = 6

            if logo_path:
                try:
                    img = XLImage(logo_path)
                    img.height = 74
                    img.width  = 180
                    img.anchor = "A1"
                    ws.add_image(img)
                except Exception:
                    pass

            header_row = 6
            for cell in ws[header_row]:
                cell.fill      = hdr_fill
                cell.font      = hdr_font
                cell.alignment = hdr_align

            reason_col_idx = None
            date_col_idx   = None
            for idx, col in enumerate(sheet_df.columns, start=1):
                if col == "Open PO Reason":
                    reason_col_idx = idx
                if col == "Expected Closure Date":
                    date_col_idx = idx

            last_data_row = header_row + len(sheet_df)

            for row in ws.iter_rows(min_row=header_row+1, max_row=last_data_row):
                for cell in row:
                    cell.alignment = data_align

            if reason_col_idx:
                reason_formula = '"' + ",".join(OPEN_REASONS) + '"'
                dv_reason = DataValidation(
                    type="list",
                    formula1=reason_formula,
                    allow_blank=True,
                    showDropDown=False,
                    showErrorMessage=True,
                    errorTitle="Invalid",
                    error="Please select a reason from the list.",
                )
                ws.add_data_validation(dv_reason)
                reason_letter = get_column_letter(reason_col_idx)
                dv_reason.sqref = f"{reason_letter}{header_row+1}:{reason_letter}{last_data_row+500}"

            if date_col_idx:
                month_list = closure_month_options()
                month_formula = '"' + ",".join(month_list) + '"'
                dv_date = DataValidation(
                    type="list",
                    formula1=month_formula,
                    allow_blank=True,
                    showDropDown=False,
                    showErrorMessage=True,
                    errorTitle="Invalid",
                    error="Please select a month from the list.",
                )
                ws.add_data_validation(dv_date)
                date_letter = get_column_letter(date_col_idx)
                dv_date.sqref = f"{date_letter}{header_row+1}:{date_letter}{last_data_row+500}"
                no_fill = PatternFill(fill_type=None)
                for row_num in range(header_row + 1, last_data_row + 500):
                    ws[f"{date_letter}{row_num}"].fill = no_fill

            text_cols = ["VAT number", "CR number"]
            for idx, col in enumerate(sheet_df.columns, start=1):
                if col in text_cols:
                    col_letter = get_column_letter(idx)
                    for row_num in range(header_row + 1, last_data_row + 1):
                        cell = ws[f"{col_letter}{row_num}"]
                        cell.number_format = "@"
                        if cell.value is not None:
                            try:
                                cell.value = str(int(float(str(cell.value))))
                            except Exception:
                                cell.value = str(cell.value)

            data_col_count = len(sheet_df.columns)
            last_col_letter = get_column_letter(data_col_count)
            ws.auto_filter.ref = f"A{header_row}:{last_col_letter}{last_data_row}"

            for col_idx, col_cells in enumerate(ws.columns, start=1):
                max_len = 0
                for cell in col_cells:
                    try:
                        val_len = len(str(cell.value)) if cell.value is not None else 0
                        if val_len > max_len:
                            max_len = val_len
                    except Exception:
                        pass
                ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 50)

            if reason_col_idx:
                ws.column_dimensions[get_column_letter(reason_col_idx)].width = 38
            if date_col_idx:
                ws.column_dimensions[get_column_letter(date_col_idx)].width = 22
            for idx, col in enumerate(sheet_df.columns, start=1):
                if col == "Comment":
                    ws.column_dimensions[get_column_letter(idx)].width = 40
                    break

            no_fill = PatternFill(fill_type=None)
            for row_num in range(1, last_data_row + 500):
                for extra_col in range(data_col_count + 1, data_col_count + 10):
                    ws.cell(row=row_num, column=extra_col).fill = no_fill

            ws.freeze_panes = f"A{header_row+1}"
            ws.sheet_view.showGridLines = False

            if "Sector" in sheet_df.columns and len(sheet_df) > 0:
                sec_tab = str(sheet_df["Sector"].iloc[0]).strip().lower()
                sec_colors_map = {
                    "e&c sector":                 "1A4449",
                    "industrial sector":          "DC8436",
                    "supporting services sector": "8E1717",
                    "real estate sector":         "00303F",
                    "corporate":                  "243322",
                }
                ws.sheet_properties.tabColor = sec_colors_map.get(sec_tab, "243322")

    return out.getvalue()

# ── PO Summary Report Excel builder with Cleaned Data sheet ─────
def to_summary_report_excel(df, logo_path=None, open_df=None, cleaned_data_df=None):
    if logo_path is None and LOGO_PATH.exists():
        logo_path = str(LOGO_PATH)
    if open_df is None:
        open_df = df
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.drawing.image import Image as XLImage

    out = BytesIO()
    amt_col = "Total Amount (SAR)" if "Total Amount (SAR)" in df.columns else "Total amount"

    DARK_GREEN  = "243322"
    DARK_TEAL   = "1A4449"
    GOLD        = "C57E1F"
    GREEN_OK    = "3B6D11"
    GRAY_INV    = "889985"
    WHITE       = "FFFFFF"
    LIGHT_GREEN = "EAF3DE"
    LIGHT_GRAY  = "F8F8F8"
    HDR_GRAY    = "2D4A3E"

    SECTOR_COLORS = {
        "e&c sector":                 "1A4449",
        "industrial sector":          "DC8436",
        "supporting services sector": "8E1717",
        "real estate sector":         "00303F",
        "corporate":                  "243322",
    }
    DEFAULT_SECTOR_COLOR = "243322"

    def get_sector_color(sector_name):
        return SECTOR_COLORS.get(str(sector_name).strip().lower(), DEFAULT_SECTOR_COLOR)

    CARD_COLORS = ["243322", "1A4449", "DC8436", "8E1717", "00303F", "889985"]

    STATUS_COLORS = {
        "confirmed":  ("243322", "EAF3DE", "3A6B55"),
        "draft":      ("C57E1F", "FAEEDA", "854F0B"),
        "rejected":   ("A32D2D", "FCEBEB", "A32D2D"),
        "pending":    ("185FA5", "E6F1FB", "0C447C"),
        "approved":   ("3B6D11", "EAF3DE", "27500A"),
        "in review":  ("854F0B", "FAEEDA", "633806"),
        "cancelled":  ("5F5E5A", "F1EFE8", "444441"),
    }
    DEFAULT_STATUS_COLOR = ("2D4A3E", "E1F5EE", "085041")

    hdr_fill  = PatternFill(start_color=DARK_GREEN, end_color=DARK_GREEN, fill_type="solid")
    hdr_font  = Font(color=WHITE, bold=True, name="Calibri", size=11)
    hdr_align = Alignment(horizontal="left", vertical="center")

    def write_banner(ws, total_cols, title_text, subtitle_text, color=None):
        max_cols = max(total_cols, 10)
        last_col = get_column_letter(max_cols)
        bg_color = color if color else DARK_GREEN

        banner_fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        for r in range(1, 5):
            for col_i in range(1, max_cols + 1):
                ws.cell(row=r, column=col_i).fill = banner_fill

        ws.merge_cells(f"A1:{last_col}1")
        ws.merge_cells(f"A4:{last_col}4")

        logo_end_col = "B"
        ws.merge_cells(f"A2:{logo_end_col}3")
        ws.merge_cells(f"C2:{last_col}2")
        ws.merge_cells(f"C3:{last_col}3")

        title_cell = ws["C2"]
        title_cell.value     = title_text
        title_cell.font      = Font(bold=True, size=18, color=WHITE, name="Calibri")
        title_cell.alignment = Alignment(horizontal="center", vertical="bottom")
        title_cell.fill      = banner_fill

        sub_cell = ws["C3"]
        sub_cell.value     = subtitle_text
        sub_cell.font      = Font(size=10, color=WHITE, bold=True, name="Calibri")
        sub_cell.alignment = Alignment(horizontal="center", vertical="top")
        sub_cell.fill      = banner_fill

        for r in range(2, 4):
            for col_i in range(1, 3):
                ws.cell(row=r, column=col_i).fill = banner_fill

        ws.row_dimensions[1].height = 6
        ws.row_dimensions[2].height = 36
        ws.row_dimensions[3].height = 22
        ws.row_dimensions[4].height = 6
        ws.row_dimensions[5].height = 6

        if logo_path:
            try:
                img = XLImage(logo_path)
                img.height = 74
                img.width  = 180
                img.anchor = "A1"
                ws.add_image(img)
            except Exception:
                pass

    sum_bg = PatternFill(start_color=DARK_GREEN, end_color=DARK_GREEN, fill_type="solid")
    lbl_f  = Font(color="AAAAAA", name="Calibri", size=8)
    big_f  = Font(color=WHITE, bold=True, name="Calibri", size=18)
    gld_f  = Font(color="C57E1F", bold=True, name="Calibri", size=16)
    sub_f  = Font(color="AAAAAA", name="Calibri", size=8)
    cen    = Alignment(horizontal="center", vertical="center")
    lft    = Alignment(horizontal="left",   vertical="center", indent=1)

    def fill_row_range(ws, r_start, r_end, c_start, c_end, fill):
        for row_i in range(r_start, r_end+1):
            for col_i in range(c_start, c_end+1):
                ws.cell(row=row_i, column=col_i).fill = fill

    def metric_block(ws, r, col, width, label, value, sub, v_font):
        for ri in range(r, r+4):
            ws.merge_cells(start_row=ri, start_column=col, end_row=ri, end_column=col+width-1)
        fill_row_range(ws, r, r+3, col, col+width-1, sum_bg)
        ws.cell(row=r,   column=col, value=label).font = lbl_f;  ws.cell(row=r,   column=col).alignment = cen
        ws.cell(row=r+1, column=col, value=value).font = v_font; ws.cell(row=r+1, column=col).alignment = cen
        ws.cell(row=r+2, column=col, value=sub).font   = sub_f;  ws.cell(row=r+2, column=col).alignment = cen

    with pd.ExcelWriter(out, engine="openpyxl") as w:

        # ── SHEET: ENHANCED SUMMARY (MULTI-YEAR) ───────────────────
        ws_enh = w.book.create_sheet("Enhanced Summary", 1)  # Position after original Summary
        
        # Set column widths
        ws_enh.column_dimensions['A'].width = 12
        ws_enh.column_dimensions['B'].width = 18
        for i in range(3, 50):
            ws_enh.column_dimensions[get_column_letter(i)].width = 13
        
        enh_row = 1
        
        # Banner Header
        for r in range(1, 5):
            for c in range(1, 25):
                ws_enh.cell(row=r, column=c).fill = PatternFill(start_color=DARK_GREEN, end_color=DARK_GREEN, fill_type="solid")
        
        ws_enh.merge_cells("A1:X1")
        ws_enh.merge_cells("A2:X4")
        
        title_enh = ws_enh["A2"]
        title_enh.value = "Open PO Executive Summary Report (Multi-Year)"
        title_enh.font = Font(bold=True, size=16, color=WHITE, name="Calibri")
        title_enh.alignment = Alignment(horizontal="center", vertical="center")
        
        ws_enh.row_dimensions[1].height = 8
        ws_enh.row_dimensions[2].height = 28
        ws_enh.row_dimensions[3].height = 10
        ws_enh.row_dimensions[4].height = 8
        
        if logo_path:
            try:
                img_enh = XLImage(logo_path)
                img_enh.height = 60
                img_enh.width = 140
                ws_enh.add_image(img_enh, "A2")
            except:
                pass
        
        enh_row = 6
        
        # Get years sorted descending
        years_list = sorted(df["Year"].dropna().unique().tolist()) if "Year" in df.columns else []
        years_list = [int(y) for y in years_list]
        years_list.sort(reverse=True)
        
        # Build sections for each year
        for year_idx, year in enumerate(years_list):
            year_df = df[df["Year"] == year].copy()
            
            if len(year_df) == 0:
                continue
            
            # Year Header
            year_header = ws_enh.cell(row=enh_row, column=1, value=f"  SECTOR OVERVIEW {year}")
            year_header.font = Font(bold=True, size=11, color=WHITE, name="Calibri")
            year_header.fill = PatternFill(start_color=HDR_GRAY, end_color=HDR_GRAY, fill_type="solid")
            
            for c in range(1, 25):
                ws_enh.cell(row=enh_row, column=c).fill = PatternFill(start_color=HDR_GRAY, end_color=HDR_GRAY, fill_type="solid")
            
            ws_enh.merge_cells(start_row=enh_row, start_column=1, end_row=enh_row, end_column=24)
            ws_enh.row_dimensions[enh_row].height = 18
            enh_row += 1
            
            # ── KPI CARDS ──
            kpi_row = enh_row
            
            # Calculate metrics
            total_po_count = len(year_df)
            total_po_spend = year_df[amt_col].sum() if amt_col in year_df.columns else 0
            
            open_count = len(year_df[year_df.get("PO status", "") == "Open order"])
            open_spend = year_df[year_df.get("PO status", "") == "Open order"][amt_col].sum() if amt_col in year_df.columns else 0
            
            received_count = len(year_df[year_df.get("PO status", "") == "Received"])
            received_spend = year_df[year_df.get("PO status", "") == "Received"][amt_col].sum() if amt_col in year_df.columns else 0
            
            invoiced_count = len(year_df[year_df.get("PO status", "") == "Invoiced"])
            invoiced_spend = year_df[year_df.get("PO status", "") == "Invoiced"][amt_col].sum() if amt_col in year_df.columns else 0
            
            def format_spend(val):
                if val >= 1_000_000:
                    return f"SAR {val/1_000_000:.2f}M"
                elif val >= 1_000:
                    return f"SAR {val/1_000:.1f}K"
                return f"SAR {val:,.0f}"
            
            # KPI Data
            kpi_data = [
                ("TOTAL PO", total_po_count, format_spend(total_po_spend), DARK_GREEN),
                ("OPEN PO", open_count, format_spend(open_spend), DARK_TEAL),
                ("TOTAL RECEIVED", received_count, format_spend(received_spend), GREEN_OK),
                ("TOTAL INVOICED", invoiced_count, format_spend(invoiced_spend), GRAY_INV),
            ]
            
            card_col = 1
            card_width = 5
            
            for label, count, spend_str, color in kpi_data:
                card_fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                border = Border(
                    left=Side(style='thin', color='FFFFFF'),
                    right=Side(style='thin', color='FFFFFF'),
                    top=Side(style='thin', color='FFFFFF'),
                    bottom=Side(style='thin', color='FFFFFF')
                )
                
                for ri in range(kpi_row, kpi_row + 4):
                    ws_enh.merge_cells(start_row=ri, start_column=card_col, end_row=ri, end_column=card_col + card_width - 1)
                    ws_enh.row_dimensions[ri].height = 18
                
                # Label row
                lbl_cell = ws_enh.cell(row=kpi_row, column=card_col)
                lbl_cell.value = label
                lbl_cell.font = Font(name="Calibri", size=9, bold=True, color=WHITE)
                lbl_cell.fill = card_fill
                lbl_cell.alignment = Alignment(horizontal="center", vertical="center")
                lbl_cell.border = border
                
                # Count row
                cnt_cell = ws_enh.cell(row=kpi_row + 1, column=card_col)
                cnt_cell.value = count
                cnt_cell.font = Font(name="Calibri", size=18, bold=True, color=WHITE)
                cnt_cell.fill = card_fill
                cnt_cell.alignment = Alignment(horizontal="center", vertical="center")
                cnt_cell.number_format = "#,##0"
                cnt_cell.border = border
                
                # Spend label row
                spd_lbl_cell = ws_enh.cell(row=kpi_row + 2, column=card_col)
                spd_lbl_cell.value = "Total Spend"
                spd_lbl_cell.font = Font(name="Calibri", size=8, color="CCCCCC")
                spd_lbl_cell.fill = card_fill
                spd_lbl_cell.alignment = Alignment(horizontal="center", vertical="center")
                spd_lbl_cell.border = border
                
                # Spend value row
                spd_cell = ws_enh.cell(row=kpi_row + 3, column=card_col)
                spd_cell.value = spend_str
                spd_cell.font = Font(name="Calibri", size=10, bold=True, color="FFEB99")
                spd_cell.fill = card_fill
                spd_cell.alignment = Alignment(horizontal="center", vertical="center")
                spd_cell.border = border
                
                card_col += card_width
            
            enh_row = kpi_row + 5
            enh_row += 1
            
            # ── ENTITY DETAIL TABLE ──
            headers = [
                "Entity", "Sector",
                "Total PO", "Total PO Spend",
                "Open Order", "Open Order Spend",
                "Received", "Received Spend",
                "Invoiced", "Invoiced Spend",
                "Finalized", "Finalized Spend",
                "Approved", "Approved Spend",
                "Confirmed", "Confirmed Spend",
                "Draft", "Draft Spend",
                "InReview", "InReview Spend",
                "Rejected", "Rejected Spend",
            ]
            
            hdr_fill = PatternFill(start_color=DARK_GREEN, end_color=DARK_GREEN, fill_type="solid")
            hdr_font = Font(color=WHITE, bold=True, name="Calibri", size=8)
            hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
            for col_idx, header in enumerate(headers, 1):
                hdr_cell = ws_enh.cell(row=enh_row, column=col_idx, value=header)
                hdr_cell.fill = hdr_fill
                hdr_cell.font = hdr_font
                hdr_cell.alignment = hdr_align
            
            ws_enh.row_dimensions[enh_row].height = 20
            enh_row += 1
            
            # Write entity rows
            entities = sorted(year_df["Entity"].dropna().unique().tolist()) if "Entity" in year_df.columns else []
            
            for entity_idx, entity in enumerate(entities):
                entity_df = year_df[year_df["Entity"] == entity]
                sector = entity_df["Sector"].iloc[0] if "Sector" in entity_df.columns and len(entity_df) > 0 else ""
                
                row_fill = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid") if entity_idx % 2 == 1 else PatternFill(fill_type=None)
                
                row_data = [entity, sector]
                
                # Total PO
                total_count = len(entity_df)
                total_spend = entity_df[amt_col].sum() if amt_col in entity_df.columns else 0
                row_data.extend([total_count, total_spend])
                
                # PO Status breakdown
                for status in ["Open order", "Received", "Invoiced"]:
                    status_df = entity_df[entity_df.get("PO status", "") == status]
                    status_count = len(status_df)
                    status_spend = status_df[amt_col].sum() if amt_col in status_df.columns else 0
                    row_data.extend([status_count, status_spend])
                
                # Approval Status breakdown
                open_po_df = entity_df[entity_df.get("PO status", "") == "Open order"]
                for appr_status in ["Finalized", "Approved", "Confirmed", "Draft", "InReview", "Rejected"]:
                    appr_df = open_po_df[open_po_df.get("Approval status", "") == appr_status]
                    appr_count = len(appr_df)
                    appr_spend = appr_df[amt_col].sum() if amt_col in appr_df.columns else 0
                    row_data.extend([appr_count, appr_spend])
                
                # Write row
                for col_idx, value in enumerate(row_data, 1):
                    cell = ws_enh.cell(row=enh_row, column=col_idx, value=value)
                    cell.fill = row_fill
                    cell.alignment = Alignment(
                        horizontal="right" if isinstance(value, (int, float)) else "left",
                        vertical="center"
                    )
                    cell.font = Font(name="Calibri", size=9)
                    
                    if isinstance(value, float):
                        cell.number_format = "#,##0.00"
                    elif isinstance(value, int) and col_idx > 2:
                        cell.number_format = "#,##0"
                
                ws_enh.row_dimensions[enh_row].height = 14
                enh_row += 1
            
            enh_row += 2
        
        ws_enh.freeze_panes = "A7"
        ws_enh.sheet_view.showGridLines = False


        # ── SHEETS 2+: ENTITY + YEAR ──────────────────────────
        if "Entity" in df.columns and "Year" in df.columns:
            combos = open_df.groupby(["Entity", "Year"]).size().reset_index()
            combos = combos.sort_values(["Entity", "Year"])

            for _, combo in combos.iterrows():
                entity = str(combo["Entity"])
                year   = str(int(combo["Year"]))
                sheet_name = f"{entity} {year}"[:31]

                sheet_df = open_df[(open_df["Entity"] == combo["Entity"]) & (open_df["Year"] == combo["Year"])].copy()

                if "Sector" in sheet_df.columns and len(sheet_df) > 0:
                    sector_val_tab = sheet_df["Sector"].iloc[0]
                    tab_color = get_sector_color(sector_val_tab)
                else:
                    tab_color = DARK_GREEN

                for col_name in ["Open PO Reason", "Expected Closure Date", "Comment"]:
                    if col_name not in sheet_df.columns:
                        sheet_df[col_name] = ""

                total_cols = len(sheet_df.columns)

                sheet_df.to_excel(w, index=False, sheet_name=sheet_name, startrow=18)
                ws = w.sheets[sheet_name]

                write_banner(ws, total_cols,
                             "Open PO Tracker",
                             f"{year}",
                             color=tab_color)

                approval_col = "Approval status" if "Approval status" in sheet_df.columns else None
                vendor_col   = "Vendor name" if "Vendor name" in sheet_df.columns else None
                sector_val   = sheet_df["Sector"].iloc[0] if "Sector" in sheet_df.columns and len(sheet_df) > 0 else ""

                banner_row = 6
                total_spend_val = sheet_df[amt_col].sum() if amt_col in sheet_df.columns else 0
                unique_vendors  = sheet_df[vendor_col].nunique() if vendor_col else 0
                total_pos       = len(sheet_df)

                sec_theme   = tab_color
                banner_fill = PatternFill(start_color=sec_theme, end_color=sec_theme, fill_type="solid")
                wht_font    = Font(color=WHITE, bold=True, name="Calibri", size=10)
                muted_font  = Font(color="BBBBBB", name="Calibri", size=8)
                big_font    = Font(color=WHITE, bold=True, name="Calibri", size=22)
                gold_font   = Font(color="C57E1F", bold=True, name="Calibri", size=18)
                center_al   = Alignment(horizontal="center", vertical="center")

                for r in range(banner_row, banner_row + 4):
                    ws.row_dimensions[r].height = 18
                    for c in range(1, total_cols + 1):
                        ws.cell(row=r, column=c).fill = banner_fill

                sec_width = max(total_cols // 4, 3)

                def banner_metric(ws, row, col, label, value, sub, val_font, lbl_font, sub_font, align):
                    for mi in range(row, row+4):
                        ws.merge_cells(start_row=mi, start_column=col, end_row=mi, end_column=col+sec_width-1)
                        ws.cell(row=mi, column=col).fill = banner_fill
                    ws.cell(row=row,   column=col, value=label).font = lbl_font
                    ws.cell(row=row,   column=col).alignment = align
                    ws.cell(row=row+1, column=col, value=value).font = val_font
                    ws.cell(row=row+1, column=col).alignment = align
                    ws.cell(row=row+2, column=col, value=sub).font   = sub_font
                    ws.cell(row=row+2, column=col).alignment = align

                banner_metric(ws, banner_row, 1,
                              "ENTITY", f"{entity} · {year}", sector_val,
                              Font(color=WHITE, bold=True, name="Calibri", size=14),
                              Font(color="AAAAAA", name="Calibri", size=8),
                              Font(color="AAAAAA", name="Calibri", size=8), center_al)

                banner_metric(ws, banner_row, 1+sec_width,
                              "TOTAL OPEN POs", total_pos, "All statuses",
                              big_font, Font(color="AAAAAA", name="Calibri", size=8),
                              Font(color="AAAAAA", name="Calibri", size=8), center_al)

                spend_display = f"SAR {total_spend_val/1_000_000:.2f}M" if total_spend_val >= 1_000_000 else f"SAR {total_spend_val:,.0f}"
                banner_metric(ws, banner_row, 1+sec_width*2,
                              "TOTAL SPEND", spend_display, "Open PO value",
                              gold_font, Font(color="AAAAAA", name="Calibri", size=8),
                              Font(color="AAAAAA", name="Calibri", size=8), center_al)

                banner_metric(ws, banner_row, 1+sec_width*3,
                              "UNIQUE VENDORS", unique_vendors, "Active suppliers",
                              big_font, Font(color="AAAAAA", name="Calibri", size=8),
                              Font(color="AAAAAA", name="Calibri", size=8), center_al)

                ws.row_dimensions[banner_row + 4].height = 6

                card_start_row = banner_row + 5

                if approval_col:
                    statuses   = sorted(sheet_df[approval_col].dropna().unique().tolist())
                    total_pos2 = len(sheet_df)
                    card_col   = 1
                    card_w     = max(total_cols // (len(statuses) + 1), 3)

                    for i, status in enumerate(statuses):
                        s_df     = sheet_df[sheet_df[approval_col] == status]
                        po_count = len(s_df)
                        spend    = s_df[amt_col].sum() if amt_col in s_df.columns else 0
                        pct      = round(po_count / total_pos2 * 100) if total_pos2 > 0 else 0

                        sc = STATUS_COLORS.get(status.lower(), DEFAULT_STATUS_COLOR)
                        status_color, badge_bg, badge_txt = sc

                        card_fill  = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                        top_fill   = PatternFill(start_color=sec_theme, end_color=sec_theme, fill_type="solid")
                        top_color  = sec_theme
                        badge_fill = PatternFill(start_color=badge_bg, end_color=badge_bg, fill_type="solid")

                        for c in range(card_col, card_col + card_w):
                            ws.merge_cells(start_row=card_start_row, start_column=card_col, end_row=card_start_row, end_column=card_col+card_w-1)
                            ws.cell(row=card_start_row, column=card_col).fill = top_fill
                            break
                        ws.row_dimensions[card_start_row].height = 4

                        for r in range(card_start_row+1, card_start_row+7):
                            ws.merge_cells(start_row=r, start_column=card_col, end_row=r, end_column=card_col+card_w-1)
                            ws.cell(row=r, column=card_col).fill = card_fill
                            ws.row_dimensions[r].height = 16

                        lbl = ws.cell(row=card_start_row+1, column=card_col, value=status.upper())
                        lbl.font      = Font(name="Calibri", size=8, bold=True, color="888888")
                        lbl.alignment = Alignment(horizontal="left", vertical="center", indent=1)

                        cnt = ws.cell(row=card_start_row+2, column=card_col, value=po_count)
                        cnt.font      = Font(name="Calibri", size=20, bold=True, color=top_color)
                        cnt.alignment = Alignment(horizontal="left", vertical="center", indent=1)

                        sub1 = ws.cell(row=card_start_row+3, column=card_col, value="Open POs")
                        sub1.font      = Font(name="Calibri", size=8, color="AAAAAA")
                        sub1.alignment = Alignment(horizontal="left", vertical="center", indent=1)

                        spend_str = f"SAR {spend/1_000_000:.2f}M" if spend >= 1_000_000 else f"SAR {spend:,.0f}"
                        spd = ws.cell(row=card_start_row+4, column=card_col, value=spend_str)
                        spd.font      = Font(name="Calibri", size=11, bold=True, color="C57E1F")
                        spd.alignment = Alignment(horizontal="left", vertical="center", indent=1)

                        sub2 = ws.cell(row=card_start_row+5, column=card_col, value=f"Total Spend  |  {pct}% of POs")
                        sub2.font      = Font(name="Calibri", size=8, color="AAAAAA")
                        sub2.alignment = Alignment(horizontal="left", vertical="center", indent=1)

                        card_col += card_w

                    ws.row_dimensions[card_start_row + 7].height = 8

                header_row = card_start_row + 8
                sec_hdr_fill = PatternFill(start_color=sec_theme, end_color=sec_theme, fill_type="solid")
                for cell in ws[header_row]:
                    cell.fill      = sec_hdr_fill
                    cell.font      = hdr_font
                    cell.alignment = hdr_align

                last_data_row = header_row + len(sheet_df)
                data_align = Alignment(horizontal="left", vertical="center")
                for drow in ws.iter_rows(min_row=header_row+1, max_row=last_data_row):
                    for cell in drow:
                        cell.alignment = data_align

                for idx, col_name in enumerate(sheet_df.columns, start=1):
                    col_letter = get_column_letter(idx)
                    if col_name in ["VAT number", "CR number"]:
                        for rn in range(header_row+1, last_data_row+1):
                            cell = ws[f"{col_letter}{rn}"]
                            cell.number_format = "@"
                            if cell.value is not None:
                                try:    cell.value = str(int(float(str(cell.value))))
                                except: cell.value = str(cell.value)
                    elif col_name in ["Total Amount (SAR)", "Total amount"]:
                        has_decimals = any(
                            v % 1 != 0
                            for v in sheet_df[col_name].dropna()
                            if isinstance(v, (int, float))
                        )
                        amt_fmt = "#,##0.00" if has_decimals else "#,##0"
                        for rn in range(header_row+1, last_data_row+1):
                            ws[f"{col_letter}{rn}"].number_format = amt_fmt

                OPEN_REASONS_SUM = [
                    "Pending Delivery",
                    "Pending Invoice / Payment Processing",
                    "Pending Documentation / Approval",
                    "Partial Receipt / Quantity Adjustment",
                    "Pending PO Closure",
                    "Open PO Linked to Project Duration",
                    "Under Internal Review",
                    "Other",
                ]
                from openpyxl.worksheet.datavalidation import DataValidation as DV2
                for idx, col_name in enumerate(sheet_df.columns, start=1):
                    col_letter = get_column_letter(idx)
                    if col_name == "Open PO Reason":
                        reason_formula = '"' + ",".join(OPEN_REASONS_SUM) + '"'
                        dv_r = DV2(type="list", formula1=reason_formula,
                                   allow_blank=True, showDropDown=False,
                                   showErrorMessage=True, errorTitle="Invalid",
                                   error="Please select a reason from the list.")
                        ws.add_data_validation(dv_r)
                        dv_r.sqref = f"{col_letter}{header_row+1}:{col_letter}{last_data_row+500}"
                        ws.column_dimensions[col_letter].width = 38
                        no_fill2 = PatternFill(fill_type=None)
                        for rn in range(header_row+1, last_data_row+500):
                            ws[f"{col_letter}{rn}"].fill = no_fill2

                    elif col_name == "Expected Closure Date":
                        month_list = closure_month_options()
                        date_formula = '"' + ",".join(month_list) + '"'
                        dv_d = DV2(type="list", formula1=date_formula,
                                   allow_blank=True, showDropDown=False,
                                   showErrorMessage=True, errorTitle="Invalid",
                                   error="Please select a month from the list.")
                        ws.add_data_validation(dv_d)
                        dv_d.sqref = f"{col_letter}{header_row+1}:{col_letter}{last_data_row+500}"
                        ws.column_dimensions[col_letter].width = 22
                        no_fill3 = PatternFill(fill_type=None)
                        for rn in range(header_row+1, last_data_row+500):
                            ws[f"{col_letter}{rn}"].fill = no_fill3

                    elif col_name == "Comment":
                        ws.column_dimensions[col_letter].width = 40

                last_col_letter = get_column_letter(total_cols)
                ws.auto_filter.ref = f"A{header_row}:{last_col_letter}{last_data_row}"

                for col_idx, col_cells in enumerate(ws.columns, start=1):
                    col_hdr = ws.cell(row=header_row, column=col_idx).value
                    if col_hdr in ["Open PO Reason", "Expected Closure Date", "Comment"]:
                        continue
                    max_len = max((len(str(c.value)) for c in col_cells if c.value), default=8)
                    ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 50)

                no_fill = PatternFill(fill_type=None)
                for rn in range(1, last_data_row + 50):
                    for ec in range(total_cols + 1, total_cols + 10):
                        ws.cell(row=rn, column=ec).fill = no_fill

                ws.freeze_panes = f"A{header_row+1}"
                ws.sheet_view.showGridLines = False
                ws.sheet_properties.tabColor = tab_color

        # ── ADD CLEANED DATA SHEET AT THE END ──────────────────
        if cleaned_data_df is not None and not cleaned_data_df.empty:
            cleaned_data_df.to_excel(w, index=False, sheet_name="Cleaned Data")
            ws_clean = w.sheets["Cleaned Data"]

            hdr_fill_clean = PatternFill(start_color="243322", end_color="243322", fill_type="solid")
            hdr_font_clean = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
            hdr_align_clean = Alignment(horizontal="left", vertical="center")

            for cell in ws_clean[1]:
                cell.fill = hdr_fill_clean
                cell.font = hdr_font_clean
                cell.alignment = hdr_align_clean

            data_align = Alignment(horizontal="left", vertical="center")
            for row in ws_clean.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = data_align

            for col_idx, col_cells in enumerate(ws_clean.columns, start=1):
                max_len = 0
                for cell in col_cells:
                    try:
                        cell_len = len(str(cell.value)) if cell.value is not None else 0
                        if cell_len > max_len:
                            max_len = cell_len
                    except Exception:
                        pass
                ws_clean.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 50)

            text_cols = ["VAT number", "CR number"]
            for idx, col_cells in enumerate(ws_clean.columns, start=1):
                header = ws_clean.cell(row=1, column=idx).value
                if header in text_cols:
                    col_letter = get_column_letter(idx)
                    for row_num in range(2, ws_clean.max_row + 1):
                        cell = ws_clean[f"{col_letter}{row_num}"]
                        cell.number_format = "@"
                        if cell.value is not None:
                            try:
                                cell.value = str(int(float(str(cell.value))))
                            except Exception:
                                cell.value = str(cell.value)

            last_col = get_column_letter(ws_clean.max_column)
            last_row = ws_clean.max_row
            ws_clean.auto_filter.ref = f"A1:{last_col}{last_row}"

    return out.getvalue()
