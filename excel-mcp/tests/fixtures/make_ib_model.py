"""Deterministic generator for the real-model IB golden fixture (R1, R2, D14)."""

import sys
from pathlib import Path

_EXCEL_MCP_DIR = Path(__file__).resolve().parents[2]
if str(_EXCEL_MCP_DIR) not in sys.path:
    sys.path.insert(0, str(_EXCEL_MCP_DIR))

import core.format_headless  # noqa: F401 - activates header/footer parsing patch
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.properties import PageSetupProperties


def build(path: str) -> None:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()

    # Reusable styles
    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    input_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    header_font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    body_font = Font(name="Arial", size=10, bold=False)
    subtotal_font = Font(name="Arial", size=10, bold=True)
    total_font = Font(name="Arial", size=10, bold=True)

    int_format = '_-* #,##0_-;-* #,##0_-;-* "-"_-;_-@_-'
    pct_format = "%_0.0%"

    thin_border_side = Side(style="thin", color="D9D9D9")
    med_border_side = Side(style="medium", color="D9D9D9")
    double_border_side = Side(style="double", color="D9D9D9")

    grid_border = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=thin_border_side,
    )
    subtotal_border = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=med_border_side,
    )
    total_border = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=double_border_side,
    )

    def apply_standard_page_setup(ws):
        ws.freeze_panes = "B2"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.oddFooter.left.text = ""
        ws.oddFooter.center.text = "&F | &A"
        ws.oddFooter.right.text = ""
        for col_letter in ("A", "B", "C", "D", "E", "F"):
            ws.column_dimensions[col_letter].width = 16
        for r in range(1, 30):
            ws.row_dimensions[r].height = 15

    # 1. Income Statement (IS)
    ws_is = wb.active
    ws_is.title = "IS"
    apply_standard_page_setup(ws_is)

    is_data = [
        ["Line Item", "2023A", "2024A", "2025E"],
        ["Revenue", 100000, 120000, 150000],
        ["Cost of Goods Sold", 40000, 48000, 60000],
        ["Gross Profit", "=B2-B3", "=C2-C3", "=D2-D3"],
        ["SG&A Expense", 20000, 24000, 30000],
        ["EBITDA", "=B4-B5", "=C4-C5", "=D4-D5"],
        ["Depreciation & Amortization", 5000, 6000, 7500],
        ["EBIT", "=B6-B7", "=C6-C7", "=D6-D7"],
        ["Interest Expense", 2000, 1800, 1500],
        ["EBT", "=B8-B9", "=C8-C9", "=D8-D9"],
        ["Tax", "=B10*0.25", "=C10*0.25", "=D10*0.25"],
        ["Net Income", "=B10-B11", "=C10-C11", "=D10-D11"],
    ]

    for row_idx, row in enumerate(is_data, start=1):
        for col_idx, val in enumerate(row, start=1):
            cell = ws_is.cell(row=row_idx, column=col_idx, value=val)
            if row_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.border = grid_border
                cell.alignment = Alignment(horizontal="center" if col_idx > 1 else "left")
            elif row_idx in (4, 6, 8, 10):
                cell.font = subtotal_font
                cell.border = subtotal_border
                if col_idx > 1:
                    cell.number_format = int_format
            elif row_idx == 12:
                cell.font = total_font
                cell.border = total_border
                if col_idx > 1:
                    cell.number_format = int_format
            else:
                cell.font = body_font
                cell.border = grid_border
                if col_idx > 1:
                    cell.number_format = int_format
                    if row_idx in (2, 3, 5, 7, 9):
                        cell.fill = input_fill

    # 2. Balance Sheet (BS)
    ws_bs = wb.create_sheet("BS")
    apply_standard_page_setup(ws_bs)

    bs_data = [
        ["Line Item", "2023A", "2024A", "2025E"],
        ["Cash & Cash Equivalents", "=CF!B10", "=CF!C10", "=CF!D10"],
        ["Accounts Receivable", 15000, 18000, 22500],
        ["Inventory", 10000, 12000, 15000],
        ["Current Assets", "=SUM(B2:B4)", "=SUM(C2:C4)", "=SUM(D2:D4)"],
        ["Property, Plant & Equipment", 50000, 55000, 60000],
        ["Total Assets", "=B5+B6", "=C5+C6", "=D5+D6"],
        ["Accounts Payable", 8000, 9600, 12000],
        ["Long Term Debt", 25000, 20000, 15000],
        ["Total Liabilities", "=B8+B9", "=C8+C9", "=D8+D9"],
        ["Common Stock", 20000, 20000, 20000],
        ["Retained Earnings", 22000, "=B12+IS!C12", "=C12+IS!D12"],
        ["Total Liabilities & Equity", "=B10+B11+B12", "=C10+C11+C12", "=D10+D11+D12"],
    ]

    for row_idx, row in enumerate(bs_data, start=1):
        for col_idx, val in enumerate(row, start=1):
            cell = ws_bs.cell(row=row_idx, column=col_idx, value=val)
            if row_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.border = grid_border
                cell.alignment = Alignment(horizontal="center" if col_idx > 1 else "left")
            elif row_idx in (5, 7, 10):
                cell.font = subtotal_font
                cell.border = subtotal_border
                if col_idx > 1:
                    cell.number_format = int_format
            elif row_idx == 13:
                cell.font = total_font
                cell.border = total_border
                if col_idx > 1:
                    cell.number_format = int_format
            else:
                cell.font = body_font
                cell.border = grid_border
                if col_idx > 1:
                    cell.number_format = int_format
                    if row_idx in (3, 4, 6, 8, 9, 11):
                        cell.fill = input_fill

    # 3. Cash Flow Statement (CF)
    ws_cf = wb.create_sheet("CF")
    apply_standard_page_setup(ws_cf)

    cf_data = [
        ["Line Item", "2023A", "2024A", "2025E"],
        ["Net Income", "=IS!B12", "=IS!C12", "=IS!D12"],
        ["Depreciation", "=IS!B7", "=IS!C7", "=IS!D7"],
        ["Change in Working Capital", -2000, -2500, -3000],
        ["Cash Flow from Operations", "=SUM(B2:B4)", "=SUM(C2:C4)", "=SUM(D2:D4)"],
        ["Capital Expenditures", -5000, -6000, -7000],
        ["Cash Flow from Investing", "=B6", "=C6", "=D6"],
        ["Debt Repayment", -5000, -5000, -5000],
        ["Cash Flow from Financing", "=B8", "=C8", "=D8"],
        ["Ending Cash", "=B5+B7+B9", "=C5+C7+C9", "=D5+D7+D9"],
    ]

    for row_idx, row in enumerate(cf_data, start=1):
        for col_idx, val in enumerate(row, start=1):
            cell = ws_cf.cell(row=row_idx, column=col_idx, value=val)
            if row_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.border = grid_border
                cell.alignment = Alignment(horizontal="center" if col_idx > 1 else "left")
            elif row_idx in (5, 7, 9):
                cell.font = subtotal_font
                cell.border = subtotal_border
                if col_idx > 1:
                    cell.number_format = int_format
            elif row_idx == 10:
                cell.font = total_font
                cell.border = total_border
                if col_idx > 1:
                    cell.number_format = int_format
            else:
                cell.font = body_font
                cell.border = grid_border
                if col_idx > 1:
                    cell.number_format = int_format
                    if row_idx in (4, 6, 8):
                        cell.fill = input_fill

    # Flag intentional circular cell per R2
    # Flagged with CIRCULAR_INTENTIONAL label and formula CF!X1 = CF!X1+1
    ws_cf["W1"] = "CIRCULAR_INTENTIONAL"
    ws_cf["W1"].font = body_font
    ws_cf["X1"] = "=CF!X1+1"
    ws_cf["X1"].font = body_font

    # 4. Comps Sheet (Comps)
    ws_comps = wb.create_sheet("Comps")
    apply_standard_page_setup(ws_comps)

    comps_data = [
        ["Company", "Category", "Revenue", "EBITDA", "EV", "EV/EBITDA"],
        ["Alpha Inc", "Software", 50000, 15000, 150000, "=E2/D2"],
        ["Beta Corp", "Hardware", 80000, 16000, 120000, "=E3/D3"],
        ["Gamma Ltd", "Software", 45000, 13500, 135000, "=E4/D4"],
        ["Delta LLC", "Hardware", 70000, 14000, 98000, "=E5/D5"],
        ["Total / Average", "All", "=SUM(C2:C5)", "=SUM(D2:D5)", "=SUM(E2:E5)", "=AVERAGE(F2:F5)"],
    ]

    for row_idx, row in enumerate(comps_data, start=1):
        for col_idx, val in enumerate(row, start=1):
            cell = ws_comps.cell(row=row_idx, column=col_idx, value=val)
            if row_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.border = grid_border
                cell.alignment = Alignment(horizontal="center" if col_idx > 2 else "left")
            elif row_idx == 6:
                cell.font = total_font
                cell.border = total_border
                if col_idx in (3, 4, 5):
                    cell.number_format = int_format
                elif col_idx == 6:
                    cell.number_format = pct_format
            else:
                cell.font = body_font
                cell.border = grid_border
                if col_idx in (3, 4, 5):
                    cell.number_format = int_format
                    cell.fill = input_fill
                elif col_idx == 6:
                    cell.number_format = pct_format

    # SUMIFS pivot-summary table (D3 fallback) on Comps sheet at A10:C13
    summary_headers = ["Category", "Total Revenue (SUMIFS)", "Total EBITDA (SUMIFS)"]
    for col_idx, h in enumerate(summary_headers, start=1):
        cell = ws_comps.cell(row=10, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = grid_border

    categories = [
        (
            "Software",
            '=SUMIFS(C$2:C$5, B$2:B$5, "Software")',
            '=SUMIFS(D$2:D$5, B$2:B$5, "Software")',
        ),
        (
            "Hardware",
            '=SUMIFS(C$2:C$5, B$2:B$5, "Hardware")',
            '=SUMIFS(D$2:D$5, B$2:B$5, "Hardware")',
        ),
        ("Summary Total", "=SUM(B11:B12)", "=SUM(C11:C12)"),
    ]

    for r_offset, (cat, rev_f, ebitda_f) in enumerate(categories, start=11):
        c1 = ws_comps.cell(row=r_offset, column=1, value=cat)
        c2 = ws_comps.cell(row=r_offset, column=2, value=rev_f)
        c3 = ws_comps.cell(row=r_offset, column=3, value=ebitda_f)
        if r_offset == 13:
            for c in (c1, c2, c3):
                c.font = total_font
                c.border = total_border
        else:
            for c in (c1, c2, c3):
                c.font = body_font
                c.border = grid_border
        c2.number_format = int_format
        c3.number_format = int_format

    # Add Chart on Comps sheet per R1
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Comps Revenue & EBITDA"
    chart.y_axis.title = "Amount ($)"
    chart.x_axis.title = "Company"
    data = Reference(ws_comps, min_col=3, min_row=1, max_col=4, max_row=5)
    cats = Reference(ws_comps, min_col=1, min_row=2, max_row=5)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 10
    chart.width = 16
    ws_comps.add_chart(chart, "H2")

    wb.save(str(p))
    wb.close()


if __name__ == "__main__":
    build(str(Path(__file__).parent / "ib_model.xlsx"))
