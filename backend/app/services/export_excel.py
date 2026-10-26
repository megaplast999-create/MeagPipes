"""Standalone Excel customer quotation export service using openpyxl.

Generates password-free, fully editable .xlsx customer workbooks
without Excel or Windows COM dependencies, running cleanly on Linux/Render.
"""

from __future__ import annotations

import io
from datetime import datetime
from typing import Any, Dict, List, Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter


def number_to_indian_words(amount: float) -> str:
    """Converts a monetary amount into standard Indian currency words (Crores, Lakhs, Thousands)."""
    num = int(amount)
    if num == 0:
        return "Rupees Zero Only"

    crore = num // 10000000
    num = num % 10000000

    lakh = num // 100000
    num = num % 100000

    thousand = num // 1000
    num = num % 1000

    hundred = num // 100
    remainder = num % 100

    parts = []

    def convert_two_digits(n: int) -> str:
        units = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
                 "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
                 "Seventeen", "Eighteen", "Nineteen"]
        tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
        if n < 20:
            return units[n]
        t = n // 10
        u = n % 10
        return f"{tens[t]} {units[u]}".strip()

    if crore > 0:
        parts.append(f"{convert_two_digits(crore)} Crore")
    if lakh > 0:
        parts.append(f"{convert_two_digits(lakh)} Lakh")
    if thousand > 0:
        parts.append(f"{convert_two_digits(thousand)} Thousand")
    if hundred > 0:
        parts.append(f"{convert_two_digits(hundred)} Hundred")
    if remainder > 0:
        parts.append(convert_two_digits(remainder))

    words_str = " ".join(parts).strip()
    return f"Rupees {words_str} Only"


class ExcelExportService:
    """Generates customer-ready, unlocked, standalone Excel quotations."""

    # Styling Palettes (Mega Pipes Corporate Theme)
    FONT_NAME = "Calibri"
    HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    SUBHEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    ACCENT_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    TOTAL_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    FONT_HEADER = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
    FONT_TITLE = Font(name=FONT_NAME, size=16, bold=True, color="1F4E79")
    FONT_BOLD = Font(name=FONT_NAME, size=11, bold=True)
    FONT_REGULAR = Font(name=FONT_NAME, size=11)
    FONT_MUTED = Font(name=FONT_NAME, size=9, italic=True, color="595959")

    BORDER_THIN = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3"),
    )
    BORDER_DOUBLE_BOTTOM = Border(
        top=Side(style="thin", color="000000"),
        bottom=Side(style="double", color="000000"),
    )

    UNLOCKED_PROTECTION = Protection(locked=False)

    @classmethod
    def generate_quotation_workbook(
        cls,
        quotation_data: Dict[str, Any],
        calculation_results: Dict[str, Any],
    ) -> io.BytesIO:
        """Builds a complete, unlocked, standalone customer quotation workbook with Offer and Synopsis sheets."""
        wb = openpyxl.Workbook()
        # Default sheet -> OFFER
        ws_offer = wb.active
        ws_offer.title = "OFFER"
        ws_synopsis = wb.create_sheet(title="Synopsis")

        cls._populate_offer_sheet(ws_offer, quotation_data, calculation_results)
        cls._populate_synopsis_sheet(ws_synopsis, quotation_data, calculation_results)

        # Ensure NO worksheet or workbook protection is active anywhere
        for ws in wb.worksheets:
            ws.protection.disable()
            # Explicitly unlock all populated cells
            for row in ws.iter_rows():
                for cell in row:
                    cell.protection = cls.UNLOCKED_PROTECTION

        wb.security.lockStructure = False
        wb.security.lockWindows = False

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output

    @classmethod
    def _populate_offer_sheet(
        cls,
        ws: openpyxl.worksheet.worksheet.Worksheet,
        quote: Dict[str, Any],
        calc: Dict[str, Any],
    ) -> None:
        """Fills the commercial Offer sheet with metadata, lines, and retained local formulas."""
        ws.views.sheetView[0].showGridLines = True

        # Company Header
        ws["A1"] = "MEGA PIPES — COMMERCIAL OFFER"
        ws["A1"].font = cls.FONT_TITLE
        ws["A2"] = f"Generated on {datetime.now().strftime('%d-%b-%Y')} | Valid for 15 Days"
        ws["A2"].font = cls.FONT_MUTED

        # Customer & Project Details
        ws["A4"] = "Customer Name:"
        ws["B4"] = quote.get("customer_name", "N/A")
        ws["D4"] = "Enquiry No:"
        ws["E4"] = quote.get("enquiry_no", "N/A")

        ws["A5"] = "Project Name:"
        ws["B5"] = quote.get("project_name", "N/A")
        ws["D5"] = "Date:"
        ws["E5"] = datetime.now().strftime("%d-%m-%Y")

        ws["A6"] = "Delivery Location:"
        ws["B6"] = quote.get("delivery_location", "N/A")
        ws["D6"] = "Mode:"
        ws["E6"] = calc.get("calculation_mode", "REFERENCE")

        for r in range(4, 7):
            ws[f"A{r}"].font = cls.FONT_BOLD
            ws[f"D{r}"].font = cls.FONT_BOLD

        # Table Column Headers
        headers = [
            ("A8", "Item"),
            ("B8", "Nominal Size"),
            ("C8", "Wall Thk (mm)"),
            ("D8", "OD (mm)"),
            ("E8", "ID (mm)"),
            ("F8", "Quantity"),
            ("G8", "Unit"),
            ("H8", "Unit Rate (₹)"),
            ("I8", "Total Amount (₹)"),
        ]
        for pos, title in headers:
            cell = ws[pos]
            cell.value = title
            cell.fill = cls.HEADER_FILL
            cell.font = cls.FONT_HEADER
            cell.alignment = Alignment(horizontal="center", vertical="center")

        lines = calc.get("lines", [])
        start_row = 9
        current_row = start_row

        for idx, line in enumerate(lines, start=1):
            applied_rate = line.get("applied_unit_price", line.get("total_per_m", 0.0))
            quantity = line.get("metres", line.get("quantity", 0.0))

            ws[f"A{current_row}"] = idx
            ws[f"B{current_row}"] = f"{line.get('size_mm', 0):.0f} mm ({line.get('basis', 'ID')})"
            ws[f"C{current_row}"] = line.get("thickness_mm", 0.0)
            ws[f"D{current_row}"] = round(line.get("od_mm", 0.0), 1)
            ws[f"E{current_row}"] = round(line.get("inner_d_mm", 0.0), 1)
            ws[f"F{current_row}"] = quantity
            ws[f"G{current_row}"] = line.get("unit", "Metres")
            ws[f"H{current_row}"] = applied_rate

            # Local dynamic Excel formula for Line Amount: Quantity * Unit Rate
            ws[f"I{current_row}"] = f"=F{current_row}*H{current_row}"

            # Number Formatting
            ws[f"A{current_row}"].alignment = Alignment(horizontal="center")
            ws[f"C{current_row}"].number_format = "#,##0.0"
            ws[f"D{current_row}"].number_format = "#,##0.0"
            ws[f"E{current_row}"].number_format = "#,##0.0"
            ws[f"F{current_row}"].number_format = "#,##0"
            ws[f"H{current_row}"].number_format = "₹#,##0"
            ws[f"I{current_row}"].number_format = "₹#,##0"

            for col in "ABCDEFGHI":
                c = ws[f"{col}{current_row}"]
                c.border = cls.BORDER_THIN
                c.font = cls.FONT_REGULAR

            current_row += 1

        last_data_row = current_row - 1 if lines else start_row

        # Summary Rows (Subtotal, GST, Grand Total)
        subtotal_row = current_row
        gst_row = current_row + 1
        grand_total_row = current_row + 2

        # Subtotal formula
        ws[f"G{subtotal_row}"] = "Subtotal (Excl. GST):"
        ws[f"G{subtotal_row}"].font = cls.FONT_BOLD
        ws[f"I{subtotal_row}"] = f"=SUM(I{start_row}:I{last_data_row})"
        ws[f"I{subtotal_row}"].font = cls.FONT_BOLD
        ws[f"I{subtotal_row}"].number_format = "₹#,##0"
        ws[f"I{subtotal_row}"].border = cls.BORDER_THIN

        # GST formula (18%)
        ws[f"G{gst_row}"] = "GST @ 18%:"
        ws[f"G{gst_row}"].font = cls.FONT_BOLD
        ws[f"I{gst_row}"] = f"=ROUND(I{subtotal_row}*0.18, 0)"
        ws[f"I{gst_row}"].font = cls.FONT_BOLD
        ws[f"I{gst_row}"].number_format = "₹#,##0"
        ws[f"I{gst_row}"].border = cls.BORDER_THIN

        # Grand Total formula
        ws[f"G{grand_total_row}"] = "Grand Total (Incl. GST):"
        ws[f"G{grand_total_row}"].font = cls.FONT_BOLD
        ws[f"I{grand_total_row}"] = f"=I{subtotal_row}+I{gst_row}"
        ws[f"I{grand_total_row}"].font = cls.FONT_BOLD
        ws[f"I{grand_total_row}"].number_format = "₹#,##0"
        ws[f"I{grand_total_row}"].fill = cls.TOTAL_FILL
        ws[f"I{grand_total_row}"].border = cls.BORDER_DOUBLE_BOTTOM

        # Amount in Words and GST Terms Note
        words_row = grand_total_row + 2
        subtotal_val = calc.get("subtotal_excl_gst", 0.0)
        ws[f"A{words_row}"] = f"Total Amount in Words: {number_to_indian_words(subtotal_val)}"
        ws[f"A{words_row}"].font = Font(name=cls.FONT_NAME, size=11, bold=True, italic=True)

        gst_note_row = words_row + 1
        ws[f"A{gst_note_row}"] = "All taxes and duties like GST shall be extra at actuals. The prevailing rate at the time of dispatch shall be levied. Current Rate of GST is 18%."
        ws[f"A{gst_note_row}"].font = cls.FONT_REGULAR

        # Auto-fit column widths
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    @classmethod
    def _populate_synopsis_sheet(
        cls,
        ws: openpyxl.worksheet.worksheet.Worksheet,
        quote: Dict[str, Any],
        calc: Dict[str, Any],
    ) -> None:
        """Fills the executive Synopsis sheet with component breakdowns and logistics totals."""
        ws.views.sheetView[0].showGridLines = True

        ws["A1"] = "MEGA PIPES — EXECUTIVE SYNOPSIS & COST BREAKDOWN"
        ws["A1"].font = cls.FONT_TITLE

        # Summary KPIs
        ws["A3"] = "Total Quantity (Metres):"
        ws["B3"] = calc.get("total_metres", 0.0)
        ws["A4"] = "Total Weight (Metric Tonnes):"
        ws["B4"] = calc.get("total_mt", 0.0)
        ws["A5"] = "Estimated Trailer Trips:"
        ws["B5"] = calc.get("total_trips", 0.0)
        ws["A6"] = "Trip Rate (₹ / Trip):"
        ws["B6"] = quote.get("trip_rate", 98000.0)

        for r in range(3, 7):
            ws[f"A{r}"].font = cls.FONT_BOLD
            ws[f"B{r}"].font = cls.FONT_REGULAR
            ws[f"B{r}"].number_format = "#,##0.0"

        # Cost Breakdown Table Header
        headers = [
            ("A8", "Item"),
            ("B8", "Nominal Size"),
            ("C8", "Bare Pipe (₹/m)"),
            ("D8", "External Coating (₹/m)"),
            ("E8", "Internal Coating (₹/m)"),
            ("F8", "Freight (₹/m)"),
            ("G8", "Unit Price (₹/m)"),
            ("H8", "Line Total (₹)"),
        ]
        for pos, title in headers:
            cell = ws[pos]
            cell.value = title
            cell.fill = cls.SUBHEADER_FILL
            cell.font = cls.FONT_BOLD
            cell.alignment = Alignment(horizontal="center", vertical="center")

        lines = calc.get("lines", [])
        current_row = 9

        for idx, line in enumerate(lines, start=1):
            ws[f"A{current_row}"] = idx
            ws[f"B{current_row}"] = f"{line.get('size_mm', 0):.0f} mm ({line.get('basis', 'ID')})"
            ws[f"C{current_row}"] = line.get("bare_per_m", 0.0)
            ws[f"D{current_row}"] = line.get("external_per_m", 0.0)
            ws[f"E{current_row}"] = line.get("internal_per_m", 0.0)
            ws[f"F{current_row}"] = line.get("freight_per_m", 0.0)
            ws[f"G{current_row}"] = line.get("total_per_m", 0.0)
            ws[f"H{current_row}"] = line.get("line_value", 0.0)

            ws[f"A{current_row}"].alignment = Alignment(horizontal="center")
            for c in "CDEFGH":
                ws[f"{c}{current_row}"].number_format = "₹#,##0"
                ws[f"{c}{current_row}"].font = cls.FONT_REGULAR
                ws[f"{c}{current_row}"].border = cls.BORDER_THIN
            ws[f"A{current_row}"].border = cls.BORDER_THIN
            ws[f"B{current_row}"].border = cls.BORDER_THIN

            current_row += 1

        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 14)
