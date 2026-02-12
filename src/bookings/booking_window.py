"""
Módulo Booking Window: genera la matriz semana x mes y exporta a Excel.
"""

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from config import MONTH_NAMES_ES

# USD format matching the template
USD_FMT = '[$$-409]#,##0'


def _week_to_row(week_num):
    """Convert week number to row in Booking Window sheet.
    Week 53 is row 2, Week 52 is row 4, ..., Week 1 is row 106."""
    return 2 + (53 - week_num) * 2


def build_booking_matrix(data_rows):
    """Build the booking window matrix: for each sale week, sum Total Venta USD
    by departure month.

    Args:
        data_rows: List of DATA sheet row dicts.

    Returns:
        Dict {week: {month_key: total_usd}} where month_key 1-12 = 2026,
        14-25 = 2027 months.
    """
    matrix = {}

    for r in data_rows:
        fecha_inicio = r["E"]  # departure date
        total_usd = r["M"]    # Total Venta USD
        fecha = r["D"]         # sale date

        if not fecha or not fecha_inicio:
            continue

        sale_week = fecha.isocalendar()[1]
        dep_year = fecha_inicio.year
        dep_month = fecha_inicio.month

        if sale_week not in matrix:
            matrix[sale_week] = {}

        if dep_year == 2026:
            matrix[sale_week][dep_month] = (
                matrix[sale_week].get(dep_month, 0) + total_usd
            )
        elif dep_year == 2027:
            matrix[sale_week][dep_month + 13] = (
                matrix[sale_week].get(dep_month + 13, 0) + total_usd
            )

    return matrix


def write_booking_to_excel(wb, matrix):
    """Write the booking window matrix into the 'Booking Window 2026' sheet.

    Args:
        wb: Open openpyxl Workbook.
        matrix: Booking matrix from build_booking_matrix().
    """
    ws_bw = wb["Booking Window 2026"]

    for week_num, month_totals in matrix.items():
        if week_num < 1 or week_num > 53:
            continue

        value_row = _week_to_row(week_num)
        pct_row = value_row + 1

        total_2026 = 0
        total_2027 = 0

        # Write 2026 months (cols C-N = 3-14)
        for month in range(1, 13):
            val = month_totals.get(month, 0)
            if val:
                col_idx = month + 2
                cell = ws_bw.cell(value_row, col_idx)
                cell.value = round(val, 2)
                cell.number_format = USD_FMT
                total_2026 += val

        # Write total 2026 (col O = 15)
        if total_2026:
            cell = ws_bw.cell(value_row, 15)
            cell.value = round(total_2026, 2)
            cell.number_format = USD_FMT

        # Write 2027 months (cols P-AA = 16-27)
        for month_key in range(14, 26):
            val = month_totals.get(month_key, 0)
            if val:
                month_2027 = month_key - 13
                col_idx = month_2027 + 15
                cell = ws_bw.cell(value_row, col_idx)
                cell.value = round(val, 2)
                cell.number_format = USD_FMT
                total_2027 += val

        # Write total 2027 (col AB = 28)
        if total_2027:
            cell = ws_bw.cell(value_row, 28)
            cell.value = round(total_2027, 2)
            cell.number_format = USD_FMT

        # Write percentage formulas in the row below
        if total_2026:
            for col in range(3, 15):
                cell = ws_bw.cell(pct_row, col)
                cell.value = (
                    f"={get_column_letter(col)}{value_row}/$O${value_row}"
                )
                cell.number_format = '0%'


def export_booking_xlsx(matrix, output_path):
    """Export the booking window matrix as a standalone Excel file.

    Args:
        matrix: Booking matrix from build_booking_matrix().
        output_path: Path for the output .xlsx file.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Booking Window"

    # Header
    months_2026 = [MONTH_NAMES_ES[m].upper()[:3] + " 2026" for m in range(1, 13)]
    months_2027 = [MONTH_NAMES_ES[m].upper()[:3] + " 2027" for m in range(1, 13)]
    header = ["SEMANA"] + months_2026 + ["TOTAL 2026"] + months_2027 + ["TOTAL 2027"]
    ws.append(header)

    # Bold header
    for cell in ws[1]:
        cell.font = cell.font.copy(bold=True)

    # Data rows (sorted by week descending, like the template)
    for week_num in sorted(matrix.keys(), reverse=True):
        month_totals = matrix[week_num]
        row = [f"WEEK {week_num}"]

        total_2026 = 0
        for month in range(1, 13):
            val = round(month_totals.get(month, 0), 2)
            row.append(val)
            total_2026 += val
        row.append(round(total_2026, 2))

        total_2027 = 0
        for month_key in range(14, 26):
            val = round(month_totals.get(month_key, 0), 2)
            row.append(val)
            total_2027 += val
        row.append(round(total_2027, 2))

        ws.append(row)

    # Format USD columns (B onwards)
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=ws.max_column):
        for cell in row:
            cell.number_format = USD_FMT

    wb.save(output_path)
    wb.close()
    print(f"  Booking Window Excel exportado: {output_path}")
