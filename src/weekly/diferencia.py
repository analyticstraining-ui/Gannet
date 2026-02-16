"""
Calcula y escribe la columna "Diferencia %" en la hoja "Weekly SL y LLC".

Calcula directamente desde data_rows (no desde celdas cacheadas del pivot).
Fórmula por semana: (Venta 2026 USD - Venta 2025 USD) / Venta 2025 USD
Total general: promedio de las diferencias semanales.
"""

from collections import defaultdict
from datetime import datetime

from openpyxl.styles import Font, PatternFill


def write_diferencia_pct(wb, data_rows):
    """
    Escribe la columna Diferencia % en la hoja "Weekly SL y LLC".

    Calcula los totales por semana directamente desde data_rows para evitar
    usar valores cacheados del pivot (que no se actualizan hasta abrir Excel).

    Args:
        wb: Workbook abierto de openpyxl.
        data_rows: Lista de dicts con claves A-Z (salida de build_data_rows).

    Returns:
        Número de semanas procesadas, o 0 si la hoja no existe.
    """
    if "Weekly SL y LLC" not in wb.sheetnames:
        return 0

    ws = wb["Weekly SL y LLC"]

    # Encontrar fila de headers ("Semana" en col A)
    hdr_row = None
    for row in range(1, ws.max_row + 1):
        if str(ws.cell(row, 1).value or "").strip() == "Semana":
            hdr_row = row
            break

    if not hdr_row:
        return 0

    # Determinar años del reporte (leer filtros del pivot)
    current_year = None
    prev_year = None
    for row in range(1, hdr_row):
        a_val = str(ws.cell(row, 1).value or "").strip()
        e_val = str(ws.cell(row, 5).value or "").strip()
        if a_val == "Año":
            current_year = ws.cell(row, 2).value
        if e_val == "Año":
            prev_year = ws.cell(row, 6).value

    if not current_year:
        current_year = datetime.now().year
    if not prev_year:
        prev_year = current_year - 1

    # Calcular Venta USD por semana desde data_rows
    venta_current = defaultdict(float)  # {semana: total_venta_usd}
    venta_prev = defaultdict(float)

    for r in data_rows:
        año = r.get("S")
        semana = r.get("H")
        venta_usd = float(r.get("M") or 0)

        if semana is None:
            continue

        if año == current_year:
            venta_current[semana] += venta_usd
        elif año == prev_year:
            venta_prev[semana] += venta_usd

    # Header "Diferencia %" en la misma fila que los demás headers
    ws.cell(hdr_row, 9).value = "Diferencia %"
    ws.cell(hdr_row, 9).font = Font(bold=True)

    # Encontrar fila de Total general
    total_row = None
    for row in range(hdr_row + 1, ws.max_row + 1):
        if str(ws.cell(row, 1).value or "").strip() == "Total general":
            total_row = row
            break

    # Escribir Diferencia % para cada semana
    dif_values = []
    for row in range(hdr_row + 1, total_row or ws.max_row + 1):
        semana = ws.cell(row, 1).value
        if semana is None:
            continue

        cur = venta_current.get(semana, 0)
        prev = venta_prev.get(semana, 0)

        if prev != 0:
            dif = (cur - prev) / prev
        else:
            dif = 0

        dif_values.append(dif)
        cell = ws.cell(row, 9, dif)
        cell.number_format = '0%'

        print(f"  Sem {semana}: Venta2026={cur:,.2f}, Venta2025={prev:,.2f} → Dif={dif:.1%}")

    # Total general: PROMEDIO de las diferencias, con fondo azul
    if dif_values and total_row:
        avg = sum(dif_values) / len(dif_values)
        cell_total = ws.cell(total_row, 9, avg)
        cell_total.number_format = '0.0%'
        cell_total.font = Font(bold=True)
        cell_total.fill = PatternFill(
            start_color="BDD7EE", end_color="BDD7EE", fill_type="solid"
        )

    return len(dif_values)
