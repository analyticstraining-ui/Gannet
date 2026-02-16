"""
Genera el Reporte Mensual Consolidado de Travel Advisors.

Usa un template con pivot tables pre-configurados (igual que el Weekly Report).
El template debe crearse una vez en Excel usando tools/create_ta_seed.py.

Flujo:
  1. Copiar template
  2. Escribir datos enriquecidos en hoja DATA NEW
  3. Actualizar pivot source ranges
  4. refreshOnLoad = True → pivots se refrescan al abrir en Excel
"""

import os
import csv
import shutil
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config import BASE_DIR, MONTH_NAMES_ES, TA_TEMPLATE_PATH

# ── Estilos ──────────────────────────────────────────────────────────────
FONT = Font(name="Aptos Narrow", size=11)
FONT_HDR = Font(name="Aptos Narrow", size=11, bold=True, color="FFFFFF")

THIN = Side(style="thin", color="BFBFBF")
MEDIUM = Side(style="medium", color="808080")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_BORDER = Border(left=THIN, right=THIN, top=MEDIUM, bottom=MEDIUM)

DATA_NEW_HDR = PatternFill("solid", fgColor="4472C4")
HDR_FILL = PatternFill("solid", fgColor="548235")

PCT_FMT = '0.0%'
DATE_FMT = 'DD/MM/YYYY'
NUM_FMT = '#,##0.00'

# Columnas que usan formulas XLOOKUP en vez de valores pre-computados
FORMULA_KEYS = {"L", "M", "O", "P", "Q", "AB_renta_com", "AC_renta_com_usd"}

# FX lookup table position: cols AO-AQ (41-43)
FX_COL_START = 41  # AO

# Definicion de columnas DATA NEW
DATA_NEW_COLS = [
    ("A", "Compania"), ("B", "folio"), ("C", "cerrada"),
    ("D", "fecha"), ("E", "fecha_inicio"), ("F", "fecha_fin"),
    ("G", "vendedor"), ("H", "Semana"), ("I", "usuarios_invitados"),
    ("J", "total_cliente"), ("K", "moneda"),
    ("L", "Total Venta EUR"), ("M", "Total Venta USD"),
    ("N", "Rentabilidad"), ("O", "Rentabilidad en EUR"),
    ("P", "Rentabilidad en USD"), ("Q", "% Rentabilidad"),
    ("R", "Mes"), ("S", "Ano"), ("T", "Mes Inicio"), ("U", "Ano Inicio"),
    ("V", "Fecha 45 dias fin"), ("W", "mes 45 dias fin"),
    ("X", "Ano 45 dias fin"),
    ("Y_oficina", "Oficina"), ("Z_linea", "Linea de Negocio"),
    ("AA_com", "Comisionamiento"),
    ("AB_renta_com", "Renta after COM"),
    ("AC_renta_com_usd", "Renta after COM USD"),
    ("AD_fecha_inc", "Fecha Incorporacion"),
]

DATA_NEW_WIDTHS = [10, 8, 7, 12, 12, 12, 14, 8, 10, 14, 8,
                   16, 16, 14, 16, 16, 14, 6, 6, 10, 10,
                   16, 14, 10, 12, 16, 16, 16, 18, 16]


# ── Helpers ──────────────────────────────────────────────────────────────

def _cell(ws, row, col, val, font=FONT, fmt=None, fill=None, border=None,
          indent=0, alignment=None):
    """Escribe un valor con formato completo."""
    cell = ws.cell(row, col, val)
    cell.font = font
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = fill
    if border:
        cell.border = border
    if indent:
        cell.alignment = Alignment(indent=indent, horizontal="left")
    elif alignment:
        cell.alignment = alignment
    return cell


# ── Plantilla Corsario ───────────────────────────────────────────────────

def load_plantilla_corsario():
    """Carga la Plantilla Corsario desde el CSV en templates/."""
    path = os.path.join(BASE_DIR, "templates", "plantilla_corsario.csv")
    plantilla = []
    with open(path, encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            com = row["comisionamiento"]
            try:
                com = float(com) if com else 0
            except ValueError:
                com = 0
            fecha = None
            if row["fecha_inicio"]:
                try:
                    fecha = datetime.strptime(row["fecha_inicio"], "%Y-%m-%d")
                except ValueError:
                    pass
            plantilla.append({
                "usuario": row["usuario_corsario"].strip(),
                "linea_negocio": row["linea_negocio"].strip(),
                "comisionamiento": com,
                "oficina": row["pais_trabajo"].strip(),
                "fecha_inicio": fecha,
            })
    return plantilla


def _build_lookup(plantilla):
    """Crea dict {vendedor: {oficina, linea_negocio, comisionamiento, fecha_inicio}}."""
    return {p["usuario"]: p for p in plantilla}


# ── Enriquecer data_rows ─────────────────────────────────────────────────

def enrich_data_rows(data_rows, plantilla):
    """Agrega columnas Y-AD a cada row basado en la plantilla.

    Nota: L, M, O, P, Q, AB, AC se escriben como formulas XLOOKUP
    en la hoja, no como valores pre-computados.
    """
    lookup = _build_lookup(plantilla)
    enriched = []
    for r in data_rows:
        row = dict(r)  # copia
        vendedor = str(row.get("G") or "").strip()
        info = lookup.get(vendedor, {})

        row["Y_oficina"] = info.get("oficina", "")
        row["Z_linea"] = info.get("linea_negocio", "")
        row["AA_com"] = info.get("comisionamiento", 0)
        row["AD_fecha_inc"] = info.get("fecha_inicio")
        enriched.append(row)
    return enriched


# ── Plantilla Corsario en hoja ───────────────────────────────────────────

def _write_plantilla_block(ws, plantilla, col_start=35, row_start=1):
    """Escribe la plantilla corsario en cols AI-AM (35-39) desde row_start."""
    headers = ["usuario Corsario", "Linea de Negocio", "Comisionamiento",
               "Pais de trabajo", "Fecha Inicio"]
    for i, h in enumerate(headers):
        _cell(ws, row_start, col_start + i, h,
              font=FONT_HDR, fill=HDR_FILL, border=HDR_BORDER)

    for j, p in enumerate(plantilla):
        r = row_start + 1 + j
        _cell(ws, r, col_start, p["usuario"], border=THIN_BORDER)
        _cell(ws, r, col_start + 1, p["linea_negocio"], border=THIN_BORDER)
        _cell(ws, r, col_start + 2, p["comisionamiento"], fmt=PCT_FMT, border=THIN_BORDER)
        _cell(ws, r, col_start + 3, p["oficina"], border=THIN_BORDER)
        if p["fecha_inicio"]:
            _cell(ws, r, col_start + 4, p["fecha_inicio"], fmt=DATE_FMT, border=THIN_BORDER)
        else:
            _cell(ws, r, col_start + 4, None, border=THIN_BORDER)

    ws.column_dimensions[get_column_letter(col_start)].width = 15
    ws.column_dimensions[get_column_letter(col_start + 1)].width = 15.5
    ws.column_dimensions[get_column_letter(col_start + 2)].width = 16.5
    ws.column_dimensions[get_column_letter(col_start + 3)].width = 13.5
    ws.column_dimensions[get_column_letter(col_start + 4)].width = 11.5


# ── FX lookup table ───────────────────────────────────────────────────

def _write_fx_table(ws, fx):
    """Escribe la tabla de tipos de cambio en cols AO-AQ (41-43)."""
    _cell(ws, 1, FX_COL_START, "Moneda",
          font=FONT_HDR, fill=HDR_FILL, border=HDR_BORDER)
    _cell(ws, 1, FX_COL_START + 1, "FX EUR",
          font=FONT_HDR, fill=HDR_FILL, border=HDR_BORDER)
    _cell(ws, 1, FX_COL_START + 2, "FX USD",
          font=FONT_HDR, fill=HDR_FILL, border=HDR_BORDER)

    currencies = ["EUR", "USD", "CHF", "GBP", "GPB", "JPY", "MXN"]
    for i, curr in enumerate(currencies):
        r = 2 + i
        _cell(ws, r, FX_COL_START, curr, border=THIN_BORDER)
        if curr in fx:
            _cell(ws, r, FX_COL_START + 1, fx[curr]["EUR"],
                  fmt='0.000000', border=THIN_BORDER)
            _cell(ws, r, FX_COL_START + 2, fx[curr]["USD"],
                  fmt='0.000000', border=THIN_BORDER)

    ws.column_dimensions[get_column_letter(FX_COL_START)].width = 10
    ws.column_dimensions[get_column_letter(FX_COL_START + 1)].width = 12
    ws.column_dimensions[get_column_letter(FX_COL_START + 2)].width = 12


# ── Formulas XLOOKUP ─────────────────────────────────────────────────

def _get_formula(key, row_num):
    """Retorna la formula XLOOKUP/calculada para una columna."""
    if key == "L":       # Total Venta EUR = total_cliente * FX EUR
        return f'=J{row_num}*_xlfn.XLOOKUP(K{row_num},AO:AO,AP:AP)'
    elif key == "M":     # Total Venta USD = total_cliente * FX USD
        return f'=J{row_num}*_xlfn.XLOOKUP(K{row_num},AO:AO,AQ:AQ)'
    elif key == "O":     # Rentabilidad en EUR = Rentabilidad * FX EUR
        return f'=N{row_num}*_xlfn.XLOOKUP(K{row_num},AO:AO,AP:AP)'
    elif key == "P":     # Rentabilidad en USD = Rentabilidad * FX USD
        return f'=N{row_num}*_xlfn.XLOOKUP(K{row_num},AO:AO,AQ:AQ)'
    elif key == "Q":     # % Rentabilidad = Rentabilidad / total_cliente
        return f'=IFERROR(N{row_num}/J{row_num},0)'
    elif key == "AB_renta_com":      # Renta after COM = Rent EUR * (1 - COM)
        return f'=O{row_num}*(1-AA{row_num})'
    elif key == "AC_renta_com_usd":  # Renta after COM USD = Rent USD * (1 - COM)
        return f'=P{row_num}*(1-AA{row_num})'
    return None


# ── Escribir DATA NEW ───────────────────────────────────────────────────

def _write_data_new(ws, enriched, plantilla, fx):
    """Escribe datos enriquecidos en DATA NEW con formulas XLOOKUP para FX."""

    # Limpiar datos existentes (preservar estructura)
    for row in range(2, ws.max_row + 1):
        for col in range(1, len(DATA_NEW_COLS) + 1):
            ws.cell(row, col).value = None

    # Verificar/escribir headers
    for ci, (_, header) in enumerate(DATA_NEW_COLS, 1):
        _cell(ws, 1, ci, header, font=FONT_HDR, fill=DATA_NEW_HDR, border=HDR_BORDER)

    # Escribir datos con formulas XLOOKUP
    for ri, row_data in enumerate(enriched, 2):
        for ci, (key, _) in enumerate(DATA_NEW_COLS, 1):
            if key in FORMULA_KEYS:
                # Columnas calculadas: escribir formula
                formula = _get_formula(key, ri)
                cell = ws.cell(ri, ci, formula)
                cell.font = FONT
                cell.border = THIN_BORDER
                if key == "Q":
                    cell.number_format = PCT_FMT
                else:
                    cell.number_format = NUM_FMT
            else:
                # Columnas de datos: escribir valor
                val = row_data.get(key)
                if val is not None:
                    fmt = None
                    if key in ("D", "E", "F", "V", "AD_fecha_inc"):
                        if isinstance(val, datetime):
                            fmt = DATE_FMT
                    elif key in ("J", "N"):
                        fmt = NUM_FMT
                    elif key == "AA_com":
                        fmt = PCT_FMT
                    _cell(ws, ri, ci, val, fmt=fmt, border=THIN_BORDER)
                else:
                    _cell(ws, ri, ci, None, border=THIN_BORDER)

    # Anchos
    for i, w in enumerate(DATA_NEW_WIDTHS):
        ws.column_dimensions[get_column_letter(i + 1)].width = w

    # Plantilla Corsario en cols AI-AM (35-39)
    _write_plantilla_block(ws, plantilla, col_start=35, row_start=1)

    # FX lookup table en cols AO-AQ (41-43)
    _write_fx_table(ws, fx)


# ── Actualizar pivots ───────────────────────────────────────────────────

def _update_pivots(wb, n_data_rows):
    """Actualiza source ranges, fuerza DATA NEW como fuente, y limpia cache."""
    n_cols = len(DATA_NEW_COLS)
    last_col = get_column_letter(n_cols)
    new_ref = f"A1:{last_col}{n_data_rows + 1}"

    pivot_count = 0
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if hasattr(ws, '_pivots') and ws._pivots:
            for pivot in ws._pivots:
                if pivot.cache and pivot.cache.cacheSource:
                    src = pivot.cache.cacheSource
                    if src.worksheetSource:
                        # Forzar TODOS los pivots a DATA NEW (no solo los vacios)
                        src.worksheetSource.sheet = "DATA NEW"
                        src.worksheetSource.ref = new_ref
                # Limpiar cache viejo para forzar recalculo
                if pivot.cache:
                    pivot.cache.refreshOnLoad = True
                    pivot.cache.records = None
                pivot_count += 1

    return pivot_count


# ── Funcion principal ────────────────────────────────────────────────────

def generate_ta_monthly_report(data_rows, output_dir, fx, year=None, month=None):
    """
    Genera el Reporte Mensual Consolidado de TAs.

    Copia el template con pivots pre-configurados, escribe DATA NEW
    con datos enriquecidos y formulas XLOOKUP para conversion FX.

    Args:
        data_rows: Lista de dicts con claves A-Z (salida de build_data_rows).
        output_dir: Directorio de salida.
        fx: Dict de tipos de cambio {currency: {"EUR": rate, "USD": rate}}.
        year: Ano del reporte (default: ano actual).
        month: Mes del reporte (default: mes actual).

    Returns:
        Ruta del archivo generado.
    """
    today = datetime.now()
    year = year or today.year
    month = month or today.month
    month_name = MONTH_NAMES_ES.get(month, str(month)).capitalize()

    # Verificar template
    if not os.path.exists(TA_TEMPLATE_PATH):
        print(f"  ERROR: No se encuentra el template: {TA_TEMPLATE_PATH}")
        print(f"  Ejecuta 'python3 tools/create_ta_seed.py' y sigue las instrucciones.")
        return None

    # Cargar plantilla y enriquecer datos
    plantilla = load_plantilla_corsario()
    enriched = enrich_data_rows(data_rows, plantilla)

    print(f"  Plantilla Corsario: {len(plantilla)} TAs")
    print(f"  Datos enriquecidos: {len(enriched)} filas")

    # Copiar template
    filename = f"Reporte_TAs_{month_name}_{year}.xlsx"
    filepath = os.path.join(output_dir, filename)
    os.makedirs(output_dir, exist_ok=True)
    shutil.copy2(TA_TEMPLATE_PATH, filepath)

    # Abrir y escribir datos
    wb = load_workbook(filepath)

    if "DATA NEW" not in wb.sheetnames:
        print(f"  ERROR: El template no tiene hoja 'DATA NEW'")
        wb.close()
        return None

    ws = wb["DATA NEW"]
    _write_data_new(ws, enriched, plantilla, fx)
    print(f"  DATA NEW: {len(enriched)} filas escritas (con formulas XLOOKUP)")

    # Actualizar pivots
    n_pivots = _update_pivots(wb, len(enriched))
    print(f"  Pivot tables actualizados: {n_pivots} (refreshOnLoad=True)")

    # Guardar
    wb.save(filepath)
    wb.close()

    return filepath
