"""
Genera el Reporte AP & AR (Cuentas por Pagar / Cobrar).

AP = Accounts Payable  → cuánto se debe a proveedores
AR = Accounts Receivable → cuánto deben los clientes

Flujo:
  1. Copiar template
  2. Leer reserva (folio, moneda, fecha_inicio)
  3. Leer pago_proveedor → pivots AP (total vs ya pagado)
  4. Leer pago_cliente   → pivots AR (total vs ya cobrado)
  5. Escribir Data AP y Data AR con datos + fórmulas XLOOKUP
  6. Crear Alertas Pago (próximos a vencer / ya excedidos)
  7. Actualizar FX, filtros de pivots, force refresh
"""

import os
import smtplib
import shutil
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.properties import CalcProperties

from config import AP_AR_TEMPLATE_PATH, MONTH_NAMES_ES
from src.data_loader import load_reserva
from src.weekly.fx_sheet import write_fx_sheet

# Códigos de "venta directa" por entidad (se excluyen de los pivots)
VENTA_DIRECTA = {
    "SL": {"00B"},
    "LLC": {"4E2", "E04"},
}

# Hojas a conservar en el output
_KEEP_SHEETS = {"Data AP", "Data AR", "AP Pivot", "AR Pivot"}

# Monedas para la tabla FX
_FX_CURRENCIES = ["EUR", "USD", "CHF", "GBP", "GPB", "JPY", "MXN"]


# ── Carga de CSVs ────────────────────────────────────────────────────────

def _find_csv(data_dir, base_name):
    """Busca un CSV por nombre, eligiendo el más reciente si hay variantes."""
    import glob
    exact = os.path.join(data_dir, f"{base_name}.csv")
    if os.path.exists(exact):
        candidates = [exact]
    else:
        candidates = []

    pattern = os.path.join(data_dir, f"{base_name}*.csv")
    candidates.extend(glob.glob(pattern))

    if not candidates:
        return None

    candidates = list(set(candidates))
    candidates.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    return candidates[0]


def _load_payment_csv(data_dir, filename):
    """Lee un CSV de pagos (pago_proveedor o pago_cliente)."""
    path = _find_csv(data_dir, filename)
    if not path:
        raise FileNotFoundError(f"No se encuentra {filename}.csv en {data_dir}")
    print(f"    Leyendo {os.path.basename(path)}...")
    df = pd.read_csv(path, encoding="latin-1")
    print(f"    {len(df)} registros leídos.")
    return df


# ── Construcción de pivots ───────────────────────────────────────────────

def _build_ap_pivots(pago_prov_df, vd_codes):
    """Construye los dos pivots de AP desde pago_proveedor.

    Pivot 1 (total comprometido): cancelado=0, forma_pago NOT venta directa
    Pivot 2 (ya pagado):          + fecha_aplicacion != '0000-00-00'

    Returns (pivot1_dict, pivot2_dict): {reserva: sum_monto}
    """
    base = pago_prov_df[
        (pago_prov_df["cancelado"] == 0) &
        (~pago_prov_df["forma_pago"].isin(vd_codes))
    ].copy()

    pivot1 = base.groupby("reserva")["monto"].sum().to_dict()

    paid = base[base["fecha_aplicacion"] != "0000-00-00"]
    pivot2 = paid.groupby("reserva")["monto"].sum().to_dict()

    return pivot1, pivot2


def _build_ar_pivots(pago_cli_df, vd_codes):
    """Construye los dos pivots de AR desde pago_cliente.

    Pivot 1 (total cliente):  cancelado=0, forma_pago NOT venta directa
    Pivot 2 (ya cobrado):     + fecha_proceso != '0000-00-00'

    Returns (pivot1_dict, pivot2_dict): {reserva: sum_monto}
    """
    base = pago_cli_df[
        (pago_cli_df["cancelado"] == 0) &
        (~pago_cli_df["forma_pago"].isin(vd_codes))
    ].copy()

    pivot1 = base.groupby("reserva")["monto"].sum().to_dict()

    received = base[base["fecha_proceso"] != "0000-00-00"]
    pivot2 = received.groupby("reserva")["monto"].sum().to_dict()

    return pivot1, pivot2


# ── Lookup de reserva ────────────────────────────────────────────────────

def _build_reserva_lookup(reserva_df):
    """Crea {folio: {moneda, fecha_inicio}} desde reserva filtrada."""
    lookup = {}
    for _, r in reserva_df.iterrows():
        folio = r["folio"]
        moneda = str(r.get("moneda", "EUR")).strip()
        fecha_inicio = r.get("fecha_inicio")
        if pd.notna(fecha_inicio):
            if isinstance(fecha_inicio, str):
                try:
                    fecha_inicio = pd.to_datetime(fecha_inicio)
                except Exception:
                    fecha_inicio = None
        else:
            fecha_inicio = None
        lookup[folio] = {"moneda": moneda, "fecha_inicio": fecha_inicio}
    return lookup


# ── Escritura de hojas de datos ──────────────────────────────────────────
#
# Layout unificado AP y AR (versión reparada):
#   A:  folio
#   B:  moneda           = XLOOKUP(A, Q:Q, S:S)
#   C:  total            = XLOOKUP(A, U:U, V:V)
#   D:  total EUR        = C * XLOOKUP(B, Z:Z, AA:AA)
#   E:  total USD        = C * XLOOKUP(B, Z:Z, AB:AB)
#   F:  ya pagado        = XLOOKUP(A, X:X, Y:Y)
#   G:  pagado EUR       = F * XLOOKUP(B, Z:Z, AA:AA)
#   H:  pagado USD       = F * XLOOKUP(B, Z:Z, AB:AB)
#   I:  restante         = C - F
#   J:  restante USD     = I * XLOOKUP(B, Z:Z, AB:AB)
#   K:  fecha_inicio     = XLOOKUP(A, Q:Q, R:R)
#   L:  fecha -30        = K - 30
#   M:  mes 30 días      = MONTH(L)
#   N:  año 30 días      = YEAR(L)
#
#   Q:  folio (lookup)   R: fecha_inicio   S: moneda
#   U:  folio (pivot1)   V: sum monto (total)
#   X:  folio (pivot2)   Y: sum monto (pagado/cobrado)
#   Z:  DIVISA           AA: Fx EUR        AB: Fx USD

def _write_data_sheet(ws, folios, reserva_lookup, pivot1, pivot2, fx):
    """Escribe una hoja Data AP o Data AR (layout unificado)."""
    # Limpiar datos existentes (preservar fila 1 headers)
    for row in range(2, ws.max_row + 1):
        for col in range(1, 29):  # A-AB
            ws.cell(row, col).value = None

    # A-N: folios + fórmulas XLOOKUP
    for i, folio in enumerate(folios):
        r = i + 2
        ws.cell(r, 1).value = folio                                           # A
        ws.cell(r, 2).value = f'=_xlfn.XLOOKUP(A{r},Q:Q,S:S,"not found")'    # B moneda
        ws.cell(r, 3).value = f'=_xlfn.XLOOKUP(A{r},U:U,V:V)'               # C total
        ws.cell(r, 4).value = f'=C{r}*_xlfn.XLOOKUP(B{r},Z:Z,AA:AA)'        # D EUR
        ws.cell(r, 5).value = f'=C{r}*_xlfn.XLOOKUP(B{r},Z:Z,AB:AB)'        # E USD
        ws.cell(r, 6).value = f'=_xlfn.XLOOKUP(A{r},X:X,Y:Y)'               # F pagado
        ws.cell(r, 7).value = f'=F{r}*_xlfn.XLOOKUP(B{r},Z:Z,AA:AA)'        # G EUR
        ws.cell(r, 8).value = f'=F{r}*_xlfn.XLOOKUP(B{r},Z:Z,AB:AB)'        # H USD
        ws.cell(r, 9).value = f'=C{r}-F{r}'                                  # I restante
        ws.cell(r, 10).value = f'=I{r}*_xlfn.XLOOKUP(B{r},Z:Z,AB:AB,0)'     # J rest USD
        ws.cell(r, 11).value = f'=_xlfn.XLOOKUP(A{r},Q:Q,R:R,"NO")'         # K fecha
        ws.cell(r, 12).value = f'=K{r}-30'                                   # L -30
        ws.cell(r, 13).value = f'=MONTH(L{r})'                               # M mes
        ws.cell(r, 14).value = f'=YEAR(L{r})'                                # N año

    # Q-S: lookup (folio, fecha_inicio, moneda)
    for i, folio in enumerate(folios):
        r = i + 2
        info = reserva_lookup.get(folio, {})
        ws.cell(r, 17).value = folio                           # Q
        fi = info.get("fecha_inicio")
        if fi is not None:
            ws.cell(r, 18).value = fi                          # R fecha_inicio
            ws.cell(r, 18).number_format = 'DD/MM/YYYY'
        ws.cell(r, 19).value = info.get("moneda", "EUR")      # S moneda

    # U-V: pivot 1 (total comprometido / total cliente)
    for i, folio in enumerate(folios):
        r = i + 2
        ws.cell(r, 21).value = folio                           # U
        ws.cell(r, 22).value = round(pivot1.get(folio, 0), 2)  # V

    # X-Y: pivot 2 (ya pagado / ya cobrado)
    for i, folio in enumerate(folios):
        r = i + 2
        ws.cell(r, 24).value = folio                           # X
        ws.cell(r, 25).value = round(pivot2.get(folio, 0), 2)  # Y

    # Z-AB: FX table (row 3 = header, rows 4+ = currencies)
    ws.cell(3, 26).value = "DIVISA"   # Z3
    ws.cell(3, 27).value = "Fx EUR"   # AA3
    ws.cell(3, 28).value = "Fx USD"   # AB3
    for i, curr in enumerate(_FX_CURRENCIES):
        r = 4 + i
        ws.cell(r, 26).value = curr
        if curr in fx:
            ws.cell(r, 27).value = fx[curr]["EUR"]
            ws.cell(r, 28).value = fx[curr]["USD"]


# ── Alertas de pago ─────────────────────────────────────────────────────

def _load_proveedor_lookup(data_dir):
    """Carga proveedor.csv y crea {clave: {nombre, email, ciudad}}."""
    path = _find_csv(data_dir, "proveedor")
    if not path:
        return {}
    df = pd.read_csv(path, encoding="latin-1")
    lookup = {}
    for _, r in df.iterrows():
        clave = r.get("clave")
        if pd.isna(clave):
            continue
        lookup[int(clave)] = {
            "nombre": str(r.get("nombre", "")).strip() if pd.notna(r.get("nombre")) else "",
            "email": str(r.get("correo_e_contacto", "")).strip() if pd.notna(r.get("correo_e_contacto")) else "",
            "ciudad": str(r.get("ciudad", "")).strip() if pd.notna(r.get("ciudad")) else "",
        }
    return lookup


def _build_alertas_data(pago_prov_df, prov_lookup):
    """Identifica pagos a proveedores sin aplicar, próximos a vencer o ya excedidos.

    Condiciones por fila:
    - fecha_aplicacion == '0000-00-00' (pago no aplicado)
    - monto_monedero == 0 (no pagado vía monedero)
    - fecha_limite dentro de 7 días → próximos a vencer
    - fecha_limite ya pasó → ya excedidos
    """
    today = datetime.now().date()

    unpaid = pago_prov_df[
        (pago_prov_df["fecha_aplicacion"].astype(str).str.strip() == "0000-00-00") &
        (pago_prov_df["monto_monedero"] == 0)
    ].copy()

    if unpaid.empty:
        return [], []

    unpaid["fecha_limite_dt"] = pd.to_datetime(unpaid["fecha_limite"], errors="coerce")

    proximos = []
    excedidos = []

    for _, row in unpaid.iterrows():
        fl = row["fecha_limite_dt"]
        if pd.isna(fl):
            continue
        fl_date = fl.date()
        days_remaining = (fl_date - today).days

        if days_remaining > 7:
            continue

        prov_code = row.get("proveedor", "")
        prov_info = prov_lookup.get(int(prov_code), {}) if pd.notna(prov_code) else {}

        record = {
            "reserva": row["reserva"],
            "vendedor": str(row.get("vendedor", "")),
            "fecha": str(row.get("fecha", "")),
            "proveedor_code": prov_code,
            "proveedor_nombre": prov_info.get("nombre", ""),
            "proveedor_email": prov_info.get("email", ""),
            "proveedor_ciudad": prov_info.get("ciudad", ""),
            "fecha_limite": fl_date,
            "monto": row.get("monto", 0),
        }

        if days_remaining < 0:
            excedidos.append(record)
        else:
            proximos.append(record)

    proximos.sort(key=lambda x: x["fecha_limite"])
    excedidos.sort(key=lambda x: x["fecha_limite"])

    return proximos, excedidos


def _write_alertas_sheet(wb, proximos, excedidos):
    """Crea la hoja Alertas Pago con dos tablas: próximos a vencer y ya excedidos."""
    ws = wb.create_sheet("Alertas Pago")

    title_font_white = Font(bold=True, size=13, color="FFFFFF")
    header_font_white = Font(bold=True, size=11, color="FFFFFF")
    orange_fill = PatternFill("solid", fgColor="FF8C00")
    red_fill = PatternFill("solid", fgColor="CC0000")
    header_fill = PatternFill("solid", fgColor="4472C4")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    cols = ["Reserva", "Vendedor", "Fecha", "Proveedor", "Nombre Proveedor",
            "Email Contacto", "Ciudad", "Fecha Límite", "Monto"]
    # Keys matching each column in the record dict
    keys = ["reserva", "vendedor", "fecha", "proveedor_code", "proveedor_nombre",
            "proveedor_email", "proveedor_ciudad", "fecha_limite", "monto"]

    def write_table(start_row, title, fill, records):
        # Título
        ws.cell(start_row, 1).value = title
        ws.cell(start_row, 1).font = title_font_white
        for c in range(1, len(cols) + 1):
            ws.cell(start_row, c).fill = fill
        ws.merge_cells(start_row=start_row, start_column=1,
                       end_row=start_row, end_column=len(cols))

        # Headers
        hr = start_row + 1
        for c, name in enumerate(cols, 1):
            cell = ws.cell(hr, c)
            cell.value = name
            cell.font = header_font_white
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center")

        if not records:
            ws.cell(hr + 1, 1).value = "Sin registros"
            return hr + 2

        for i, rec in enumerate(records):
            r = hr + 1 + i
            for c, key in enumerate(keys, 1):
                cell = ws.cell(r, c)
                cell.value = rec.get(key, "")
                cell.border = thin_border
                if key == "fecha_limite":
                    cell.number_format = 'DD/MM/YYYY'
                elif key == "monto":
                    cell.number_format = '#,##0.00'

        return hr + 1 + len(records)

    # Tabla 1: Próximos a vencer
    end_row = write_table(
        1, f"PROXIMOS A VENCER (7 dias o menos) - {len(proximos)} registros",
        orange_fill, proximos,
    )

    # Tabla 2: Ya excedidos (2 filas de separación)
    write_table(
        end_row + 2, f"YA EXCEDIDOS - {len(excedidos)} registros",
        red_fill, excedidos,
    )

    # Ajustar anchos de columna
    widths = {"A": 12, "B": 15, "C": 14, "D": 12, "E": 28,
              "F": 30, "G": 16, "H": 16, "I": 14}
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w

    return len(proximos), len(excedidos)


# ── Email de alertas (solo Madrid / SL) ─────────────────────────────────

_ALERTA_EMAIL_TO = "analytics.training@gannetworld.com"


def _send_alertas_email(proximos, entity_label):
    """Envía un email con la tabla de pagos próximos a vencer (solo Madrid).

    Requiere variables de entorno:
      SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS
    """
    smtp_host = os.environ.get("SMTP_HOST", "")
    smtp_port = int(os.environ.get("SMTP_PORT", "587"))
    smtp_user = os.environ.get("SMTP_USER", "")
    smtp_pass = os.environ.get("SMTP_PASS", "")

    if not all([smtp_host, smtp_user, smtp_pass]):
        print("  Email: variables SMTP no configuradas (SMTP_HOST, SMTP_USER, SMTP_PASS)")
        print("  Saltando envío de email.")
        return False

    today_str = datetime.now().strftime("%d/%m/%Y")

    # Construir tabla HTML
    rows_html = ""
    for rec in proximos:
        fl = rec["fecha_limite"]
        fl_str = fl.strftime("%d/%m/%Y") if hasattr(fl, "strftime") else str(fl)
        rows_html += f"""<tr>
            <td>{rec['reserva']}</td>
            <td>{rec['vendedor']}</td>
            <td>{rec['fecha']}</td>
            <td>{rec['proveedor_code']}</td>
            <td>{rec['proveedor_nombre']}</td>
            <td>{rec['proveedor_email']}</td>
            <td>{rec['proveedor_ciudad']}</td>
            <td>{fl_str}</td>
            <td style="text-align:right">{rec['monto']:,.2f}</td>
        </tr>"""

    html = f"""<html><body>
    <h2>Alerta Pagos Proveedores - {entity_label}</h2>
    <p>Fecha de ejecucion: {today_str}</p>
    <p>Hay <b>{len(proximos)}</b> pagos proximos a vencer (7 dias o menos):</p>
    <table border="1" cellpadding="6" cellspacing="0"
           style="border-collapse:collapse; font-family:Arial; font-size:12px;">
      <tr style="background:#4472C4; color:white;">
        <th>Reserva</th><th>Vendedor</th><th>Fecha</th>
        <th>Proveedor</th><th>Nombre</th><th>Email</th><th>Ciudad</th>
        <th>Fecha Limite</th><th>Monto</th>
      </tr>
      {rows_html}
    </table>
    <br><p style="color:gray; font-size:11px;">Generado automaticamente por Gannet Reports.</p>
    </body></html>"""

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"Alerta Pagos Proveedores {entity_label} - {today_str}"
    msg["From"] = smtp_user
    msg["To"] = _ALERTA_EMAIL_TO
    msg.attach(MIMEText(html, "html"))

    try:
        with smtplib.SMTP(smtp_host, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_pass)
            server.sendmail(smtp_user, [_ALERTA_EMAIL_TO], msg.as_string())
        print(f"  Email enviado a {_ALERTA_EMAIL_TO}")
        return True
    except Exception as e:
        print(f"  Email falló: {e}")
        return False


# ── Limpieza y pivot refresh ─────────────────────────────────────────────

def _cleanup_sheets(wb):
    """Elimina hojas innecesarias del template."""
    to_delete = [sn for sn in wb.sheetnames if sn not in _KEEP_SHEETS]
    for sn in to_delete:
        del wb[sn]


def _force_pivot_refresh(wb):
    """Fuerza el recálculo de pivots y fórmulas al abrir."""
    if wb.calculation is None:
        wb.calculation = CalcProperties()
    wb.calculation.fullCalcOnLoad = True

    seen = set()
    for ws in wb.worksheets:
        for pt in getattr(ws, '_pivots', []):
            cache = pt.cache
            cid = id(cache)
            if cid in seen:
                continue
            seen.add(cid)

            cache.refreshOnLoad = True
            cache.recordCount = 0

            try:
                if cache.records is not None:
                    cache.records.r = []
            except (AttributeError, TypeError):
                pass


# ── Función principal ────────────────────────────────────────────────────

def generate_ap_ar_report(data_dir, output_dir, fx, company, entity_label,
                          year=None, month=None):
    """
    Genera el Reporte AP & AR para una entidad.

    Args:
        data_dir: Directorio de CSVs (data/espana o data/mexico).
        output_dir: Directorio de salida.
        fx: Dict de tipos de cambio {currency: {"EUR": rate, "USD": rate}}.
        company: "SL" o "LLC".
        entity_label: "Madrid" o "Mexico" (para el nombre del archivo).
        year: Año del reporte (default: año actual).
        month: Mes del reporte (default: mes actual).

    Returns:
        Ruta del archivo generado o None si falla.
    """
    today = datetime.now()
    year = year or today.year
    month = month or today.month
    month_name = MONTH_NAMES_ES.get(month, str(month)).capitalize()

    # Verificar template
    if not os.path.exists(AP_AR_TEMPLATE_PATH):
        print(f"  ERROR: No se encuentra el template: {AP_AR_TEMPLATE_PATH}")
        return None

    vd_codes = VENTA_DIRECTA.get(company, set())
    print(f"  Entidad: {entity_label} ({company})")
    print(f"  Códigos venta directa excluidos: {vd_codes}")

    # 1. Cargar reserva (filtrada, no canceladas)
    reserva_df = load_reserva(data_dir)
    reserva_lookup = _build_reserva_lookup(reserva_df)
    print(f"  Reservas activas: {len(reserva_lookup)}")

    # 2. Cargar pagos y proveedores
    pago_prov_df = _load_payment_csv(data_dir, "pago_proveedor")
    pago_cli_df = _load_payment_csv(data_dir, "pago_cliente")
    prov_lookup = _load_proveedor_lookup(data_dir)
    print(f"  Proveedores cargados: {len(prov_lookup)}")

    # 3. Construir pivots AP
    ap_pivot1, ap_pivot2 = _build_ap_pivots(pago_prov_df, vd_codes)
    ap_folios = sorted(ap_pivot1.keys())
    print(f"  AP: {len(ap_folios)} reservas con pagos a proveedores")

    # 4. Construir pivots AR
    ar_pivot1, ar_pivot2 = _build_ar_pivots(pago_cli_df, vd_codes)
    ar_folios = sorted(ar_pivot1.keys())
    print(f"  AR: {len(ar_folios)} reservas con pagos de clientes")

    # 5. Alertas de pago
    proximos, excedidos = _build_alertas_data(pago_prov_df, prov_lookup)
    print(f"  Alertas: {len(proximos)} próximos a vencer, {len(excedidos)} ya excedidos")

    # 6. Copiar template y abrir
    filename = f"Report_AP_&_AR_{entity_label}_{month_name}_{year}.xlsx"
    filepath = os.path.join(output_dir, filename)
    os.makedirs(output_dir, exist_ok=True)
    shutil.copy2(AP_AR_TEMPLATE_PATH, filepath)

    wb = load_workbook(filepath)

    # 7. Limpiar hojas innecesarias
    _cleanup_sheets(wb)
    print(f"  Hojas: {wb.sheetnames}")

    # 8. Escribir Data AP
    ws_ap = wb["Data AP"]
    _write_data_sheet(ws_ap, ap_folios, reserva_lookup, ap_pivot1, ap_pivot2, fx)
    print(f"  Data AP: {len(ap_folios)} filas escritas")

    # 9. Escribir Data AR
    ws_ar = wb["Data AR"]
    _write_data_sheet(ws_ar, ar_folios, reserva_lookup, ar_pivot1, ar_pivot2, fx)
    print(f"  Data AR: {len(ar_folios)} filas escritas")

    # 10. Alertas Pago
    n_prox, n_exc = _write_alertas_sheet(wb, proximos, excedidos)
    print(f"  Alertas Pago: {n_prox} próximos, {n_exc} excedidos")

    # 11. FX Rates (tasas diarias del mes, misma hoja que el Weekly)
    month_label, n_days = write_fx_sheet(wb)
    print(f"  FX Rates: {n_days} días de {month_label}")

    # 12. Force refresh (pivots + fórmulas)
    _force_pivot_refresh(wb)
    print(f"  Force refresh configurado")

    # 13. Guardar
    wb.save(filepath)
    wb.close()

    # 14. Email de alertas (solo Madrid/SL, solo si hay próximos)
    if company == "SL" and proximos:
        _send_alertas_email(proximos, entity_label)

    return filepath
