#!/usr/bin/env python3
"""
Generador automático de reportes Gannet.

Uso:
    python3 main.py                              # todos los reportes
    python3 main.py --weekly                     # solo Weekly Report
    python3 main.py --bookings                   # solo Bookings por entidad
    python3 main.py --dashboard                  # solo Dashboard
    python3 main.py --individual                 # solo Reportes individuales por TA
    python3 main.py --ta-monthly                 # solo Reporte mensual consolidado TAs
    python3 main.py --ap-ar                      # solo Reporte AP & AR
    python3 main.py --weekly --dashboard         # combinaciones
    python3 main.py --week 7                     # semana específica
    python3 main.py --entity espana              # solo España
    python3 main.py --entity mexico              # solo México
    python3 main.py --week 7 --entity mexico     # combinado
"""

import argparse
import os
import sys
from datetime import datetime

from config import ENTITIES, TEMPLATE_PATH, BASE_DIR
from src.data_loader import load_reserva, load_dreserva, compute_rentabilidad
from src.fx_rates import get_latest_fx, preload_fx_months
from src.validators import detect_errors
from src.weekly import (
    build_data_rows, build_serv_rows, generate_weekly_excel,
    write_diferencia_pct, write_fx_sheet, write_errores_sheet,
)
from src.bookings import build_booking_matrix, write_booking_to_excel, export_bookings_xlsx
from src.dashboard import generate_dashboard
from src.individual import generate_individual_reports
from src.ta_monthly import generate_ta_monthly_report
from src.ap_ar import generate_ap_ar_report

REPORT_FLAGS = ["weekly", "bookings", "dashboard", "individual", "ta_monthly", "ap_ar"]


def _load_data(entities, today, week_num):
    """Carga datos de todas las entidades y retorna las estructuras comunes."""
    all_data_rows = []
    all_serv_rows = []
    all_errors = []
    total_serv_count = 0
    entity_info = {}

    for key, cfg in entities.items():
        company = cfg["company"]
        label = cfg["label"]
        data_dir = cfg["data_dir"]

        if not os.path.exists(data_dir):
            print(f"\n  Directorio de datos no encontrado para {label}: {data_dir}")
            print(f"  Saltando {label}...")
            continue

        print(f"\n{'─' * 60}")
        print(f"  Cargando {label} ({company})")
        print(f"{'─' * 60}")

        print(f"  Leyendo datos...")
        reserva_df = load_reserva(data_dir)
        dreserva_df = load_dreserva(data_dir)
        rentabilidad_df = compute_rentabilidad(dreserva_df)

        print(f"  Descargando tasas históricas por fecha de reserva...")
        preload_fx_months(reserva_df["fecha"].dropna())

        data_rows = build_data_rows(reserva_df, rentabilidad_df, company)
        serv_rows, serv_count = build_serv_rows(dreserva_df, reserva_df, company)
        print(f"  DATA: {len(data_rows)} filas | DATA SERV: {serv_count} filas")

        # Bookings por entidad (siempre se generan al cargar datos)
        output_dir_entity = cfg["output_dir"]
        os.makedirs(output_dir_entity, exist_ok=True)
        bookings_path = os.path.join(
            output_dir_entity, f"Bookings_{company}_{week_num}_{today.year}.xlsx"
        )
        export_bookings_xlsx(data_rows, bookings_path, company)

        # Validar
        errors = detect_errors(reserva_df, rentabilidad_df)
        for err in errors:
            err["Compañia"] = company
            all_errors.append(err)
        if errors:
            print(f"\n  {len(errors)} posibles errores detectados")
        else:
            print(f"  Sin errores detectados.")

        all_data_rows.extend(data_rows)
        all_serv_rows.extend(serv_rows)
        total_serv_count += serv_count

        entity_info[key] = {
            "label": label, "company": company,
            "n_data": len(data_rows), "n_serv": serv_count,
            "n_errors": len(errors),
        }

    return all_data_rows, all_serv_rows, all_errors, total_serv_count, entity_info


def _run_weekly(all_data_rows, all_serv_rows, all_errors, total_serv_count,
                fx_latest, output_dir, output_xlsx, week_num):
    """Genera el Weekly Report combinado."""
    print(f"\n{'─' * 60}")
    print(f"  Weekly Report — Week {week_num}")
    print(f"{'─' * 60}")

    if not os.path.exists(TEMPLATE_PATH):
        print(f"  ERROR: No se encuentra la plantilla: {TEMPLATE_PATH}")
        return

    print(f"  Total: DATA={len(all_data_rows)} filas, DATA SERV={total_serv_count} filas")

    wb = generate_weekly_excel(
        TEMPLATE_PATH, output_xlsx, all_data_rows, all_serv_rows,
        total_serv_count, fx_latest
    )

    combined_matrix = build_booking_matrix(all_data_rows)
    print(f"  Booking Window: semanas con datos: {sorted(combined_matrix.keys())}")
    write_booking_to_excel(wb, combined_matrix)

    n_weeks = write_diferencia_pct(wb, all_data_rows)
    print(f"  Weekly SL y LLC: Diferencia % ({n_weeks} semanas)")

    month_label, n_days = write_fx_sheet(wb)
    print(f"  Hoja FX RATES: {n_days} días de {month_label}")

    n_errors = write_errores_sheet(wb, all_errors)
    if n_errors:
        print(f"  Hoja ERRORES: {n_errors} posibles errores")

    print(f"  Guardando {output_xlsx}...")
    wb.save(output_xlsx)
    wb.close()
    print(f"  Weekly Report generado exitosamente.")


def _run_bookings(all_data_rows, output_dir, week_num, today):
    """Genera el archivo de Bookings combinado."""
    print(f"\n{'─' * 60}")
    print(f"  Bookings combinado")
    print(f"{'─' * 60}")

    bookings_path = os.path.join(output_dir, f"Bookings_ALL_{week_num}_{today.year}.xlsx")
    combined_matrix = build_booking_matrix(all_data_rows)
    from openpyxl import Workbook
    from src.bookings import export_bookings_xlsx as _export
    _export(all_data_rows, bookings_path, "ALL")
    print(f"  Bookings combinado: {bookings_path}")


def _run_dashboard(all_data_rows, all_serv_rows, fx_latest, output_dir, week_num, today):
    """Genera el Dashboard."""
    print(f"\n{'─' * 60}")
    print(f"  Dashboard")
    print(f"{'─' * 60}")

    dashboard_path = os.path.join(output_dir, f"Dashboard_{week_num}_{today.year}.xlsx")
    generate_dashboard(all_data_rows, all_serv_rows, fx_latest, dashboard_path)
    print(f"  Dashboard generado: {dashboard_path}")


def _run_individual(all_data_rows, output_dir, today, report_month, report_year):
    """Genera los reportes individuales por TA."""
    print(f"\n{'─' * 60}")
    print(f"  Reportes Individuales por TA")
    print(f"{'─' * 60}")

    generate_individual_reports(all_data_rows, output_dir, year=report_year, month=report_month)


def _run_ta_monthly(all_data_rows, output_dir, fx_latest, today):
    """Genera el Reporte Mensual Consolidado de TAs."""
    print(f"\n{'─' * 60}")
    print(f"  Reporte Mensual Consolidado de TAs")
    print(f"{'─' * 60}")

    ta_report_path = generate_ta_monthly_report(
        all_data_rows, output_dir, fx_latest, year=today.year, month=today.month
    )
    if ta_report_path:
        print(f"  Reporte generado: {ta_report_path}")
    else:
        print(f"  Saltando reporte (template no encontrado)")


_AP_AR_ENTITY_LABELS = {"SL": "Madrid", "LLC": "Mexico"}


def _run_ap_ar(entities, output_dir, fx_latest, today):
    """Genera el Reporte AP & AR para cada entidad."""
    print(f"\n{'─' * 60}")
    print(f"  Reporte AP & AR (Cuentas por Pagar / Cobrar)")
    print(f"{'─' * 60}")

    for key, cfg in entities.items():
        company = cfg["company"]
        data_dir = cfg["data_dir"]
        entity_label = _AP_AR_ENTITY_LABELS.get(company, cfg["label"])

        report_path = generate_ap_ar_report(
            data_dir=data_dir,
            output_dir=output_dir,
            fx=fx_latest,
            company=company,
            entity_label=entity_label,
            year=today.year,
            month=today.month,
        )
        if report_path:
            print(f"  Reporte generado: {report_path}")
        else:
            print(f"  Saltando {entity_label} (template no encontrado)")


def main():
    parser = argparse.ArgumentParser(
        description="Genera reportes de Gannet",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  python3 main.py                              # todos los reportes
  python3 main.py --weekly                     # solo Weekly Report
  python3 main.py --dashboard                  # solo Dashboard
  python3 main.py --weekly --dashboard         # Weekly + Dashboard
  python3 main.py --ta-monthly                 # solo Reporte mensual TAs
  python3 main.py --ap-ar                      # solo Reporte AP & AR
  python3 main.py --individual --report-month 1  # Individuales de enero
  python3 main.py --week 7 --entity mexico     # semana y entidad específica
        """
    )

    # ── Flags de selección de reportes ────────────────────────────────
    report_group = parser.add_argument_group("Reportes (sin flags = todos)")
    report_group.add_argument(
        "--weekly", action="store_true",
        help="Weekly Report (Week_N_YYYY.xlsx)"
    )
    report_group.add_argument(
        "--bookings", action="store_true",
        help="Bookings combinado"
    )
    report_group.add_argument(
        "--dashboard", action="store_true",
        help="Dashboard con gráficos (Dashboard_N_YYYY.xlsx)"
    )
    report_group.add_argument(
        "--individual", action="store_true",
        help="Reportes individuales por TA"
    )
    report_group.add_argument(
        "--ta-monthly", action="store_true", dest="ta_monthly",
        help="Reporte mensual consolidado de TAs"
    )
    report_group.add_argument(
        "--ap-ar", action="store_true", dest="ap_ar",
        help="Reporte AP & AR (Cuentas por Pagar / Cobrar)"
    )

    # ── Parámetros generales ──────────────────────────────────────────
    parser.add_argument(
        "--week", type=int, default=None,
        help="Número de semana (default: semana actual)"
    )
    parser.add_argument(
        "--entity", type=str, default="all",
        choices=["espana", "mexico", "all"],
        help="Entidad a procesar (default: all)"
    )
    parser.add_argument(
        "--report-month", type=int, default=None, dest="report_month",
        help="Mes para reportes individuales (default: mes anterior)"
    )

    args = parser.parse_args()

    # Si no se pasó ningún flag de reporte → ejecutar todos
    run_all = not any(getattr(args, f) for f in REPORT_FLAGS)
    run = {f: (run_all or getattr(args, f)) for f in REPORT_FLAGS}

    today = datetime.now()
    week_num = args.week or int(today.strftime('%W'))

    selected = [f for f in REPORT_FLAGS if run[f]]
    print(f"═══ Gannet Reports - Semana {week_num} ═══")
    if not run_all:
        print(f"  Reportes seleccionados: {', '.join(selected)}")

    # ── Entidades ─────────────────────────────────────────────────────
    if args.entity == "all":
        entities = ENTITIES
    else:
        entities = {args.entity: ENTITIES[args.entity]}

    # ── Paso 1: FX rates ──────────────────────────────────────────────
    print("\n[1] Obteniendo tipos de cambio...")
    fx_latest = get_latest_fx()

    # ── Paso 2: Cargar datos ──────────────────────────────────────────
    print("\n[2] Cargando datos por entidad...")
    all_data_rows, all_serv_rows, all_errors, total_serv_count, entity_info = \
        _load_data(entities, today, week_num)

    if not all_data_rows:
        print("\nERROR: No se cargaron datos de ninguna entidad.")
        sys.exit(1)

    output_dir = os.path.join(BASE_DIR, "output")
    os.makedirs(output_dir, exist_ok=True)

    # ── Paso 3: Ejecutar reportes seleccionados ───────────────────────

    if run["weekly"]:
        output_xlsx = os.path.join(output_dir, f"Week_{week_num}_{today.year}.xlsx")
        _run_weekly(all_data_rows, all_serv_rows, all_errors, total_serv_count,
                    fx_latest, output_dir, output_xlsx, week_num)

    if run["bookings"]:
        _run_bookings(all_data_rows, output_dir, week_num, today)

    if run["dashboard"]:
        _run_dashboard(all_data_rows, all_serv_rows, fx_latest, output_dir, week_num, today)

    if run["individual"]:
        if args.report_month:
            report_month = args.report_month
            report_year = today.year
        elif today.month == 1:
            report_month = 12
            report_year = today.year - 1
        else:
            report_month = today.month - 1
            report_year = today.year
        _run_individual(all_data_rows, output_dir, today, report_month, report_year)

    if run["ta_monthly"]:
        _run_ta_monthly(all_data_rows, output_dir, fx_latest, today)

    if run["ap_ar"]:
        _run_ap_ar(entities, output_dir, fx_latest, today)

    # ── Resumen ───────────────────────────────────────────────────────
    print(f"\n{'═' * 60}")
    print(f"  RESUMEN - Semana {week_num}")
    print(f"{'═' * 60}")
    for key, info in entity_info.items():
        print(f"\n  {info['label']} ({info['company']}):")
        print(f"    DATA: {info['n_data']} | DATA SERV: {info['n_serv']} | Errores: {info['n_errors']}")
    print(f"\n  Reportes generados: {', '.join(selected)}")
    print(f"\n═══ Proceso completado ═══")


if __name__ == "__main__":
    main()
