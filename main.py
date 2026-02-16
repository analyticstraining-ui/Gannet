#!/usr/bin/env python3
"""
Generador automático del Weekly Sales Report de Gannet.

Flujo:
    1. Obtener tipos de cambio (últimos + históricos por fecha de reserva)
    2. Por cada entidad: cargar datos, validar
    3. Combinar → generar Weekly Report con pivots + booking window + FX + errores

Uso:
    python3 main.py                          # ambas entidades, semana anterior
    python3 main.py --week 7                 # semana específica
    python3 main.py --entity espana          # solo España
    python3 main.py --entity mexico          # solo México
    python3 main.py --week 7 --entity mexico # combinado
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


def main():
    parser = argparse.ArgumentParser(
        description="Genera el Weekly Sales Report de Gannet"
    )
    parser.add_argument(
        "--week", type=int, default=None,
        help="Número de semana (default: semana anterior)"
    )
    parser.add_argument(
        "--entity", type=str, default="all",
        choices=["espana", "mexico", "all"],
        help="Entidad a procesar (default: all)"
    )
    parser.add_argument(
        "--report-month", type=int, default=None,
        help="Mes para reportes individuales (default: mes anterior)"
    )
    args = parser.parse_args()

    today = datetime.now()
    week_num = args.week or int(today.strftime('%W'))

    print(f"═══ Gannet Reports - Semana {week_num} ═══")

    # Validate template
    if not os.path.exists(TEMPLATE_PATH):
        print(f"ERROR: No se encuentra la plantilla: {TEMPLATE_PATH}")
        sys.exit(1)

    # Determine entities to process
    if args.entity == "all":
        entities = ENTITIES
    else:
        entities = {args.entity: ENTITIES[args.entity]}

    # ── Step 1: FX rates ────────────────────────────────────────────
    print("\n[1] Obteniendo tipos de cambio...")
    fx_latest = get_latest_fx()

    # ── Step 2: Cargar datos por entidad ────────────────────────────
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
            print(f"\n⚠ Directorio de datos no encontrado para {label}: {data_dir}")
            print(f"  Saltando {label}...")
            continue

        print(f"\n{'─' * 60}")
        print(f"  [2] {label} ({company})")
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

        # Bookings per entity
        output_dir_entity = cfg["output_dir"]
        os.makedirs(output_dir_entity, exist_ok=True)
        bookings_path = os.path.join(output_dir_entity, f"Bookings_{company}_{week_num}_{today.year}.xlsx")
        export_bookings_xlsx(data_rows, bookings_path, company)

        # Validate
        errors = detect_errors(reserva_df, rentabilidad_df)
        for err in errors:
            err["Compañia"] = company
            all_errors.append(err)
        if errors:
            print(f"\n  ⚠ {len(errors)} posibles errores detectados")
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

    if not all_data_rows:
        print("\nERROR: No se cargaron datos de ninguna entidad.")
        sys.exit(1)

    # ── Step 3: Weekly Report combinado ─────────────────────────────
    output_dir = os.path.join(BASE_DIR, "output")
    os.makedirs(output_dir, exist_ok=True)
    output_xlsx = os.path.join(output_dir, f"Week_{week_num}_{today.year}.xlsx")

    print(f"\n{'─' * 60}")
    print(f"  [3] Weekly Report — Week {week_num}")
    print(f"{'─' * 60}")
    print(f"  Total: DATA={len(all_data_rows)} filas, DATA SERV={total_serv_count} filas")

    wb = generate_weekly_excel(
        TEMPLATE_PATH, output_xlsx, all_data_rows, all_serv_rows,
        total_serv_count, fx_latest
    )

    # Booking window
    combined_matrix = build_booking_matrix(all_data_rows)
    print(f"  Booking Window: semanas con datos: {sorted(combined_matrix.keys())}")
    write_booking_to_excel(wb, combined_matrix)

    # Diferencia %
    n_weeks = write_diferencia_pct(wb, all_data_rows)
    print(f"  Weekly SL y LLC: Diferencia % ({n_weeks} semanas)")

    # FX RATES
    month_label, n_days = write_fx_sheet(wb)
    print(f"  Hoja FX RATES: {n_days} días de {month_label}")

    # ERRORES
    n_errors = write_errores_sheet(wb, all_errors)
    if n_errors:
        print(f"  Hoja ERRORES: {n_errors} posibles errores")

    print(f"  Guardando {output_xlsx}...")
    wb.save(output_xlsx)
    wb.close()
    print(f"  Weekly Report generado exitosamente.")

    # ── Step 4: Dashboard ───────────────────────────────────────────
    dashboard_path = os.path.join(output_dir, f"Dashboard_{week_num}_{today.year}.xlsx")
    print(f"  Generando Dashboard...")
    generate_dashboard(all_data_rows, all_serv_rows, fx_latest, dashboard_path)
    print(f"  Dashboard generado: {dashboard_path}")

    # ── Step 5: Reportes individuales por TA ────────────────────────
    print(f"\n{'─' * 60}")
    print(f"  [5] Reportes Individuales por TA")
    print(f"{'─' * 60}")
    # Mes anterior por defecto (si estamos en enero, usar diciembre del año anterior)
    if args.report_month:
        report_month = args.report_month
        report_year = today.year
    elif today.month == 1:
        report_month = 12
        report_year = today.year - 1
    else:
        report_month = today.month - 1
        report_year = today.year
    generate_individual_reports(all_data_rows, output_dir, year=report_year, month=report_month)

    # ── Step 6: Reporte Mensual Consolidado TAs ──────────────────────
    print(f"\n{'─' * 60}")
    print(f"  [6] Reporte Mensual Consolidado de TAs")
    print(f"{'─' * 60}")
    ta_report_path = generate_ta_monthly_report(
        all_data_rows, output_dir, fx_latest, year=today.year, month=today.month
    )
    if ta_report_path:
        print(f"  Reporte generado: {ta_report_path}")
    else:
        print(f"  Saltando reporte (template no encontrado)")

    # ── Summary ─────────────────────────────────────────────────────
    print(f"\n{'═' * 60}")
    print(f"  RESUMEN - Semana {week_num}")
    print(f"{'═' * 60}")
    for key, info in entity_info.items():
        print(f"\n  {info['label']} ({info['company']}):")
        print(f"    DATA: {info['n_data']} | DATA SERV: {info['n_serv']} | Errores: {info['n_errors']}")
    print(f"\n  Weekly Report: {output_xlsx}")
    print(f"    DATA: {len(all_data_rows)} filas | DATA SERV: {total_serv_count} filas")
    print(f"    FX: Tasas históricas del BCE por fecha de reserva")

    print(f"\n═══ Proceso completado ═══")
    print("  Abrir Week en Microsoft Excel para que los pivots se refresquen.")


if __name__ == "__main__":
    main()
