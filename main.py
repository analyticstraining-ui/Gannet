#!/usr/bin/env python3
"""
Generador automático del Weekly Sales Report de Gannet.

Flujo:
    1. Obtener tipos de cambio (una vez)
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
from src.fx_rates import get_fx_rates
from src.validators import detect_errors
from src.weekly import build_data_rows, build_serv_rows, generate_weekly_excel
from src.bookings import build_booking_matrix, write_booking_to_excel, export_bookings_xlsx


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
    args = parser.parse_args()

    today = datetime.now()
    week_num = args.week or (today.isocalendar()[1] - 1)

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

    # ── Step 1: FX rates (once) ───────────────────────────────────
    print("\n[1] Obteniendo tipos de cambio...")
    fx = get_fx_rates()

    # ══════════════════════════════════════════════════════════════
    # Step 2: Cargar datos por entidad
    # ══════════════════════════════════════════════════════════════
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

        # Load data
        print(f"  Leyendo datos...")
        reserva_df = load_reserva(data_dir)
        dreserva_df = load_dreserva(data_dir)
        rentabilidad_df = compute_rentabilidad(dreserva_df)

        # Build rows
        data_rows = build_data_rows(reserva_df, rentabilidad_df, fx, company)
        serv_rows, serv_count = build_serv_rows(dreserva_df, reserva_df, company)
        print(f"  DATA: {len(data_rows)} filas | DATA SERV: {serv_count} filas")

        # Bookings per entity
        output_dir_entity = cfg["output_dir"]
        os.makedirs(output_dir_entity, exist_ok=True)
        bookings_path = os.path.join(output_dir_entity, f"Bookings_{company}_{week_num}_{today.year}.xlsx")
        export_bookings_xlsx(data_rows, bookings_path, company)

        # Validate and report errors
        errors = detect_errors(reserva_df, rentabilidad_df)
        for err in errors:
            all_errors.append((company, err))
        if errors:
            print(f"\n  ⚠ {len(errors)} posibles errores detectados")
        else:
            print(f"  Sin errores detectados.")

        # Accumulate for combined Weekly Report
        all_data_rows.extend(data_rows)
        all_serv_rows.extend(serv_rows)
        total_serv_count += serv_count

        entity_info[key] = {
            "label": label,
            "company": company,
            "n_data": len(data_rows),
            "n_serv": serv_count,
            "n_errors": len(errors),
        }

    if not all_data_rows:
        print("\nERROR: No se cargaron datos de ninguna entidad.")
        sys.exit(1)

    # ══════════════════════════════════════════════════════════════
    # Step 3: WEEKLY REPORT combinado
    # ══════════════════════════════════════════════════════════════
    output_dir = os.path.join(BASE_DIR, "output")
    os.makedirs(output_dir, exist_ok=True)
    output_xlsx = os.path.join(output_dir, f"Week_{week_num}_{today.year}.xlsx")

    print(f"\n{'─' * 60}")
    print(f"  [3] Weekly Report — Week {week_num}")
    print(f"{'─' * 60}")
    print(f"  Total: DATA={len(all_data_rows)} filas, DATA SERV={total_serv_count} filas")

    wb = generate_weekly_excel(
        TEMPLATE_PATH, output_xlsx, all_data_rows, all_serv_rows,
        total_serv_count, fx
    )

    # Booking window combinado dentro del Weekly
    combined_matrix = build_booking_matrix(all_data_rows)
    weeks_with_data = sorted(combined_matrix.keys())
    print(f"  Booking Window: semanas con datos: {weeks_with_data}")
    write_booking_to_excel(wb, combined_matrix)

    # ── Hoja FX RATES ────────────────────────────────────────────
    from openpyxl.styles import Font, PatternFill, Alignment
    if "FX RATES" in wb.sheetnames:
        del wb["FX RATES"]
    ws_fx = wb.create_sheet("FX RATES")
    hdr_font = Font(bold=True, color="FFFFFF")
    hdr_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    ws_fx.column_dimensions["A"].width = 12
    ws_fx.column_dimensions["B"].width = 16
    ws_fx.column_dimensions["C"].width = 16
    for c, h in enumerate(["Moneda", "→ EUR", "→ USD"], 1):
        cell = ws_fx.cell(1, c, h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal="center")
    for i, (moneda, rates) in enumerate(sorted(fx.items()), 2):
        ws_fx.cell(i, 1, moneda)
        cell_eur = ws_fx.cell(i, 2, rates.get("EUR", 0))
        cell_usd = ws_fx.cell(i, 3, rates.get("USD", 0))
        cell_eur.number_format = '0.000000'
        cell_usd.number_format = '0.000000'
    print(f"  Hoja FX RATES: {len(fx)} monedas")

    # ── Hoja ERRORES ─────────────────────────────────────────────
    if all_errors:
        if "ERRORES" in wb.sheetnames:
            del wb["ERRORES"]
        ws_err = wb.create_sheet("ERRORES")
        err_fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
        ws_err.column_dimensions["A"].width = 12
        ws_err.column_dimensions["B"].width = 80
        for c, h in enumerate(["Compania", "Error detectado"], 1):
            cell = ws_err.cell(1, c, h)
            cell.font = hdr_font
            cell.fill = err_fill
            cell.alignment = Alignment(horizontal="center")
        for i, (comp, err) in enumerate(all_errors, 2):
            ws_err.cell(i, 1, comp)
            ws_err.cell(i, 2, err)
        print(f"  Hoja ERRORES: {len(all_errors)} posibles errores")

    print(f"  Guardando {output_xlsx}...")
    wb.save(output_xlsx)
    wb.close()
    print(f"  Weekly Report generado exitosamente.")

    # ── Summary ───────────────────────────────────────────────────
    print(f"\n{'═' * 60}")
    print(f"  RESUMEN - Semana {week_num}")
    print(f"{'═' * 60}")
    for key, info in entity_info.items():
        print(f"\n  {info['label']} ({info['company']}):")
        print(f"    DATA: {info['n_data']} | DATA SERV: {info['n_serv']} | Errores: {info['n_errors']}")
    print(f"\n  Weekly Report: {output_xlsx}")
    print(f"    DATA: {len(all_data_rows)} filas | DATA SERV: {total_serv_count} filas")

    print(f"\n═══ Proceso completado ═══")
    print("  Abrir Week en Microsoft Excel para que los pivots se refresquen.")


if __name__ == "__main__":
    main()
