#!/usr/bin/env python3
"""
Generador automático del Weekly Sales Report de Gannet.

Flujo:
    1. Obtener tipos de cambio (una vez)
    2. Por cada entidad: cargar datos, generar Bookings (control interno), validar
    3. Combinar ambos bookings → generar Weekly Report con pivots + booking window

Uso:
    python3 main.py                          # ambas entidades, semana actual
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
from src.bookings import (
    build_booking_matrix, write_booking_to_excel, export_booking_xlsx,
    export_bookings_xlsx,
)


def main():
    parser = argparse.ArgumentParser(
        description="Genera el Weekly Sales Report de Gannet"
    )
    parser.add_argument(
        "--week", type=int, default=None,
        help="Número de semana (default: semana actual ISO)"
    )
    parser.add_argument(
        "--entity", type=str, default="all",
        choices=["espana", "mexico", "all"],
        help="Entidad a procesar (default: all)"
    )
    args = parser.parse_args()

    today = datetime.now()
    week_num = args.week or today.isocalendar()[1]

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
    # Step 2: BOOKINGS por entidad (control interno)
    #   - Se crean PRIMERO porque el Weekly se basa en ellos
    # ══════════════════════════════════════════════════════════════
    all_data_rows = []
    all_serv_rows = []
    total_serv_count = 0
    entity_info = {}

    for key, cfg in entities.items():
        company = cfg["company"]
        label = cfg["label"]
        data_dir = cfg["data_dir"]
        output_dir = cfg["output_dir"]

        if not os.path.exists(data_dir):
            print(f"\n⚠ Directorio de datos no encontrado para {label}: {data_dir}")
            print(f"  Saltando {label}...")
            continue

        print(f"\n{'─' * 60}")
        print(f"  [2] Bookings {label} ({company}) — Control Interno")
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

        # Bookings per entity (same data as DATA tab, separated)
        os.makedirs(output_dir, exist_ok=True)
        bookings_path = os.path.join(output_dir, f"Bookings {company}.xlsx")
        export_bookings_xlsx(data_rows, bookings_path, company)

        # Validate and report errors
        errors = detect_errors(reserva_df, rentabilidad_df)
        if errors:
            print(f"\n  ⚠ {len(errors)} posibles errores para reportar:")
            for err in errors:
                print(f"    - {err}")
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
            "bookings_path": bookings_path,
        }

    if not all_data_rows:
        print("\nERROR: No se cargaron datos de ninguna entidad.")
        sys.exit(1)

    # ══════════════════════════════════════════════════════════════
    # Step 3: WEEKLY REPORT combinado (SL + LLC → 1 archivo)
    #   - Copia ambos bookings a 1 sheet + pivots + booking window
    # ══════════════════════════════════════════════════════════════
    output_dir = os.path.join(BASE_DIR, "output")
    os.makedirs(output_dir, exist_ok=True)
    output_xlsx = os.path.join(output_dir, f"Week {week_num}.xlsx")

    print(f"\n{'─' * 60}")
    print(f"  [3] Weekly Report combinado — Week {week_num}")
    print(f"{'─' * 60}")
    print(f"  Copiando bookings SL + LLC a 1 sheet...")
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

    print(f"  Guardando {output_xlsx}...")
    wb.save(output_xlsx)
    wb.close()
    print(f"  Weekly Report generado exitosamente.")

    # ── Summary ───────────────────────────────────────────────────
    print(f"\n{'═' * 60}")
    print(f"  RESUMEN - Semana {week_num}")
    print(f"{'═' * 60}")
    for key, info in entity_info.items():
        print(f"\n  Bookings {info['label']} ({info['company']}):")
        print(f"    DATA: {info['n_data']} | DATA SERV: {info['n_serv']} | Errores: {info['n_errors']}")
        print(f"    Bookings: {info['bookings_path']}")
    print(f"\n  Weekly Report: {output_xlsx}")
    print(f"    DATA: {len(all_data_rows)} filas | DATA SERV: {total_serv_count} filas")

    print(f"\n═══ Proceso completado ═══")
    print("  Abrir Week en Microsoft Excel para que los pivots se refresquen.")


if __name__ == "__main__":
    main()
