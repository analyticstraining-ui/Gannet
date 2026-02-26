#!/usr/bin/env python3
"""
Verificación del Reporte AP & AR usando FX rates del template manual.

Usa los mismos datos (CSVs) pero con los FX rates estáticos del template
para poder comparar 1:1 con el reporte hecho a mano.

Uso:
    python3 verify_ap_ar.py                        # todas las entidades
    python3 verify_ap_ar.py --entity espana         # solo España
    python3 verify_ap_ar.py --entity mexico         # solo México
"""

import argparse
import os
import sys
from datetime import datetime

from config import ENTITIES, BASE_DIR

# ── FX rates EXACTOS del template AP & AR ─────────────────────────────
TEMPLATE_FX = {
    "EUR": {"EUR": 1.0, "USD": 1.16},
    "USD": {"EUR": 0.86, "USD": 1.0},
    "CHF": {"EUR": 1.07, "USD": 1.25},
    "GBP": {"EUR": 1.15, "USD": 1.34},
    "GPB": {"EUR": 1.15, "USD": 1.34},
    "JPY": {"EUR": 0.0055, "USD": 0.0064},
    "MXN": {"EUR": 0.048, "USD": 0.056},
}

_AP_AR_ENTITY_LABELS = {"SL": "Madrid", "LLC": "Mexico"}


def main():
    parser = argparse.ArgumentParser(
        description="Genera Reporte AP & AR con FX rates del template"
    )
    parser.add_argument("--entity", type=str, default="all",
                        choices=["espana", "mexico", "all"],
                        help="Entidad a procesar (default: all)")
    args = parser.parse_args()

    today = datetime.now()

    print(f"═══ VERIFICACIÓN AP & AR ═══")
    print(f"  FX rates exactos del template")
    print(f"  EUR→USD: {TEMPLATE_FX['EUR']['USD']}, USD→EUR: {TEMPLATE_FX['USD']['EUR']}")
    print(f"  MXN→EUR: {TEMPLATE_FX['MXN']['EUR']}, MXN→USD: {TEMPLATE_FX['MXN']['USD']}")

    if args.entity == "all":
        entities = ENTITIES
    else:
        entities = {args.entity: ENTITIES[args.entity]}

    output_dir = os.path.join(BASE_DIR, "output")
    os.makedirs(output_dir, exist_ok=True)

    from src.ap_ar import generate_ap_ar_report

    for key, cfg in entities.items():
        company = cfg["company"]
        data_dir = cfg["data_dir"]
        entity_label = _AP_AR_ENTITY_LABELS.get(company, cfg["label"])

        if not os.path.exists(data_dir):
            print(f"\n  Saltando {entity_label}: {data_dir} no encontrado")
            continue

        print(f"\n{'─' * 60}")
        print(f"  {entity_label} ({company}) — FX de template")
        print(f"{'─' * 60}")

        report_path = generate_ap_ar_report(
            data_dir=data_dir,
            output_dir=output_dir,
            fx=TEMPLATE_FX,
            company=company,
            entity_label=entity_label + "_VERIFY",
            year=today.year,
            month=today.month,
        )

        if report_path:
            print(f"\n  Archivo: {report_path}")
        else:
            print(f"\n  ERROR generando reporte para {entity_label}")

    print(f"\n═══ Verificación AP & AR completada ═══")


if __name__ == "__main__":
    main()
