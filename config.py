"""
Configuración central del proyecto Gannet Reports.
"""

import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ── Entity configuration ──────────────────────────────────────────────────
ENTITIES = {
    "espana": {
        "company": "SL",
        "label": "España",
        "data_dir": os.path.join(BASE_DIR, "data", "espana"),
        "output_dir": os.path.join(BASE_DIR, "output", "espana"),
    },
    "mexico": {
        "company": "LLC",
        "label": "México",
        "data_dir": os.path.join(BASE_DIR, "data", "mexico"),
        "output_dir": os.path.join(BASE_DIR, "output", "mexico"),
    },
}

TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "Week 6.xlsx")
TA_TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "Reporte_TAs_template.xlsx")
AP_AR_TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "AP y AR SL .xlsx")
COM_PEND_PROV_TEMPLATE_PATH = os.path.join(BASE_DIR, "templates", "Comisiones Pendientes Prov.xlsx")

# ── Fallback FX rates ─────────────────────────────────────────────────────
FALLBACK_FX = {
    "EUR": {"EUR": 1.0, "USD": 1.16},
    "USD": {"EUR": 0.86, "USD": 1.0},
    "GBP": {"EUR": 1.15, "USD": 1.34},
    "GPB": {"EUR": 1.15, "USD": 1.34},  # typo in source data
    "CHF": {"EUR": 1.07, "USD": 1.25},
    "JPY": {"EUR": 0.0055, "USD": 0.0064},
    "MXN": {"EUR": 0.046, "USD": 0.054},
}

# ── Spanish month names ───────────────────────────────────────────────────
MONTH_NAMES_ES = {
    1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
    5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
    9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
}
