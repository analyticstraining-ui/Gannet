"""
Detección de errores en datos de reservas.
"""

import pandas as pd


def detect_errors(reserva_df, rentabilidad_df):
    """Detect potential data issues to report.

    Returns:
        List of dicts with keys: Folio, Error, Vendedor, Fecha.
    """
    errors = []

    merged = reserva_df.merge(rentabilidad_df, on="folio", how="left")
    merged["rentabilidad"] = merged["rentabilidad"].fillna(0)

    for _, r in merged.iterrows():
        folio = r["folio"]
        vendedor = str(r.get("vendedor", "")).strip()
        fecha = r.get("fecha")
        total = float(r.get("total_cliente", 0) or 0)
        rent = float(r.get("rentabilidad", 0) or 0)
        moneda = str(r.get("moneda", "")).strip()

        def add_error(mensaje):
            errors.append({
                "Folio": folio,
                "Error": mensaje,
                "Vendedor": vendedor,
                "Fecha": fecha,
            })

        if total < 0:
            add_error(f"total_cliente negativo ({total})")

        if rent < 0:
            add_error(f"rentabilidad negativa ({rent})")

        if total > 0 and (rent / total) > 0.5:
            add_error(f"rentabilidad muy alta ({rent/total:.1%} de {total})")

        if total == 0 and rent != 0:
            add_error(f"total_cliente=0 pero tiene rentabilidad ({rent})")

        if moneda not in ("EUR", "USD", "GBP", "GPB", "CHF", "MXN", "JPY"):
            add_error(f"moneda desconocida '{moneda}'")

        if pd.isna(r.get("fecha_inicio")):
            add_error("sin fecha_inicio")
        if pd.isna(r.get("fecha_fin")):
            add_error("sin fecha_fin")

    return errors
