"""
Detección de errores en datos de reservas.
"""

import pandas as pd


def detect_errors(reserva_df, rentabilidad_df):
    """Detect potential data issues to report."""
    errors = []

    merged = reserva_df.merge(rentabilidad_df, on="folio", how="left")
    merged["rentabilidad"] = merged["rentabilidad"].fillna(0)

    for _, r in merged.iterrows():
        folio = r["folio"]
        total = float(r.get("total_cliente", 0) or 0)
        rent = float(r.get("rentabilidad", 0) or 0)
        moneda = str(r.get("moneda", "")).strip()

        if total < 0:
            errors.append(f"Folio {folio}: total_cliente negativo ({total})")

        if rent < 0:
            errors.append(f"Folio {folio}: rentabilidad negativa ({rent})")

        if total > 0 and (rent / total) > 0.5:
            errors.append(
                f"Folio {folio}: rentabilidad muy alta ({rent/total:.1%} de {total})"
            )

        if total == 0 and rent != 0:
            errors.append(
                f"Folio {folio}: total_cliente=0 pero tiene rentabilidad ({rent})"
            )

        if moneda not in ("EUR", "USD", "GBP", "GPB", "CHF", "MXN", "JPY"):
            errors.append(f"Folio {folio}: moneda desconocida '{moneda}'")

        if pd.isna(r.get("fecha_inicio")):
            errors.append(f"Folio {folio}: sin fecha_inicio")
        if pd.isna(r.get("fecha_fin")):
            errors.append(f"Folio {folio}: sin fecha_fin")

    return errors
