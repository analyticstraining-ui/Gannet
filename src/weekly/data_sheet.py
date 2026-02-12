"""
Construye el DataFrame para la hoja DATA del Weekly Report.
"""

from datetime import datetime, timedelta

import pandas as pd

from config import MONTH_NAMES_ES


def _parse_date(val):
    """Parse a date value that might be string or already datetime."""
    if pd.isna(val):
        return None
    if isinstance(val, datetime):
        return val
    try:
        return pd.to_datetime(val)
    except Exception:
        return None


def build_data_rows(reserva_df, rentabilidad_df, fx, company):
    """Build the DATA sheet rows (columns A-Z).

    Args:
        reserva_df: Filtered reserva DataFrame.
        rentabilidad_df: Rentabilidad by folio.
        fx: Exchange rates dict {currency: {EUR: rate, USD: rate}}.
        company: Entity identifier ("SL" or "LLC").

    Returns:
        List of dicts with column letter keys.
    """
    merged = reserva_df.merge(rentabilidad_df, on="folio", how="left")
    merged["rentabilidad"] = merged["rentabilidad"].fillna(0)

    rows = []
    for _, r in merged.iterrows():
        fecha = _parse_date(r.get("fecha"))
        fecha_inicio = _parse_date(r.get("fecha_inicio"))
        fecha_fin = _parse_date(r.get("fecha_fin"))

        moneda = str(r.get("moneda", "EUR")).strip()
        total_cliente = float(r.get("total_cliente", 0) or 0)
        rentabilidad = float(r.get("rentabilidad", 0) or 0)

        fx_rates = fx.get(moneda, fx.get("EUR", {"EUR": 1.0, "USD": 1.16}))
        fx_eur = fx_rates["EUR"]
        fx_usd = fx_rates["USD"]

        semana = fecha.isocalendar()[1] if fecha else None
        fecha_45 = fecha_fin + timedelta(days=45) if fecha_fin else None
        mes_45_nombre = MONTH_NAMES_ES.get(fecha_45.month) if fecha_45 else None
        pct_rent = (rentabilidad / total_cliente) if total_cliente != 0 else 0

        row = {
            "A": company,
            "B": r.get("folio"),
            "C": r.get("cerrada", 0),
            "D": fecha,
            "E": fecha_inicio,
            "F": fecha_fin,
            "G": r.get("vendedor"),
            "H": semana,
            "I": r.get("usuarios_invitados"),
            "J": total_cliente,
            "K": moneda,
            "L": round(total_cliente * fx_eur, 2),
            "M": round(total_cliente * fx_usd, 2),
            "N": rentabilidad,
            "O": round(rentabilidad * fx_eur, 2),
            "P": round(rentabilidad * fx_usd, 2),
            "Q": round(pct_rent, 6),
            "R": fecha.month if fecha else None,
            "S": fecha.year if fecha else None,
            "T": fecha_inicio.month if fecha_inicio else None,
            "U": fecha_inicio.year if fecha_inicio else None,
            "V": fecha_45,
            "W": mes_45_nombre,
            "X": fecha_45.year if fecha_45 else None,
            "Z": r.get("observaciones", ""),
        }
        rows.append(row)

    return rows
