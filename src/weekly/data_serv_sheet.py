"""
Construye los datos para la hoja DATA SERV del Weekly Report.
"""

from datetime import datetime

import pandas as pd


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


def build_serv_rows(dreserva_df, reserva_df, company):
    """Build DATA SERV data rows.

    Static values in columns B,C,E-I,K,L,O.
    Columns with formulas (A,D,J,M,N,P,Q,R,S,T,U,V) are written by excel_writer.

    Args:
        dreserva_df: Detail reserva DataFrame.
        reserva_df: Filtered reserva DataFrame.
        company: Entity identifier ("SL" or "LLC").

    Returns:
        Tuple of (list of row dicts, row count).
    """
    active_folios = set(reserva_df["folio"].values)
    df = dreserva_df[dreserva_df["folio"].isin(active_folios)].copy()
    df.sort_values(["folio", "numero"], inplace=True)
    df.reset_index(drop=True, inplace=True)

    col_name = "monto_comision" if "monto_comision" in df.columns else "comision_monto"

    rows = []
    for _, r in df.iterrows():
        inicio = _parse_date(r.get("inicio_estancia"))
        fin = _parse_date(r.get("fin_estancia"))

        row = {
            "B": company,
            "C": r.get("folio"),
            "E": r.get("proveedor"),
            "F": r.get("descripcion"),
            "G": inicio,
            "H": fin,
            "I": r.get("tipo_servicio"),
            "K": r.get("moneda"),
            "L": r.get("subtotal", 0),
            "O": r.get(col_name, 0),
        }
        rows.append(row)

    return rows, len(df)
