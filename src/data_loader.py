"""
Lectura de CSVs, filtrado y limpieza de datos.
"""

import os
import glob

import pandas as pd


def _find_csv(data_dir, base_name):
    """Find a CSV file in data_dir, handling variants like 'reserva (1).csv'."""
    # Try exact name first
    exact = os.path.join(data_dir, f"{base_name}.csv")
    if os.path.exists(exact):
        return exact

    # Try recursive search for any variant
    pattern = os.path.join(data_dir, "**", f"{base_name}*.csv")
    matches = glob.glob(pattern, recursive=True)
    if matches:
        return matches[0]

    # Try inside subdirectories (backup folders)
    pattern = os.path.join(data_dir, "**", f"{base_name}.csv")
    matches = glob.glob(pattern, recursive=True)
    if matches:
        return matches[0]

    return None


def load_reserva(data_dir):
    """Read reserva.csv, filter cancelled, sort by folio."""
    path = _find_csv(data_dir, "reserva")
    if not path:
        raise FileNotFoundError(f"No se encuentra reserva.csv en {data_dir}")

    print(f"  Leyendo {path}...")
    df = pd.read_csv(path, encoding="latin-1")
    total = len(df)
    df = df[df["cancelada"] == 0].copy()
    df.sort_values("folio", inplace=True)
    df.reset_index(drop=True, inplace=True)
    print(f"  {total} registros leídos, {len(df)} después de filtrar canceladas.")
    return df


def load_dreserva(data_dir):
    """Read dreserva.csv."""
    path = _find_csv(data_dir, "dreserva")
    if not path:
        raise FileNotFoundError(f"No se encuentra dreserva.csv en {data_dir}")

    print(f"  Leyendo {path}...")
    df = pd.read_csv(path, encoding="latin-1")
    print(f"  {len(df)} registros de detalle leídos.")
    return df


def compute_rentabilidad(dreserva_df):
    """Sum monto_comision by folio for rentabilidad."""
    col = "monto_comision" if "monto_comision" in dreserva_df.columns else "comision_monto"
    rent = dreserva_df.groupby("folio")[col].sum().reset_index()
    rent.columns = ["folio", "rentabilidad"]
    return rent
