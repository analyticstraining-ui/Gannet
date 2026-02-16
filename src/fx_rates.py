"""
API de tipos de cambio históricos por fecha de reserva.

Usa la API de Frankfurter (datos del BCE, gratuita, sin API key)
para obtener la tasa del día exacto de cada reserva.

Para fines de semana / festivos usa la tasa del último día laborable anterior.
"""

import calendar
from datetime import date, datetime, timedelta

from config import FALLBACK_FX

# Monedas a pedir (EUR es la base, no se pide)
_CURRENCIES = "USD,GBP,CHF,JPY,MXN"

# Cache global: {date: {moneda: {EUR: float, USD: float}}}
_fx_cache = {}

# Meses ya descargados (para no repetir llamadas)
_downloaded_months = set()


# Nombres de mes en español
_MONTH_NAMES = {
    1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic",
}


# ── Funciones públicas ─────────────────────────────────────────────────

def get_current_month_daily_rates():
    """
    Descarga y devuelve las tasas de cada día laborable del mes actual.

    Returns:
        (label, daily_list) donde:
        - label: "Febrero 2026"
        - daily_list: [{"date": date, "rates": {moneda: {EUR, USD}}}, ...]
    """
    today = date.today()
    year, month = today.year, today.month

    # Asegurar que el mes actual está descargado
    if (year, month) not in _downloaded_months:
        _download_month(year, month)

    _MONTH_FULL = {
        1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
        5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
        9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
    }
    label = f"{_MONTH_FULL[month]} {year}"

    daily = []
    for d in sorted(_fx_cache.keys()):
        if d.year == year and d.month == month:
            daily.append({"date": d, "rates": _fx_cache[d]})

    return label, daily


def preload_fx_months(dates):
    """
    Pre-descarga las tasas de todos los meses que aparecen en una lista de fechas.
    Acepta una pandas Series, lista de datetime/date, o cualquier iterable.
    """
    months_needed = set()
    for d in dates:
        if d is None:
            continue
        if hasattr(d, 'date'):
            d = d.date()
        if isinstance(d, date):
            months_needed.add((d.year, d.month))

    to_download = months_needed - _downloaded_months
    if not to_download:
        return

    sorted_months = sorted(to_download)
    print(f"    Descargando tasas históricas: {len(sorted_months)} meses "
          f"({sorted_months[0][0]}-{sorted_months[0][1]:02d} a "
          f"{sorted_months[-1][0]}-{sorted_months[-1][1]:02d})...")

    for year, month in sorted_months:
        _download_month(year, month)

    print(f"    Cache FX: {len(_fx_cache)} días laborables cargados.")


def get_historical_fx(target_date):
    """
    Obtiene tipos de cambio para una fecha específica.
    Si es fin de semana/festivo, retrocede al último día laborable.

    Args:
        target_date: date o datetime (se extrae .date()).

    Returns:
        dict {moneda: {EUR: float, USD: float}} con EUR, USD, GBP, GPB, CHF, JPY, MXN.
    """
    if target_date is None:
        return FALLBACK_FX.copy()

    if isinstance(target_date, datetime):
        target_date = target_date.date()

    # Asegurar que el mes está descargado
    month_key = (target_date.year, target_date.month)
    if month_key not in _downloaded_months:
        _download_month(target_date.year, target_date.month)

    return _find_nearest_rate(target_date)


def get_latest_fx():
    """
    Obtiene tipos de cambio del último día laborable disponible.
    Se usa para la lookup table (AM-AO) y la hoja FX RATES.
    """
    try:
        import requests
        resp = requests.get(
            f"https://api.frankfurter.app/latest?from=EUR&to={_CURRENCIES}",
            timeout=10
        )
        resp.raise_for_status()
        data = resp.json()
        fx = _parse_single_day(data["rates"])

        print(f"  Tipos de cambio (últimos disponibles, {data.get('date', 'hoy')}):")
        for cur, rates in sorted(fx.items()):
            print(f"    {cur} → EUR: {rates['EUR']}, USD: {rates['USD']}")
        return fx

    except Exception as e:
        print(f"  ⚠ API latest FX falló ({e}), usando valores fallback.")
        return FALLBACK_FX.copy()


# ── Funciones internas ─────────────────────────────────────────────────

def _download_month(year, month):
    """Descarga todas las tasas de un mes desde Frankfurter y las cachea."""
    last_day = calendar.monthrange(year, month)[1]
    start = f"{year}-{month:02d}-01"
    end = f"{year}-{month:02d}-{last_day:02d}"

    try:
        import requests
        url = f"https://api.frankfurter.app/{start}..{end}?from=EUR&to={_CURRENCIES}"
        resp = requests.get(url, timeout=15)
        resp.raise_for_status()
        data = resp.json()

        for date_str, rates in data.get("rates", {}).items():
            d = date.fromisoformat(date_str)
            _fx_cache[d] = _parse_single_day(rates)

    except Exception as e:
        print(f"    ⚠ FX mes {year}-{month:02d} falló ({e})")

    # Marcar como intentado (para no reintentar si falló)
    _downloaded_months.add((year, month))


def _parse_single_day(rates_from_eur):
    """
    Convierte tasas de un día (1 EUR = X moneda) al formato interno
    {moneda: {EUR: cuántos EUR es 1 moneda, USD: cuántos USD es 1 moneda}}.
    """
    usd_per_eur = rates_from_eur.get("USD", 1.16)

    fx = {}
    for currency in ["EUR", "USD", "GBP", "GPB", "CHF", "JPY", "MXN"]:
        lookup = "GBP" if currency == "GPB" else currency

        if lookup == "EUR":
            fx[currency] = {"EUR": 1.0, "USD": round(usd_per_eur, 6)}
        elif lookup in rates_from_eur:
            rate_from_eur = rates_from_eur[lookup]  # 1 EUR = X moneda
            to_eur = 1.0 / rate_from_eur             # 1 moneda = X EUR
            to_usd = to_eur * usd_per_eur            # 1 moneda = X USD
            fx[currency] = {"EUR": round(to_eur, 6), "USD": round(to_usd, 6)}
        else:
            fx[currency] = FALLBACK_FX.get(currency, {"EUR": 1.0, "USD": 1.0})

    return fx


def _find_nearest_rate(target_date):
    """
    Busca la tasa de target_date. Si no existe (fin de semana/festivo),
    retrocede hasta 7 días buscando el último día laborable con datos.
    """
    for i in range(8):
        d = target_date - timedelta(days=i)

        if d in _fx_cache:
            return _fx_cache[d]

        # Si retrocedemos al mes anterior, asegurar que esté descargado
        month_key = (d.year, d.month)
        if month_key not in _downloaded_months:
            _download_month(d.year, d.month)
            if d in _fx_cache:
                return _fx_cache[d]

    # Si no encontramos nada en 7 días, usar fallback
    return FALLBACK_FX.copy()
