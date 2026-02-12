"""
API de tipos de cambio con fallback a valores estáticos.
"""

from config import FALLBACK_FX


def get_fx_rates():
    """Fetch exchange rates from API, fallback to hardcoded rates."""
    try:
        import requests
        resp = requests.get(
            "https://api.exchangerate-api.com/v4/latest/EUR", timeout=10
        )
        resp.raise_for_status()
        data = resp.json()
        rates_from_eur = data["rates"]  # 1 EUR = X currency

        fx = {}
        for currency in ["EUR", "USD", "GBP", "GPB", "CHF", "JPY", "MXN"]:
            lookup = "GBP" if currency == "GPB" else currency
            if lookup in rates_from_eur:
                rate_from_eur = rates_from_eur[lookup]
                to_eur = 1.0 / rate_from_eur
                usd_per_eur = rates_from_eur.get("USD", 1.16)
                to_usd = to_eur * usd_per_eur
                fx[currency] = {"EUR": round(to_eur, 6), "USD": round(to_usd, 6)}
            else:
                fx[currency] = FALLBACK_FX.get(currency, {"EUR": 1.0, "USD": 1.0})

        # Ensure EUR is exact
        fx["EUR"] = {"EUR": 1.0, "USD": round(rates_from_eur.get("USD", 1.16), 6)}

        print(f"  Tipos de cambio obtenidos de API:")
        for cur, rates in sorted(fx.items()):
            print(f"    {cur} -> EUR: {rates['EUR']}, USD: {rates['USD']}")
        return fx

    except Exception as e:
        print(f"  ⚠ API de tipos de cambio falló ({e}), usando valores del template.")
        return FALLBACK_FX.copy()
