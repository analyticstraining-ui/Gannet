"""
Actualiza filtros de semana en los pivot tables del Weekly Report.

Filtra semanas 1..N en las hojas "Weekly SL y LLC" y "Bookings Fecha de Salida".
En "Bookings Fecha de Salida", pivots con Año anterior no se filtran
(tienen semanas 38-52 de ventas del año pasado).
"""

from datetime import datetime

_WEEK_FILTER_SHEETS = {"Weekly SL y LLC", "Bookings Fecha de Salida"}


def _get_pivot_año(pt, cache):
    """Lee el valor del filtro 'Año' de un pivot (pageField)."""
    if not pt.pageFields:
        return None
    for pf in pt.pageFields:
        cf = cache.cacheFields[pf.fld]
        if not cf.name or cf.name.strip().lower() != 'año':
            continue
        if pf.item is None:
            return None
        pivot_field = pt.pivotFields[pf.fld]
        if not pivot_field.items or pf.item >= len(pivot_field.items):
            return None
        x = getattr(pivot_field.items[pf.item], 'x', None)
        if x is None or not cf.sharedItems or not cf.sharedItems._fields:
            return None
        if x < len(cf.sharedItems._fields):
            try:
                return int(float(getattr(cf.sharedItems._fields[x], 'v', 0)))
            except (TypeError, ValueError):
                return None
    return None


def update_pivot_week_filters(wb, week_num):
    """Actualiza filtros de Semana en pivots de las hojas indicadas.

    En "Weekly SL y LLC": filtra ambos pivots (2025 y 2026) a semanas 1..N.
    En "Bookings Fecha de Salida": solo filtra pivots con Año >= actual.

    Args:
        wb: Workbook abierto de openpyxl.
        week_num: Número de semana máximo a mostrar.

    Returns:
        Número de pivots actualizados.
    """
    current_year = datetime.now().year
    updated = 0

    for ws in wb.worksheets:
        if ws.title not in _WEEK_FILTER_SHEETS:
            continue

        for pt in getattr(ws, '_pivots', []):
            cache = pt.cache
            if not cache or not cache.cacheFields:
                continue

            # En "Bookings Fecha de Salida", pivots con Año < actual tienen
            # semanas 38-52 (ventas 2025), no filtrar para no vaciarlos.
            # En "Weekly SL y LLC", ambos pivots usan semanas 1-N, filtrar todos.
            año = _get_pivot_año(pt, cache)
            if ws.title != "Weekly SL y LLC" and año is not None and año < current_year:
                print(f"  [{ws.title}] {pt.name}: Año={año} → sin filtro de semanas")
                continue

            semana_idx = None
            for i, cf in enumerate(cache.cacheFields):
                if cf.name and cf.name.strip().lower() == 'semana':
                    semana_idx = i
                    break

            if semana_idx is None:
                continue

            cf = cache.cacheFields[semana_idx]
            shared_vals = []
            if cf.sharedItems and cf.sharedItems._fields:
                for si in cf.sharedItems._fields:
                    shared_vals.append(getattr(si, 'v', None))

            if not shared_vals:
                continue

            pf = pt.pivotFields[semana_idx]
            if not pf.items:
                continue

            for item in pf.items:
                x = getattr(item, 'x', None)
                if x is None:
                    continue
                if x >= len(shared_vals):
                    item.h = True
                    continue
                val = shared_vals[x]
                try:
                    week_val = int(float(val))
                except (TypeError, ValueError):
                    item.h = True
                    continue
                item.h = week_val > week_num

            updated += 1
            visible = []
            for item in pf.items:
                x = getattr(item, 'x', None)
                if x is not None and x < len(shared_vals) and not getattr(item, 'h', False):
                    try:
                        visible.append(int(float(shared_vals[x])))
                    except (TypeError, ValueError):
                        pass
            if visible:
                print(f"  [{ws.title}] {pt.name}: Año={año} → semanas {min(visible)}-{max(visible)}")

    return updated
