# Gannet Reports

Automatizacion del Weekly Sales Report y Bookings de Gannet (Corsario SL y LLC).

## Que hace

A partir de los CSVs de Corsario (reserva y dreserva), genera:

| Archivo | Descripcion |
|---------|-------------|
| `output/espana/Bookings SL.xlsx` | Reservas de Espana — control interno |
| `output/mexico/Bookings LLC.xlsx` | Reservas de Mexico — control interno |
| `output/Week N.xlsx` | Weekly Report combinado (SL + LLC) con pivots y booking window |

Ademas:
- Obtiene tipos de cambio del dia (API) para convertir montos a EUR y USD
- Detecta errores en los datos (montos negativos, rentabilidad alta, fechas faltantes) para reportar a Madrid
- Configura los 9 pivot tables para auto-refresh al abrir en Excel

## Requisitos

- Python 3.10+
- Dependencias: `pip install -r requirements.txt`

## Uso

```bash
# Ambas entidades, semana actual
python3 main.py

# Semana especifica
python3 main.py --week 7

# Solo una entidad
python3 main.py --entity espana
python3 main.py --entity mexico

# Combinado
python3 main.py --week 7 --entity mexico
```

## Workflow semanal

### 1. Descargar datos de Corsario

- Corsario SL: Configuracion > Respaldar datos > Reservaciones
- Corsario LLC: lo mismo

### 2. Colocar los CSVs

```
data/
  espana/
    reserva.csv      <- de Corsario SL
    dreserva.csv     <- de Corsario SL
    proveedor.csv    <- de Corsario SL
    usuario.csv      <- de Corsario SL
  mexico/
    reserva.csv      <- de Corsario LLC (puede llamarse "reserva (1).csv", lo detecta solo)
    dreserva.csv     <- de Corsario LLC
```

### 3. Ejecutar

```bash
python3 main.py --week 7
```

### 4. Recoger output

- `output/espana/Bookings SL.xlsx` — control interno Espana
- `output/mexico/Bookings LLC.xlsx` — control interno Mexico
- `output/Week 7.xlsx` — abrir en Microsoft Excel para que los pivots se refresquen, enviarselo a Pedro

### 5. Errores

Los errores que aparezcan en consola se los reportas al equipo de Madrid.

## Ejecucion automatica (GitHub Actions)

El workflow `.github/workflows/weekly_report.yml` se ejecuta automaticamente:
- **Cada lunes a las 7:00 AM hora Madrid** (ajusta cambio horario verano/invierno)
- Se puede lanzar manualmente desde GitHub > Actions > Run workflow

Para que funcione:
1. Subir los CSVs nuevos al repo (`git push`)
2. El lunes se ejecuta solo
3. Descargar el output desde GitHub > Actions > Artifacts > weekly-report

## Estructura del proyecto

```
Gannet/
├── main.py                         # Punto de entrada
├── config.py                       # Rutas, entidades, FX fallback, meses
├── requirements.txt                # pandas, openpyxl, requests
├── .gitignore
│
├── .github/workflows/
│   └── weekly_report.yml           # GitHub Actions: lunes 7am Madrid
│
├── src/
│   ├── data_loader.py              # Lee CSVs, filtra canceladas, calcula rentabilidad
│   ├── fx_rates.py                 # API tipos de cambio + fallback
│   ├── validators.py               # Detecta errores en los datos
│   │
│   ├── weekly/                     # Modulo Weekly Report
│   │   ├── data_sheet.py           # Construye filas de hoja DATA (cols A-Z)
│   │   ├── data_serv_sheet.py      # Construye filas de hoja DATA SERV
│   │   └── excel_writer.py         # Copia plantilla, escribe datos, formulas XLOOKUP, pivots
│   │
│   └── bookings/                   # Modulo Bookings
│       ├── booking_window.py       # Matriz booking window (semana x mes) en USD
│       └── export_bookings.py      # Exporta datos DATA como Excel por entidad
│
├── data/                           # CSVs de Corsario (se sobreescriben cada semana)
│   ├── espana/
│   └── mexico/
│
├── templates/
│   └── Week 6.xlsx                 # Plantilla base (no se modifica)
│
└── output/                         # Archivos generados (gitignored)
    ├── Week N.xlsx
    ├── espana/Bookings SL.xlsx
    └── mexico/Bookings LLC.xlsx
```

## Flujo interno

```
python3 main.py --week 7
       |
       v
[1] Tipos de cambio (API o fallback)
       |
       v
[2] Bookings SL (control interno)
       ├── Lee data/espana/reserva.csv + dreserva.csv
       ├── Filtra canceladas, calcula rentabilidad
       ├── Genera output/espana/Bookings SL.xlsx
       └── Valida errores
       |
       v
[2] Bookings LLC (control interno)
       ├── Lee data/mexico/reserva.csv + dreserva.csv
       ├── Genera output/mexico/Bookings LLC.xlsx
       └── Valida errores
       |
       v
[3] Weekly Report combinado
       ├── Copia plantilla templates/Week 6.xlsx
       ├── Escribe SL + LLC juntos en hoja DATA
       ├── Escribe SL + LLC juntos en hoja DATA SERV (con XLOOKUP)
       ├── Llena Booking Window 2026 (combinado, USD)
       ├── Actualiza tipos de cambio en lookup table
       ├── Configura pivots para auto-refresh
       └── Guarda output/Week 7.xlsx
```

## Hojas del Weekly Report

| Hoja | Que contiene |
|------|-------------|
| DATA | Todas las reservas activas (SL + LLC): folio, fechas, vendedor, montos EUR/USD, rentabilidad |
| DATA SERV | Detalle por servicio con formulas XLOOKUP (proveedor, tipo servicio, conversiones FX) |
| Weekly SL y LLC | Pivot: venta y rentabilidad por semana, filtrado por compania y ano |
| Bookings Fecha de Salida | 3 pivots: ventas por semana y mes de salida (2025→2026, 2026→2026, 2026→2027) |
| Booking Window 2026 | Matriz manual: semana de venta x mes de salida en USD |
| Por servicio y TA Weekly | Pivot de DATA SERV: rentabilidad por vendedor y tipo de servicio |

## Notas

- La plantilla `Week 6.xlsx` contiene las lookup tables (proveedores SL/LLC, usuarios, tipos de servicio, FX) en columnas Y-BO de DATA SERV. No se modifican.
- Los pivots se configuran con `refreshOnLoad=True` pero solo se refrescan al abrir en **Microsoft Excel** (no funciona en Google Sheets ni LibreOffice).
- El script detecta automaticamente si el CSV se llama `reserva.csv` o `reserva (1).csv`.
- "GPB" es un typo de Corsario para GBP — el script lo maneja.
