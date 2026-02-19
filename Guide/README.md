# Gannet Reports

Automatizacion de reportes de ventas y gestion de Gannet (Corsario SL y LLC).

## Que hace

A partir de los CSVs de Corsario (reserva y dreserva), genera 5 tipos de reportes:

| Archivo | Descripcion |
|---------|-------------|
| `output/Week_N_YYYY.xlsx` | Weekly Report combinado (SL + LLC) con pivots, booking window y FX |
| `output/espana/Bookings_SL_N_YYYY.xlsx` | Reservas de Espana — control interno |
| `output/mexico/Bookings_LLC_N_YYYY.xlsx` | Reservas de Mexico — control interno |
| `output/Bookings_ALL_N_YYYY.xlsx` | Bookings combinado (SL + LLC) |
| `output/Dashboard_N_YYYY.xlsx` | Dashboard con 10 hojas de analisis y graficos |
| `output/Reportes_Individuales_Mes_Ano/` | Reporte individual por cada Travel Advisor |
| `output/Reporte_TAs_Mes_YYYY.xlsx` | Reporte mensual consolidado de todos los TAs |

Ademas:
- Obtiene tipos de cambio historicos por fecha de reserva (API BCE via Frankfurter)
- Detecta errores en los datos (montos negativos, rentabilidad alta, fechas faltantes)
- Configura pivot tables para auto-refresh al abrir en Excel (Win + Mac)
- Genera hoja SUMMARY con datos pre-calculados (compatible con Apple Numbers)

## Requisitos

- Python 3.10+
- Dependencias: `pip install -r requirements.txt` (pandas, openpyxl, requests)

## Uso

```bash
# Todos los reportes, semana actual
python3 main.py

# Solo un tipo de reporte
python3 main.py --weekly
python3 main.py --bookings
python3 main.py --dashboard
python3 main.py --individual
python3 main.py --ta-monthly

# Combinaciones
python3 main.py --weekly --dashboard

# Semana especifica
python3 main.py --week 7

# Solo una entidad
python3 main.py --entity espana
python3 main.py --entity mexico

# Reportes individuales de un mes especifico
python3 main.py --individual --report-month 1

# Combinado
python3 main.py --week 7 --entity mexico --weekly --dashboard
```

Si no se pasa ningun flag de reporte, se ejecutan todos.

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
  mexico/
    reserva.csv      <- de Corsario LLC (puede llamarse "reserva (1).csv", lo detecta solo)
    dreserva.csv     <- de Corsario LLC
```

### 3. Ejecutar

```bash
python3 main.py --week 7
```

### 4. Recoger output

- `output/Week_7_2026.xlsx` — Weekly Report, abrir en Excel para que los pivots se refresquen
- `output/espana/Bookings_SL_7_2026.xlsx` — control interno Espana
- `output/mexico/Bookings_LLC_7_2026.xlsx` — control interno Mexico
- `output/Dashboard_7_2026.xlsx` — Dashboard con graficos y analisis
- `output/Reportes_Individuales_Enero_2026/` — carpeta con un Excel por cada TA
- `output/Reporte_TAs_Febrero_2026.xlsx` — reporte mensual consolidado

### 5. Errores

Los errores detectados aparecen en consola y en la hoja ERRORES del Weekly Report. Se reportan al equipo de Madrid.

## Ejecucion automatica (GitHub Actions)

Hay 2 workflows en `.github/workflows/`:

| Workflow | Archivo | Que genera |
|----------|---------|------------|
| Weekly Report | `weekly_report.yml` | Weekly + Bookings + Dashboard |
| TA Reports | `TA's_reports.yml` | Reporte mensual TAs + Individuales |

Se pueden lanzar manualmente desde GitHub > Actions > Run workflow. Para que funcionen:
1. Subir los CSVs nuevos al repo (`git push`)
2. Descargar el output desde GitHub > Actions > Artifacts

## Estructura del proyecto

```
Gannet/
├── main.py                  <- Orquestador: python3 main.py
├── config.py                <- Configuracion (rutas, monedas, FX fallback)
├── requirements.txt         <- pandas, openpyxl, requests
│
├── src/
│   ├── data_loader.py       <- Lee CSVs y filtra canceladas
│   ├── fx_rates.py          <- Tipos de cambio historicos (API BCE)
│   ├── validators.py        <- Detecta errores en datos
│   │
│   ├── weekly/              <- Weekly Sales Report
│   │   ├── data_sheet.py    <- Construye filas de hoja DATA
│   │   ├── data_serv_sheet.py <- Construye filas de DATA SERV
│   │   ├── excel_writer.py  <- Copia template, escribe datos, formulas, pivots
│   │   ├── diferencia.py    <- Calcula Diferencia % interanual
│   │   ├── fx_sheet.py      <- Hoja FX RATES (tasas diarias BCE)
│   │   └── errores_sheet.py <- Hoja ERRORES con errores detectados
│   │
│   ├── bookings/            <- Bookings Report
│   │   ├── booking_window.py <- Matriz semana x mes en USD
│   │   └── export_bookings.py <- Export independiente por entidad
│   │
│   ├── dashboard/           <- Dashboard con graficos
│   │   └── dashboard.py     <- 10 hojas de analisis
│   │
│   ├── individual/          <- Reportes individuales por TA
│   │   └── reports.py       <- Un Excel por vendedor
│   │
│   └── ta_monthly/          <- Reporte mensual consolidado TAs
│       └── report.py        <- Pivots + SUMMARY + force refresh
│
├── data/                    <- CSVs de Corsario (tu los pones aqui)
│   ├── espana/              <- reserva.csv, dreserva.csv (+ otros de Corsario)
│   └── mexico/              <- reserva.csv, dreserva.csv (+ otros de Corsario)
│
├── templates/               <- Plantillas base (NO se modifican, solo se copian)
│   ├── Week 6.xlsx          <- Template del Weekly Report
│   ├── Reporte_TAs_template.xlsx <- Template del reporte mensual TAs
│   ├── Reporte indicidual.xlsx   <- Template de reportes individuales
│   └── plantilla_corsario.csv    <- Datos de TAs (oficina, linea negocio, comision)
│
├── tools/
│   └── create_ta_seed.py    <- Herramienta para crear template inicial de TAs
│
├── Guide/                   <- Documentacion
│   ├── README.md
│   ├── Structure
│   ├── USE
│   └── Explicacion
│
└── output/                  <- Archivos generados (no se suben a git)
```

## Hojas del Weekly Report

| Hoja | Que contiene |
|------|-------------|
| DATA | Todas las reservas activas (SL + LLC) con FX historico por fecha |
| DATA SERV | Detalle por servicio con formulas XLOOKUP |
| Weekly SL y LLC | Pivot: venta y rentabilidad por semana + Diferencia % interanual |
| Bookings Fecha de Salida | 3 pivots: ventas por semana y mes de salida |
| Booking Window 2026 | Matriz: semana de venta x mes de salida en USD |
| Por servicio y TA Weekly | Pivot: rentabilidad por vendedor y tipo de servicio |
| FX RATES | Tasas diarias del BCE del mes actual |
| ERRORES | Errores detectados en los datos |

## Hojas del Reporte Mensual TAs

| Hoja | Que contiene |
|------|-------------|
| Reportes TA {Mes} | 4 pivot tables: venta mes, YTD, por mes inicio, por mes |
| Ventas por Linea de Negocio | Pivot: ventas YTD agrupadas por Oficina > LN > Vendedor |
| SUMMARY | Datos pre-calculados en Python (funciona en Apple Numbers) |
| DATA NEW | Datos enriquecidos con plantilla corsario y FX |
| FX RATES | Tasas diarias del BCE |

## Compatibilidad multiplataforma (Reporte TAs)

| Plataforma | Pivots | SUMMARY | DATA |
|---|---|---|---|
| Excel Mac | Se refrescan al abrir | Datos completos | Completo |
| Excel Windows | Se refrescan al abrir | Datos completos | Completo |
| Apple Numbers | No soportados | Datos completos | Completo |
| Google Sheets | Se refrescan | Datos completos | Completo |

El mecanismo de force refresh usa: `fullCalcOnLoad`, `refreshOnLoad`, `recordCount=0` (fuerza mismatch), actualizacion del rango fuente, y vaciado de registros cacheados.

## Notas

- Las plantillas (`Week 6.xlsx`, `Reporte_TAs_template.xlsx`) NUNCA se modifican — el script las copia y trabaja sobre la copia.
- Los tipos de cambio son historicos: cada reserva usa la tasa del dia de su fecha de creacion (API BCE via Frankfurter).
- "GPB" es un typo de Corsario para GBP — el script lo maneja.
- El script detecta automaticamente nombres de CSV variantes (`reserva.csv`, `reserva (1).csv`, etc.).
- `.gitignore` excluye output/, __pycache__/ y .DS_Store.
