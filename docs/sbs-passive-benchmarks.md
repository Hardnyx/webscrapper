# Referencias de tasas pasivas

Dataset general: `pe.sbs.tasas_pasivas_mercado`.

Fuente: https://www.sbs.gob.pe/app/pp/EstadisticasSAEEPortal/Paginas/TIPasivaMercado.aspx?tip=B

Metodología: https://www.sbs.gob.pe/app/stats/metodologia/metodologia_ti_promedio.pdf

| Indicador | Moneda | Base | Ventana | Grupo según metodología SBS |
| --- | --- | --- | --- | --- |
| TIPMN | MN | Saldos | Saldos vigentes | Bancos y financieras |
| TIPMEX | ME | Saldos | Saldos vigentes | Bancos y financieras |
| FTIPMN | MN | Flujos | Últimos 30 días útiles | Bancos |
| FTIPMEX | ME | Flujos | Últimos 30 días útiles | Bancos |

Las tasas se conservan en porcentaje efectivo anual: `1.97` significa 1.97 %,
no 197 %. ME conserva la denominación de moneda extranjera publicada por SBS.
No se interpreta automáticamente como una serie exclusiva de dólares.

## Ejecutar

```powershell
python scripts/sync_passive_benchmarks.py --desde 2026-10-05 --hasta 2026-10-06
python scripts/sync_passive_benchmarks.py --load-only
```

El lanzador comprueba las dependencias. `--data-root PATH` selecciona el almacén
y `--output-dir PATH` selecciona el reporte. `--force` revalida las fechas.
Se utiliza HTTP y WebForms, sin navegador. Ambas fechas publicadas en los
bloques de saldos y flujos deben coincidir con la solicitada; no basta la fecha
del selector. Una respuesta parcial, una tasa ilegible o un bloqueo son errores.
Los períodos no disponibles no se sustituyen por otra fecha.

El caché conserva particiones por año, HTML y manifiestos verificados por hash.
La planificación omite fines de semana; no incorpora un calendario de feriados.
Un rango incompleto termina con código 1 y no exporta un reporte completo.

```python
from fuentes_financieras import source
market = source('pe.sbs.tasas_pasivas_mercado')
market.sync(desde='2026-10-05', hasta='2026-10-06')
general = market.load(metric=['TIPMN', 'FTIPMN'])
```

Contrato: `period`, `period_date`, `frequency`, `metric`, `currency`, `rate`,
`unit`, `basis`, `observation_window`, `entity_scope`, `reference_kind`, `source`,
`source_url`, `methodology_url`, `retrieved_at`.

## Promedios por producto

```python
from fuentes_financieras.benchmarks import product_benchmarks
products = product_benchmarks(source('pe.sbs.tasas_pasivas').load())
```

Esta función extrae las filas **Promedio** publicadas por SBS en el dataset de
tasas por entidad. No calcula una media simple de entidades y no descarga datos.
El proveedor de mercado no consulta el proveedor de tasas por entidad.

Conserva tipo B/F/C/R, moneda, cuadro, persona, producto o tramo, fecha,
frecuencia y fuentes. B/F corresponden a flujos de los últimos 30 días útiles;
C/R, a flujos del mes calendario. Una tasa vacía sigue vacía: no se convierte
en cero ni se imputa. Promedios duplicados para la misma clave son un error.

El reporte `outputs/passive_benchmarks/referencias_tasas_pasivas.xlsx` contiene
las referencias generales del rango solicitado y todos los promedios por
producto presentes en el almacén local, con sus períodos explícitos. No recorta
los promedios mensuales al rango diario ni los presenta como datos de ese día.
Si faltan tasas locales, se informa que los promedios no fueron evaluados.
Las hojas con datos tienen tablas, filtros, estilo claro 9 y encabezado fijo.

## Comparabilidad

Para comparar un depósito con un promedio por producto, el consumidor debe
igualar tipo de entidad, moneda, producto o tramo, tipo de persona, período,
frecuencia, unidad y ventana. Un promedio general de todos los depósitos no
equivale al promedio de depósitos de personas jurídicas a 181–360 días.
Una referencia de banca tampoco equivale a una referencia de cajas rurales.
No se interpolan fechas ni se calculan spreads o reglas de elegibilidad aquí.
