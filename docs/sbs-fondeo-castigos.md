# Fondeo, escalas de depósitos y castigos SBS

Cinco proveedores independientes comparten transporte, planificación mensual y
caché. Cada consulta descarga únicamente su cuadro, usando enlaces publicados
por SBS. No consulta otros proveedores ni calcula decisiones de inversión.

| Dataset | Banca (B) | Financieras (F) | CMAC (C) | CRAC (R) |
| --- | --- | --- | --- | --- |
| `pe.sbs.depositos_persona` | B-2372 | B-3231 | C-1245 | C-2250 |
| `pe.sbs.depositos_escalas` | B-2321 | B-3256 | C-1211 | C-2211 |
| `pe.sbs.depositos_plazo` | B-220513 | B-3251 | Sin adapter | Sin adapter |
| `pe.sbs.adeudos` | B-2310 | B-3239 | C-1219 | C-2219 |
| `pe.sbs.castigos` | B-2369 | B-3234 | C-1253 | C-2258 |

Los índices oficiales usan
`https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c=CODIGO`.
Los archivos se descubren allí; no se fabrican direcciones ni se sustituye un
mes solicitado por el último disponible.

## Comandos

```powershell
python scripts/sync_fondeo.py --datasets personas escalas --desde 2026-08
python scripts/sync_fondeo.py --datasets adeudos --desde 2026-08 --tipos B C R
python scripts/sync_fondeo.py --datasets adeudos --desde 2026-07 --tipos F
python scripts/sync_fondeo.py --datasets plazos --desde 2026-08 --tipos B F
python scripts/sync_castigos.py --desde 2026-08
python scripts/sync_castigos.py --desde 2026-08 --load-only
```

`sync_fondeo.py` selecciona personas, escalas y adeudos por defecto. Ambos
comandos admiten `--hasta`, `--tipos`, `--data-root`, `--output-dir`, `--force`
y `--load-only`. Los ejecutores comprueban e instalan solo dependencias faltantes
o incompatibles con el mismo intérprete.

Los reportes predeterminados son `outputs/fondeo/fondeo.xlsx` y
`outputs/castigos/castigos.xlsx`. Incluyen encabezados en español, una hoja por
fuente seleccionada, tablas con filtros, `TableStyleLight9` y primera fila
inmovilizada. No modifican los anchos de columna.

Una publicación ausente o una descarga fallida impide exportar el conjunto
solicitado. `--load-only` exige que la caché cubra todos los tipos y meses
pedidos. Plazos requiere seleccionar B/F explícitamente; no omite cajas de
una solicitud B/F/C/R.

## Depósitos por persona

`depositos_persona` conserva importes por entidad, producto y persona natural,
persona jurídica sin fines de lucro u otra persona jurídica. Los cuadros B/F
incluyen vista, ahorro, plazo, CTS y total. C/R no publica una columna de vista
en este cuadro: no se inventa una observación de cero.

Todos estos importes son **miles de soles**, con monedas agregadas (`TOTAL`).
No hay desglose MN/ME en esta fuente. Los totales del sistema y los nombres que
incluyen sucursales en el exterior llevan un alcance separado.

El archivo rural de agosto de 2026 contiene una fila con identificador
numérico `0`. Se conserva como `unidentified_source_row`, sin nombre de entidad,
con `source_entity_name='0'` y aviso `source_entity_placeholder`. No se asocia a
una contraparte ni se elimina como si fuera una entidad conocida.

## Escalas de montos

`depositos_escalas` contiene **agregados del sistema**, por producto, persona y
tramo de importe. No contiene escalas individuales de cada banco o caja y no
permite calcular concentración de una contraparte ni top 10/top 20 depositantes.
Ese detalle requiere otra fuente, por ejemplo informes de clasificadoras.

Se conservan dos indicadores distintos:

| Indicador | Unidad | Alcance |
| --- | --- | --- |
| `published_number` | Número publicado (`count`) | Columna «Número» del cuadro; no se deduplican personas |
| `deposit_amount` | Miles de soles | Monto correspondiente al mismo producto, persona y tramo |

Los límites `band_lower_PEN` y `band_upper_PEN` están en **soles**, sin el
multiplicador de mil de los montos. `product_total` identifica la fila total
de cada producto; `upper_bounded` y `upper_open` identifican los tramos.
El primer límite inferior y el límite superior abierto quedan ausentes. El
texto original se conserva; la inclusión de los extremos permanece
`unspecified_by_source`, porque «de … a …» no la define explícitamente.

No se suman conteos entre productos para obtener depositantes únicos. Los
conteos totales publicados se conservan como tales. Los blancos, incluso en
vista de financieras o segmentos CTS, permanecen ausentes; no equivalen a cero.

## Monedas y plazos

`depositos_plazo` extrae los saldos publicados del Reporte 6-B para B/F. Conserva
cuenta corriente, ahorro, CTS, total y los cinco tramos de plazo: hasta 30 días,
31–90, 91–180, 181–360 y más de 360 días.

MN está en **miles de soles** y ME en **miles de dólares**. No se convierte ME a
soles ni se suman ambas monedas sin un tipo de cambio explícito del consumidor.
La definición exacta del saldo no se reconstruye como promedio diario o saldo
al cierre a partir de sus cifras.

El libro de banca de agosto de 2026 lleva fecha **1 de agosto**; el de financieras,
**31 de agosto**. `period` conserva el mes del enlace y `period_date` la fecha
original. La fecha debe pertenecer al mes solicitado, pero no se reemplaza por
el fin de mes. La banca lleva aviso `source_date_not_month_end`; las notas de
caché explican esta diferencia.

## Adeudos

`adeudos` conserva la estructura por instituciones del país e instituciones
del exterior y organismos internacionales, separando corto y largo plazo.
Las cuatro participaciones están en **puntos porcentuales**; el total de adeudos
y obligaciones financieras está en **miles de soles**. No se multiplican las
participaciones por el saldo para fabricar importes por acreedor o plazo.

El parser verifica el formato numérico Excel de porcentajes. Un nuevo formato
porcentual que altere la escala produce un error de estructura y exige revisión.
No se decide la escala por el tamaño de un número.

En la validación del 8 de octubre de 2026, el índice de financieras enlazaba
adeudos hasta **julio de 2026**, mientras B/C/R enlazaban agosto. Pedir F para
agosto devuelve `unavailable`; no reutiliza julio con una etiqueta de agosto.

## Castigos

`castigos` publica **flujos mensuales**, en miles de soles, por entidad y tipo de
crédito: corporativo, gran empresa, mediana, pequeña, microempresa, consumo,
hipotecario y total. La fuente B/F declara «en el mes de…» y C/R identifica
«flujo mensual». El parser exige esa evidencia: una cabecera acumulada o una
fecha de otro mes produce un error.

`measurement_basis='monthly_flow'` y `observation_start`/`observation_end`
definen el mes completo. No se presenta como acumulado anual ni últimos doce
meses. `-` es un marcador ausente, mientras un cero numérico se preserva como cero.

Las notas publicadas se guardan en los metadatos de caché, con fila y columna.
El cuadro de financieras advierte que los criterios de clasificación empresarial
cambiaron desde octubre de 2024 por la Resolución SBS 2368-2023. No se reasignan
las categorías anteriores ni se afirma comparabilidad histórica automática.

## Contrato y validación

Todas las filas conservan período, tipo y nombre originales, alcance, indicador,
valor, unidad, multiplicador monetario, hoja/fila/columna fuente, URL y fecha de
extracción. Los parsers rechazan encabezados desconocidos, fechas incompatibles,
columnas inesperadas y tramos discontinuos. Los valores publicados que no
reconcilian con componentes completos se preservan con aviso; no se corrigen.
La comprobación de sumas no sustituye blancos o guiones por cero.

```python
from fuentes_financieras import source

provider = source('pe.sbs.depositos_persona')
result = provider.sync(desde='2026-08', tipos=['B', 'F'], keep_raw=True)
data = provider.load(desde='2026-08', hasta='2026-08', tipos=['B', 'F'])

# This provider defaults to B/F because those are its supported groups.
source('pe.sbs.depositos_plazo').sync(desde='2026-08')
```

Se validaron con el código del repositorio **18 libros reales**: cuatro de
personas, cuatro de escalas, dos de plazos, cuatro de adeudos y cuatro de castigos.
Corresponden a agosto de 2026, salvo adeudos de financieras en julio.
También se comprobó la exportación Excel, la reutilización de las 18 particiones
sin solicitudes de red y la disponibilidad ausente de adeudos F en agosto.
Archivos, cachés, reportes y registros de esta validación permanecen fuera del
repositorio; las pruebas versionadas construyen libros pequeños en memoria.
