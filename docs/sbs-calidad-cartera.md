# Calidad de cartera y provisiones SBS

Cuatro fuentes independientes, con planificación mensual y almacenamiento
compartidos. Cada proveedor consulta únicamente sus propios índices y archivos.

| Dataset | B | F | C | R |
| --- | --- | --- | --- | --- |
| `pe.sbs.calidad_cartera` | B-2401 | B-3301 | C-1301 | C-2301 |
| `pe.sbs.categorias_riesgo_cartera` | B-2309 | B-3205 | C-120201 | C-220201 |
| `pe.sbs.morosidad_dias` | B-220512 | B-3230 | C-1230 | C-2230 |
| `pe.sbs.saldos_cartera` | B-2201 | B-3101 | C-1101 | C-2101 |

```bash
python scripts/sync_calidad_cartera.py --desde 2026-08
python scripts/sync_calidad_cartera.py --desde 2026-08 --tipos B F --datasets calidad categorias
python scripts/sync_calidad_cartera.py --desde 2026-08 --load-only
```

El comando incluye los cuatro bloques por defecto y genera
`outputs/calidad_cartera/calidad_cartera.xlsx`, con una hoja por bloque, tablas,
filtros y encabezado inmovilizado. `--hasta`, `--data-root`, `--output-dir`,
`--force` y `--load-only` permiten controlar el rango, el almacén y los reportes.
Si falta un grupo/mes o falla una captura, no se exporta el rango incompleto.
No se sustituye un mes solicitado por otro. Los originales y las notas de fuente
quedan en el almacén local, excluido de git. La caché verifica sus hashes y evita
reconsultar durante 24 horas; después puede detectar revisiones de la fuente.

```python
from fuentes_financieras import source
quality = source('pe.sbs.calidad_cartera')
quality.sync(desde='2026-08', tipos=['B', 'F', 'C', 'R'], keep_raw=True)
ratios = quality.load(desde='2026-08', hasta='2026-08')
```

## Cobertura de cada bloque

**Calidad de activos:** extrae solamente esa sección del cuadro de indicadores
financieros. Incluye morosidad según criterio SBS y mayor de 90 días, cobertura
`Provisiones / Créditos Atrasados`, cartera atrasada ajustada y alto riesgo
ajustado. B publica morosidad MN/ME y refinanciados/reestructurados; F publica
refinanciados/reestructurados; C/R publican morosidad MN/ME y alto riesgo sin
ajustar. Una métrica no publicada por un grupo no se rellena ni se calcula.
Los bloques de etiquetas repetidas del Excel se validan por separado para no
perder las entidades situadas en la segunda parte del cuadro.

Los valores de esta sección son puntos porcentuales, no fracciones. B/F lo
indican en el encabezado. Para las etiquetas C/R que omiten la unidad, el
contrato revisado usa las definiciones porcentuales del
[glosario oficial SBS de enero de 2025](https://intranet2.sbs.gob.pe/estadistica/financiera/2025/Enero/SF-0002-en2025.PDF)
y la convención de los indicadores ajustados homónimos publicados por B/F.
`unit_evidence` distingue el encabezado y la definición revisada. Si una celda
cambia al formato Excel `%`, el parser detiene la captura para revisar su escala;
no multiplica según la magnitud del número ni aplica conversiones automáticas.
El glosario se revisó para definir el contrato; no se descarga implícitamente
durante cada ejecución.

**Categorías de riesgo:** conserva Normal, Con Problemas Potenciales, Deficiente,
Dudoso y Pérdida, en porcentaje, y la exposición total en miles de soles.
B/F denominan la base créditos directos e indirectos y C/R directos y
contingentes. Las notas precisan el uso del equivalente a riesgo crediticio de
los indirectos; se conserva `credit_scope` y la nota original. Ese total no se
trata como saldo exclusivamente de créditos directos ni se recalculan importes
por categoría. No se agrega la cartera pesada: su cálculo corresponde al
consumidor, usando las categorías y la base apropiadas.

**Morosidad por días:** porcentajes de créditos con **más de** 30, 60, 90 y 120
días de incumplimiento, además de morosidad según criterio contable SBS.
`arrears_threshold_days` conserva el umbral y queda nulo para el criterio
contable. Los umbrales son acumulativos; no representan tramos excluyentes.
Se conserva la nota original sobre los criterios de atraso según tipo de crédito,
sin reinterpretar los límites ni construir reglas regulatorias.

**Saldos de cartera:** créditos netos de provisiones e ingresos no devengados,
vigentes, refinanciados y reestructurados, atrasados, vencidos, cobranza judicial,
provisiones del crédito e intereses/comisiones no devengados. La selección usa
el bloque de créditos dentro del activo, evitando confundir su fila “Provisiones”
con las provisiones de inversiones, pasivos o resultados. Reutiliza el parser y
las comprobaciones contables de estados financieros; no llama a ese proveedor
ni consulta otro dataset.

Los saldos MN, ME y TOTAL están **todos en miles de soles**, según el encabezado
del balance. ME no se convierte a dólares. Las provisiones de crédito son una
cuenta que resta al activo y se conservan negativas; no se toma su valor absoluto.
El importe neto no se presenta como cartera bruta. Las notas originales sobre
neteos de ingresos no devengados se mantienen. No se calculan ratios ni se unen
los saldos con los otros bloques automáticamente.

## Avisos y trazabilidad

Cada observación conserva período de cierre, nombre y cobertura de entidad,
indicador original, unidad, URL, fecha de extracción y hoja/fila fuente.
Los tres cuadros de riesgo añaden columna original, categoría de riesgo,
cobertura crediticia y marcador textual. Los saldos conservan también cuenta
original, estado, sección, cobertura monetaria y base de medición del balance.
Los totales se distinguen de las entidades; las sucursales extranjeras incluidas
en un encabezado no se eliminan de su cobertura.

| Aviso | Tratamiento |
| --- | --- |
| `source_value_missing` | Celda vacía o guion publicado: nulo, nunca cero inferido |
| `source_definition_missing` | El indicador ajustado remite a una nota no encontrada en el archivo |
| `published_categories_sum_mismatch` | Participaciones completas no suman 100 %, tolerancia 0.02 puntos |
| `published_percentage_out_of_range` | Categoría o umbral publicado fuera de 0–100 % |
| `published_arrears_order_mismatch` | Umbrales acumulativos completos no decrecen, tolerancia 0.02 puntos |
| `published_credit_components_mismatch` | Crédito neto o atrasado no coincide con sus componentes, tolerancia 0.02 miles de soles |

Los avisos conservan el valor publicado; no lo corrigen. El guion literal queda
identificado en `source_value_token`, separado de una celda vacía. Los ceros
numéricos permanecen como ceros. Un valor no numérico desconocido, una fecha
incorrecta, una unidad cambiada, un indicador no reconocido o una estructura
incompleta producen error, no una captura exitosa.

En agosto, los indicadores ajustados C/R llevan referencias `****` y `*****`
sin sus notas de definición completas en el archivo. Se marca esa limitación.
La nota específica de Piura, situada en el segundo bloque, sí se conserva.
El indicador ajustado B/F incorpora referencias a castigos y transferencias;
no reemplaza la descarga de los flujos de castigos, que sigue pendiente.

## Validación inicial

Los índices revisados enlazan agosto de 2026 como último mes para los dieciséis
cuadros. La validación real usa el comando del repositorio y archivos fuera de git:

| Bloque | B | F | C | R | Total |
| --- | --- | --- | --- | --- | --- |
| Calidad de activos | 168 | 42 | 104 | 48 | 362 |
| Categorías de riesgo | 126 | 42 | 78 | 30 | 276 |
| Morosidad por días | 105 | 35 | 65 | 25 | 230 |
| Saldos y provisiones | 552 | 168 | 312 | 144 | 1176 |

Total: 2044 filas. Hay 27 valores ausentes o marcadores y 38 filas con definición
referenciada ausente; dos filas tienen ambos avisos. Se verificó también la
reutilización de las dieciséis capturas sin solicitudes de red.

Las listas de entidades difieren entre cuadros. Por ejemplo, calidad de activos
incluye CRAC del Centro, mientras otros cuadros de agosto no la muestran. Se
conserva lo publicado; no se deduce autorización, vigencia legal ni equivalencia
de nombres. No se incorporan límites de inversión, elegibilidad o puntajes.
Los formatos de otros años pueden diferir; el contrato inicial comprobado es
agosto de 2026 y no garantiza armonización histórica automática.
