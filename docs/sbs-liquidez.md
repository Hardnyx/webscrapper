# Liquidez, cobertura y financiación neta estable SBS

Tres proveedores independientes con transporte y caché compartidos:

| Dataset | B | F | C | R |
| --- | --- | --- | --- | --- |
| `pe.sbs.liquidez` | B-2340 | B-3250 | C-1244 | C-2249 |
| `pe.sbs.cobertura_liquidez` | B-230809 | B-230810 | B-230811 | B-230812 |
| `pe.sbs.financiacion_neta_estable` | B-234021 | B-230213 | C-120212 | B-230820 |

Los enlaces se descubren en el índice publicado de cada cuadro, sin construir
URLs de descarga. Los archivos originales comprimidos y las notas se conservan
en el almacén local. No se consultan tasas, balances ni clasificaciones.

```bash
python scripts/sync_liquidez.py --desde 2026-08
python scripts/sync_liquidez.py --datasets cobertura --desde 2026-06
python scripts/sync_liquidez.py --datasets financiacion --desde 2026-07
```

El comando selecciona `liquidez` por defecto. `--datasets` admite uno o varios
bloques; todos usan el rango `--desde/--hasta` solicitado. Las fechas del ejemplo
son los últimos meses enlazados comprobados para cada bloque, no una única
fecha común. Los reportes predeterminados en `outputs/liquidez/` son `liquidez.xlsx`,
`cobertura_liquidez.xlsx` y `financiacion_neta_estable.xlsx`, según el bloque.
Una selección múltiple produce `liquidez_completa.xlsx`.

`--tipos B F C R`, `--force`, `--load-only`, `--data-root` y `--output-dir`
funcionan como en los otros extractores mensuales. Una falla o un mes sin enlace
impide exportar el rango solicitado. Un mes ausente no se sustituye por otro.
La caché verificada evita nuevas consultas durante 24 horas. Los valores
publicados ausentes se conservan como nulos con `source_value_missing`, nunca 0.

```python
from fuentes_financieras import source
provider = source('pe.sbs.cobertura_liquidez')
provider.sync(desde='2026-06', tipos=['B', 'F', 'C', 'R'], keep_raw=True)
data = provider.load(desde='2026-06', hasta='2026-06')
```

## Datos y unidades

- **Liquidez:** activos líquidos, pasivos de corto plazo y ratio de liquidez,
  por moneda y entidad, incluidos los agregados publicados. MN: miles de soles;
  ME: miles de dólares. Los ratios son porcentajes. La escala ME difiere de los
  estados financieros estadísticos que expresan sus columnas en soles.
- **RCL:** resumen de activos líquidos de alta calidad (ALAC), entradas y salidas
  a 30 días, con importes base y ajustados, en MN, ME y total. MN y total: miles
  de soles; ME: miles de dólares. El ratio publicado es el promedio de los
  ratios diarios del trimestre, en porcentaje. No se recalcula dividiendo
  promedios de saldos. No se extraen todavía todas las partidas de detalle.
- **RFNE:** financiación estable disponible y requerida, ponderadas y en miles
  de soles, y ratio total. No se extraen todavía todos los vencimientos ni las
  partidas de detalle. Se verifica el formato numérico porcentual del Excel:
  `source_value=1.25`, formato `0%`, produce `value=125`, `unit=percent`.
  Si el formato cambia a uno no reconocido, la extracción falla; no se deduce
  la escala por el tamaño del número. La muestra actual usa XLSX, aunque sus
  enlaces terminan en XLS; también se admite XLS con formato porcentual.

`value`, `unit` y `unit_multiplier` describen el valor normalizado; el
multiplicador monetario 1000 corresponde a la moneda indicada, no siempre soles.
`source_value` conserva el número almacenado y `source_number_format` el formato
verificado de los ratios de divulgación. El RCL actual almacena puntos
porcentuales con formato numérico común; un formato porcentual inesperado
requiere revisión de escala y detiene su extracción. `amount_basis` distingue base/ajustado, ponderado y promedio
de ratios diarios. No se calcula un total convirtiendo monedas.

## Períodos y cobertura

`period` es **el mes del índice SBS consultado**, también usado para particionar
la caché y seleccionar el rango. No garantiza una fecha de publicación.
`period_date` y `observation_end` son el cierre realmente declarado en el cuadro;
`observation_start` indica el inicio del mes o trimestre. `frequency` representa
la frecuencia del dato observado, no la frecuencia de descarga del índice.

En la revisión inicial:

| Bloque | Último enlace | Observación declarada | Filas B/F/C/R |
| --- | --- | --- | --- |
| Liquidez MN/ME | Agosto 2026 | Cierre agosto 2026 | 126 / 42 / 78 / 30 |
| Cobertura de liquidez | Junio 2026 | Promedio enero-marzo 2026 | 420 / 147 / 273 / 126 |
| Financiación neta estable | Julio 2026 | Cierre julio 2026 | 63 / 18 / 39 / 15 |

El RCL enlazado en junio declara enero-marzo en todas sus hojas. Se mantiene
ese trimestre y se marca `source_period_differs_from_index`. No se presenta como
RCL observado en junio ni se corrige su texto. Las entidades pueden diferir de
las del cuadro mensual más reciente: no se unen nombres ni se supone continuidad
legal. Los nombres de entidad provienen del encabezado interior, conservando
por separado la hoja original y los consolidados.

Se comprueban títulos, fechas, unidades, columnas, indicadores del resumen y
cierre de fuente. Las contradicciones numéricas en liquidez/RFNE se conservan
con `published_ratio_mismatch`, sin corregir el ratio publicado. En agosto hay
un aviso en liquidez ME de B. Efectiva: los importes redondeados no reproducen
el ratio publicado. No se imponen límites regulatorios ni reglas de elegibilidad.
La validación real inicial cubre estos doce archivos; los formatos históricos
pueden cambiar y requieren revisión explícita si fallan los contratos.
