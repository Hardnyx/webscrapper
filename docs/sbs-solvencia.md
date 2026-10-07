# Solvencia y patrimonio efectivo SBS

Dos datasets independientes, sin consultas implícitas a otros proveedores:

| Dataset | B | F | C | R |
| --- | --- | --- | --- | --- |
| `pe.sbs.solvencia` | B-2402 | B-3302 | C-1252 | C-2257 |
| `pe.sbs.patrimonio_efectivo` | B-2370 | B-3252 | C-1257 | C-2262 |

El primero extrae requerimientos de patrimonio efectivo por riesgo de crédito,
mercado y operacional, los APR correspondientes y su total, capital ordinario
de nivel 1 / APR, patrimonio efectivo de nivel 1 / APR y ratio de capital global.
APR significa activos y contingentes ponderados por riesgo. Los requerimientos
no son el patrimonio efectivo disponible.

El segundo extrae capital ordinario de nivel 1, capital adicional de nivel 1,
patrimonio efectivo de nivel 2 y patrimonio efectivo total. No calcula los
importes a partir de ratios ni los une automáticamente con los balances.

```bash
python scripts/sync_solvencia.py --desde 2026-07
python scripts/sync_solvencia.py --desde 2026-07 --hasta 2026-07 --tipos B F --load-only
```

`--data-root` y `--output-dir` seleccionan caché y reportes. El comando sincroniza
los dos datasets explícitamente y exporta `outputs/solvencia/solvencia.xlsx` con
tablas, filtros y dos hojas. Si falta un grupo/mes o falla un parser, no exporta
un rango incompleto. `--force` vuelve a consultar. La caché valida hashes y evita
nuevas descargas durante 24 horas; todos los períodos pueden revisarse después.
Los originales comprimidos y sus notas se conservan en la carpeta de datos.

```python
from fuentes_financieras import source
solvency = source('pe.sbs.solvencia')
solvency.sync(desde='2026-07', tipos=['B', 'F', 'C', 'R'])
ratios = solvency.load(desde='2026-07', hasta='2026-07')
capital = source('pe.sbs.patrimonio_efectivo').load(desde='2026-07', hasta='2026-07')
```

Se conserva una fila por entidad e indicador: valor, unidad, multiplicador,
período de cierre, hoja/fila original, encabezado del indicador, URL, fecha de
extracción y avisos de calidad. Los totales del sistema son agregados separados;
la entidad bancaria con sucursales en el exterior conserva esa cobertura.
La fecha auxiliar del archivo se guarda separada del período observado; no se
interpreta como una nueva observación ni como fecha de publicación garantizada.
Los ratios publicados, por ejemplo 17.12, significan 17.12 %, no 0.1712.

El acceso y la muestra inicial verifican julio de 2026 para los ocho cuadros.
En esta revisión los índices de solvencia enlazan julio como último mes, aunque
el dataset de estados financieros ya publica agosto. No se traslada el dato a
agosto ni se completa un mes ausente con el anterior.

## Unidades y avisos de la fuente

| Cuadro | Unidades conservadas |
| --- | --- |
| Requerimientos y APR B/F/C/R | Miles de soles, multiplicador 1000 |
| Ratios de capital B/F/C/R | Porcentaje |
| Patrimonio efectivo B/F | Miles de soles, multiplicador 1000 |
| Componentes del patrimonio C/R | Porcentaje, según encabezado |
| Total del patrimonio C/R | Unidad no especificada separadamente; sin multiplicador |

Los archivos C/R declaran “En porcentaje” para la composición, pero contienen
una columna total con valores de apariencia monetaria y sin unidad propia.
Se conserva el número original con `unit=unspecified_by_source` y
`source_unit_unspecified`; su escala no se deduce de su magnitud. Esa limitación
impide convertir directamente estos totales a soles. Las participaciones
tampoco se convierten automáticamente a importes.

La muestra municipal presenta composiciones que no suman 100 % para Ica,
Paita y Piura. Se mantienen y marcan como
`published_components_sum_mismatch`, sin corregir el dato publicado. El reporte
traduce estos avisos a español y el comando comunica su presencia. Un archivo
capturado correctamente puede contener datos publicados con estas limitaciones;
los avisos no equivalen a una validación económica completa de la fuente.

Se exige fecha efectiva correcta, encabezados reconocibles, unidades declaradas,
entidades y agregado, totales esenciales y coincidencia del APR total con sus
componentes (tolerancia 0.02 miles de soles). Una respuesta HTML, un bloqueo o un
libro cambiado no se considera una descarga válida.

Los formatos anteriores pueden diferir: la muestra comprobada es julio de 2026,
y las notas SBS señalan cambios de composición desde enero de 2023. No se
armonizan silenciosamente los regímenes históricos ni se fijan límites
regulatorios o reglas de elegibilidad dentro del scraper.
