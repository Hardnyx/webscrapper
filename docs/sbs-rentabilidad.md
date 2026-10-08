# Rentabilidad, eficiencia y gestión SBS

Dos proveedores independientes extraen secciones del cuadro mensual de
indicadores financieros. Comparten utilidades de lectura y transporte, pero
ninguno consulta ni recalcula resultados del otro.

| Dataset | Sección | Banca (B) | Financieras (F) | CMAC (C) | CRAC (R) |
| --- | --- | --- | --- | --- | --- |
| `pe.sbs.rentabilidad` | Rentabilidad | B-2401 | B-3301 | C-1301 | C-2301 |
| `pe.sbs.eficiencia` | Eficiencia y gestión | B-2401 | B-3301 | C-1301 | C-2301 |

Se usan únicamente los enlaces publicados en
`https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c=CODIGO`.
Una publicación ausente no se sustituye por la de otro mes.

## Uso

```powershell
python scripts/sync_rentabilidad.py --desde 2026-08
python scripts/sync_rentabilidad.py --desde 2026-08 --datasets rentabilidad --tipos B F
python scripts/sync_rentabilidad.py --desde 2026-08 --datasets eficiencia --tipos C R
python scripts/sync_rentabilidad.py --desde 2026-08 --load-only
```

El comando solicita ambos datasets por defecto y admite `--hasta`, `--tipos`,
`--data-root`, `--output-dir`, `--force` y `--load-only`. El ejecutor comprueba
versiones e instala solo paquetes faltantes o incompatibles con el mismo
intérprete de Python.

El reporte predeterminado es
`outputs/rentabilidad_eficiencia/rentabilidad_eficiencia.xlsx`, con una hoja por
dataset, encabezados en español, tablas con filtros, estilo `TableStyleLight9`
y primera fila inmovilizada. No modifica los anchos de columna.

La exportación exige cubrir todos los grupos y meses solicitados. Un fallo de
descarga, error de estructura o publicación ausente impide generar el conjunto
solicitado. `--load-only` también comprueba la cobertura.

```python
from fuentes_financieras import source

profitability = source('pe.sbs.rentabilidad')
result = profitability.sync(desde='2026-08', tipos=['B', 'F'], keep_raw=True)
data = profitability.load(desde='2026-08', hasta='2026-08', tipos=['B', 'F'])
```

## Rentabilidad

| Indicador | Significado publicado | Unidad |
| --- | --- | --- |
| `return_on_equity` | Utilidad anualizada sobre patrimonio promedio, ROE | Puntos porcentuales |
| `return_on_assets` | Utilidad anualizada sobre activo promedio, ROA | Puntos porcentuales |

B/C/R identifica utilidad **neta**; F usa el rótulo «utilidad». Se conserva
`numerator_basis='annualized_profit_as_labeled'` para F, sin añadir «neta» al
rótulo original. Se conserva la base del numerador y del denominador en cada fila.

La nota publicada define el valor anualizado como el valor del mes más el valor
a diciembre del año anterior menos el del mismo mes del año anterior. El
promedio corresponde a los últimos doce meses. El parser exige esta definición
explícita y no anualiza multiplicando una utilidad mensual por doce.

Por ejemplo, agosto de 2026 lleva `measurement_basis='rolling_12_months'`, con
ventana del 1 de septiembre de 2025 al 31 de agosto de 2026. La observación sigue
siendo mensual: no representa únicamente el resultado de agosto. La definición
publicada se conserva en `annualization_definition` y en las notas de caché.

## Eficiencia y gestión

| Contenido | B/F | C/R |
| --- | --- | --- |
| Gastos administrativos anualizados | Sobre activo productivo promedio | Sobre créditos directos e indirectos promedio |
| Gastos operativos y margen financiero total | Rótulo sin anualización explícita | Ambos términos anualizados explícitamente |
| Ingresos financieros anualizados | Sobre activo productivo promedio | Sobre activo productivo promedio |
| Productividad del personal | Créditos directos por personal | Créditos directos por empleados |
| Productividad de oficinas | Depósitos por número de oficinas | Créditos directos por número de oficinas |
| Indicador adicional | Ingresos financieros sobre ingresos totales | Depósitos sobre créditos directos |

Las diferencias de denominador y anualización llevan identificadores distintos
cuando corresponde. Los ratios B/F de gastos operativos y de composición de
ingresos quedan con `measurement_basis='unspecified_by_source'` y sin ventana
inferida. No se los etiqueta como últimos doce meses por estar junto a un ratio
anualizado.

Los indicadores con anualización explícita conservan la ventana de doce meses.
Los ratios de saldos y productividad se identifican como `point_in_time`, con
fecha del cuadro. No se recalculan ratios usando los estados financieros ni se
fabrican importes absolutos de gastos o margen. Esos saldos contables se obtienen
del proveedor independiente de [estados financieros](sbs-estados-financieros.md).

## Unidades, nombres y controles

Los ratios se conservan en puntos porcentuales, sin dividirlos entre cien.
Los importes por persona, empleado u oficina están en miles de soles por su
respectivo denominador, con unidades `thousands_PEN_per_person`,
`thousands_PEN_per_employee` y `thousands_PEN_per_office`. El multiplicador de mil
no elimina el denominador ni convierte el dato en un importe total.

Los valores negativos de rentabilidad y los ratios superiores a 100 se
preservan. No se les aplica un límite general de 0–100, porque estos cocientes
no son necesariamente participaciones. Los blancos y el marcador `-` son
valores ausentes, mientras que un cero numérico se conserva como cero.

Los nombres se mantienen como aparecen en el cuadro, incluyendo asteriscos,
notas y sucursales en el exterior. Los agregados del sistema se distinguen de
las entidades individuales. El encabezado bancario de agosto de 2026 contiene
«HSBC Bank Perú» en una fila auxiliar de la columna de «B. Falabella Perú»:
se guarda en `source_auxiliary_entity_name` con aviso, sin reemplazar el nombre
principal ni establecer una equivalencia o sucesión legal.

El parser comprueba tipo de entidad, fecha de cierre, bloques repetidos,
secciones, indicadores completos, encabezados y definición de anualización.
Rechaza columnas con datos sin entidad, valores sin indicador, indicadores
nuevos o duplicados y cambios de formato Excel que alteren la escala porcentual.
Una captura fallida no sobrescribe una versión válida de la caché.

Cada fila mantiene hoja, fila, columna, URL original y fecha de extracción. Las
notas publicadas se conservan con posición original en los metadatos de caché.
La validación cubre los cuatro libros publicados para agosto de 2026; no
certifica que los rótulos de todo el histórico permanezcan iguales.

## Validación

Se comprobaron descargas reales con el código del repositorio para los dos
datasets y B/F/C/R: **94 filas de rentabilidad y 282 de eficiencia**. También
se verificaron los reportes Excel y la reutilización de ocho particiones de
caché sin solicitudes de red. Los archivos reales, registros y reportes de
validación están fuera del repositorio. Las pruebas versionadas construyen
libros pequeños en memoria y comprueban fechas, bases, unidades, valores
ausentes, ventanas de doce meses, captura fallida y cobertura de exportación.
