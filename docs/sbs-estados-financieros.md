# Estados financieros mensuales SBS

Dataset: `pe.sbs.estados_financieros`. Proveedor independiente para banca (B),
financieras (F), cajas municipales (C) y rurales (R). Descubre los enlaces
publicados por SBS en sus índices B-2201, B-3101, C-1101 y C-2101; no inventa
URLs de descarga. XLS y XLSX se reconocen por su contenido, incluso cuando la
extensión del enlace no coincide.

```bash
python scripts/sync_estados_financieros.py --desde 2026-08
python scripts/sync_estados_financieros.py --desde 2026-08 --hasta 2026-08 --tipos B F --load-only
```

`--data-root` selecciona la caché compartida y `--output-dir` la carpeta del
reporte. `--force` vuelve a consultar. Los archivos originales se conservan
comprimidos fuera del código, junto al manifest y los datos Parquet. Una
segunda sincronización dentro de 24 horas reutiliza la caché validada. Todos
los meses pueden revisarse porque SBS puede corregir publicaciones históricas.
Un mes ausente del índice se registra como no disponible; una respuesta
bloqueada, un libro ilegible o un cambio de esquema se registra como fallo.
El comando no exporta un rango incompleto como si estuviera completo.

```python
from fuentes_financieras import source
statements = source('pe.sbs.estados_financieros')
statements.sync(desde='2026-08', hasta='2026-08', tipos=['B', 'F', 'C', 'R'])
data = statements.load(desde='2026-08', hasta='2026-08')
```

Cada fila representa una cuenta publicada, una entidad y MN/ME/TOTAL. Los
importes de las tres columnas están expresados en **miles de soles**, incluida
ME, y `unit_multiplier=1000` permite convertir a soles. No son importes en USD.
El balance contiene saldos de cierre y los resultados son **acumulados del
ejercicio**, no flujos exclusivos del mes. Celdas vacías siguen vacías; ceros y
valores negativos se conservan.

Se mantienen cuenta y encabezado originales (`source_account`,
`source_entity_name`), hoja y fila del Excel, fecha de cierre, enlace y momento
de recuperación. `account_code` identifica solamente totales y cuentas clave
verificadas: activos, pasivos, patrimonio, pasivo y patrimonio, resultado neto,
ingresos y gastos financieros. Las demás cuentas conservan su etiqueta original
y código vacío: no se presupone que rótulos repetidos como “Otros” o
“Provisiones” sean equivalentes entre secciones. Hoja y fila permiten localizar
cada observación, pero no constituyen una identidad histórica de cuenta.

Los totales del sistema se distinguen mediante `entity_scope=system_aggregate`.
La variante bancaria con sucursales en el exterior tiene cobertura explícita
`entity_with_foreign_branches`; no debe sumarse con la entidad doméstica. Las
notas explicativas quedan en los metadatos del manifest, incluidos los cambios
institucionales que SBS anota en el libro. El formato no asigna elegibilidad
crediticia ni calcula ratios.

El libro bancario de agosto de 2026 usa “BANCOM” en activo/resultados y “Banco
de Comercio” en pasivo para la misma columna. La correspondencia se aplica
exclusivamente a ese período y grupo; el encabezado original queda conservado.
Otras discrepancias de identidad fallan hasta revisarlas con evidencia.

El lector exige fecha y unidades reconocibles, las tres secciones y totales por
entidad. Verifica activos = pasivos + patrimonio y resultado neto del balance =
resultado neto del estado de resultados, con tolerancia de 0.02 miles de soles
por redondeo. Un cambio incompatible requiere revisión del parser.

Son estados **estadísticos** publicados por SBS; su agrupación puede diferir de
los formatos contables regulatorios. Se admiten consultas desde enero de 2013
por el cambio contable señalado en los boletines, sin prometer que todos los
formatos históricos sean idénticos. La validación real inicial cubre agosto de
2026 para los cuatro grupos. La presentación de cuentas no se homogeneiza
silenciosamente entre años.
