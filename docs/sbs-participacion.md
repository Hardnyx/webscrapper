# Tamaño, ranking y participación SBS

`pe.sbs.participacion` extrae los rankings mensuales publicados de créditos
**directos**, depósitos **totales** y patrimonio. Conserva importe, posición,
participación y porcentaje acumulado; no calcula nuevas posiciones ni mezcla
entidades de grupos distintos.

| Grupo | Código del índice |
| --- | --- |
| Banca (B) | B-2332 |
| Financieras (F) | B-3243 |
| Cajas municipales (C) | C-1205 |
| Cajas rurales (R) | C-2205 |

Se descubren los enlaces oficiales desde
`https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c=CODIGO`.
No se fabrican URLs ni se sustituye un mes ausente por otro disponible.

```powershell
python scripts/sync_participacion.py --desde 2026-08
python scripts/sync_participacion.py --desde 2026-07 --hasta 2026-08 --tipos B F
python scripts/sync_participacion.py --desde 2026-08 --load-only
```

El comando admite `--data-root`, `--output-dir` y `--force`. Exporta
`outputs/participacion/participacion.xlsx`, con encabezados en español,
tabla con filtros y estilo `TableStyleLight9`, primera fila inmovilizada y
anchos originales. Exige todos los grupos y meses pedidos, también al cargar
solo caché. El ejecutor comprueba e instala únicamente dependencias faltantes
o incompatibles con el mismo intérprete.

## Bases y unidades

Los montos permanecen en miles de soles y los porcentajes en puntos
porcentuales. `published_rank` conserva la posición original, por producto y
grupo, no un ranking global del sistema financiero.

El cuadro bancario declara **«No incluye sucursales en el exterior»**; esa nota
es obligatoria para validar su alcance. No se compara automáticamente con
balances que incluyan dichas sucursales. El ranking municipal incluye **CMCP
Lima** cuando figura en el cuadro, además de las CMAC. La base de comparación
es el grupo publicado, sin filtrar por la lista actual de entidades autorizadas.

La posición, el porcentaje individual y el acumulado se preservan como datos
publicados. Un guion o blanco no se reemplaza por cero. En agosto de 2026,
Mitsui Auto Finance tiene monto de depósitos cero y porcentajes marcados `-`;
se conservan ambas observaciones, aunque su significado sea distinto.

Las notas de C/R indican que los créditos están neteados de ingresos no
devengados por arrendamiento financiero y lease-back desde enero de 2013.
Las notas completas quedan en los metadatos de caché con fila y columna.
Los nombres originales no establecen equivalencias legales entre meses.

## Contrato y controles

Cada fila identifica período, entidad, producto, posición publicada, indicador,
valor, unidad, alcance, URL, fecha de extracción y hoja/fila/columna original.
Se validan fecha de cierre, unidades, bloques, encabezados, entidades y
posiciones enteras positivas. Una secuencia de posiciones inesperada o una
inconsistencia porcentual se conserva con aviso; no se corrige la fuente.
La reconciliación de porcentajes se hace únicamente con valores completos,
sin rellenar ausentes. Un fallo no sobrescribe una captura válida.

```python
from fuentes_financieras import source

provider = source('pe.sbs.participacion')
provider.sync(desde='2026-07', hasta='2026-08', tipos=['B'], keep_raw=True)
monthly = provider.load(desde='2026-07', hasta='2026-08', tipos=['B'])
```

El consumidor puede comparar meses usando las observaciones originales;
las fusiones, cambios de nombre y cambios de grupo requieren evidencia fechada.
No se genera una identidad histórica por semejanza de nombres.

Los cuatro libros reales de agosto de 2026 contienen **369 filas**: 180 B,
54 F, 99 C y 36 R. Los archivos de validación y reportes se guardan fuera del
repositorio; las pruebas versionadas construyen libros pequeños en memoria.
La cobertura histórica completa no queda certificada por validar este mes.
