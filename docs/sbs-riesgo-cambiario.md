# Riesgo cambiario SBS

Dos proveedores independientes conservan los cuadros publicados, sin calcular un score ni conectar automáticamente con otras fuentes.

| Dataset | Cobertura | Cuadro SBS | Unidad |
| --- | --- | --- | --- |
| `pe.sbs.posicion_cambiaria` | B/F/C/R | B-2368 / B-3266 / C-1260 / C-2370 | Miles de soles |
| `pe.sbs.posicion_cambiaria_capital` | C/R | C-1301 / C-2301, sección de moneda extranjera | Porcentaje publicado |

## Posición global

Se conservan cuatro componentes por entidad y agregado publicado: posición de cambio de balance, posición neta en derivados, delta de posiciones netas en opciones y posición global. La fuente declara miles de soles para todos los componentes, aunque la exposición sea en moneda extranjera (`currency=ME`). No son importes en dólares. `unit_multiplier=1000` permite pasar de miles de soles a soles.

Los signos originales se mantienen. Un guion o celda vacía permanece como valor ausente con su aviso; no se sustituye por cero. Solo si están presentes los cuatro valores se comprueba la igualdad global = balance + derivados + delta, con tolerancia absoluta de 0,02 miles de soles. Una discrepancia genera un aviso y conserva los valores publicados.

La fuente es el Reporte 2-B1, Anexo 3, requerimiento de patrimonio efectivo por riesgo cambiario, método estándar. Las notas originales se guardan con la captura.

## Ratio sobre patrimonio efectivo

Los indicadores de cajas publican `Posición Global en M.E. / Patrimonio Efectivo (%)`. La nota establece que el patrimonio efectivo corresponde al mes anterior. Cada observación conserva `period` y `period_date` del cuadro y `denominator_period` y `denominator_date` del cierre del mes anterior, incluso al cambiar de año.

El ratio se guarda en puntos porcentuales, incluyendo valores negativos o fuera de 0–100. No se calcula a partir del proveedor de solvencia ni se deduce un importe de exposición. Los formatos Excel que impliquen otra escala porcentual, las unidades ausentes y la pérdida de la nota temporal hacen fallar la captura. B/F no están cubiertos por este dataset; la CLI exige seleccionar C/R explícitamente.

## Uso

```bash
# Posiciones para los cuatro grupos; selección predeterminada.
python scripts/sync_riesgo_cambiario.py --desde 2026-07

# Ratio publicado de cajas, con denominador del mes anterior.
python scripts/sync_riesgo_cambiario.py --datasets capital --desde 2026-08 --tipos C R

# Reporte de un rango ya disponible en caché.
python scripts/sync_riesgo_cambiario.py --desde 2026-07 --load-only
```

`--hasta`, `--force`, `--data-root` y `--output-dir` funcionan como en los demás cuadros mensuales. El reporte predeterminado es `outputs/riesgo_cambiario/riesgo_cambiario.xlsx`; distintas ejecuciones sustituyen ese reporte, por lo que conviene usar carpetas distintas para selecciones que se quieran conservar. También está disponible `fuentes-sync-fx-risk` tras instalar el paquete.

```python
from fuentes_financieras import source

positions = source('pe.sbs.posicion_cambiaria')
positions.sync(desde='2026-07', tipos=['B', 'F', 'C', 'R'])
amounts = positions.load(desde='2026-07', hasta='2026-07')

ratios = source('pe.sbs.posicion_cambiaria_capital')
ratios.sync(desde='2026-08', tipos=['C', 'R'])
capital_ratio = ratios.load(desde='2026-08', hasta='2026-08', tipos=['C', 'R'])
```

La CLI no exporta un rango incompleto ni reemplaza un mes solicitado por otro. Cambios de estructura, fecha, unidad o tipo de entidad se registran como fallos, conservando una captura válida anterior. Una captura vigente en caché se reutiliza sin solicitudes de red; las revisiones se rigen por la política compartida de actualización mensual.

## Validación

Validado el 9 de octubre de 2026 con el código del repositorio: posiciones de julio de 2026 para B/F/C/R y ratio de agosto de 2026 para C/R, cuyo denominador corresponde a julio. Los índices consultados ofrecían julio como último mes de los cuadros dedicados de posición. Esta validación no certifica todo el histórico ni equipara los meses de ambas fuentes. Descargas y reportes de comprobación se mantienen fuera del repositorio.
