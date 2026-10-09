# Clasificaciones e inventario de informes SBS

Fuente: https://www.sbs.gob.pe/app/iece/paginas/MostrarResumenClasificaciones.aspx

| Dataset independiente | Contenido | Contrato |
| --- | --- | --- |
| `pe.sbs.clasificaciones_riesgo` | Clasificación institucional del resumen, cambio publicado y referencia al informe | v3 |
| `pe.sbs.informes_riesgo` | Inventario de enlaces y versiones por entidad, período y clasificadora | v1 |

## Alcance de la información

El resumen conserva la clasificación tal como aparece en la SBS; no equivale a un rating de depósitos a corto o largo plazo. La letra no se convierte a una escala numérica común entre clasificadoras.

`trend=up/down` identifica únicamente los símbolos `SUBIO.png` y `BAJO.png`, que la fuente describe como cambios respecto de la clasificación anterior. No es una perspectiva estable, positiva o negativa. Sin símbolo, `trend` queda ausente; no se interpreta como estabilidad. Los símbolos desconocidos se conservan mediante URL, título y texto alternativo, con un aviso para revisión. Un símbolo conocido con título contradictorio hace fallar la captura.

`period_code=YYYY01/YYYY02` representa el corte del resumen de marzo/septiembre. `period_date` es el cierre de ese mes derivado del código, no la fecha del comité de clasificación ni la publicación del PDF. En el inventario, `report_date` queda vacío hasta que exista evidencia en el documento.

Cada referencia conserva `report_url`, `report_agency_code`, `report_period_code`, `report_file_number`, `report_version` y `report_id`, compuesto por los cuatro identificadores SBS. Se exige HTTPS, el destino oficial y parámetros únicos, numéricos y del período solicitado. Una URL ausente conserva la clasificación con aviso; un enlace de otro período o destino desconocido genera un fallo. Un mismo documento puede estar asociado a varias entidades: no se deduplican sus asociaciones por URL.

El inventario no descarga ni analiza los documentos. `document_status=linked_not_downloaded` indica únicamente que la SBS publicó un enlace, no que se haya comprobado la disponibilidad o el contenido del archivo. Los ratings de depósitos, las perspectivas y las fechas de comité se pueden extraer mediante el módulo independiente de [documentos de riesgo](sbs-documentos-riesgo.md), dentro de sus formatos y alcance comprobados.

## Ejecutar

```bash
# Clasificaciones de todos los períodos publicados; selección predeterminada.
python scripts/sync_clasificaciones_riesgo.py

# Ambos datasets para períodos explícitos.
python scripts/sync_clasificaciones_riesgo.py --datasets clasificaciones informes --periodos 202602 202601 --keep-raw

# Solo el inventario dentro de un rango de cortes publicados.
python scripts/sync_clasificaciones_riesgo.py --datasets informes --desde 2025-03 --hasta 2026-09

# Exportar una selección ya validada, sin consultar la página.
python scripts/sync_clasificaciones_riesgo.py --datasets clasificaciones informes --periodos 202602 202601 --load-only
```

El lanzador comprueba las dependencias compatibles. El proveedor utiliza `curl_cffi`, WebForms y HTML, sin navegador. El cliente reutiliza el HTML inicial; otros períodos requieren una consulta y actualizan el estado de la página. Las respuestas parciales se analizan por longitudes UTF-16 para conservar barras verticales, saltos de línea y caracteres suplementarios dentro del contenido. Respuestas truncadas, errores de servidor, paneles inesperados y estado duplicado se rechazan.

`--periodos` no se combina con `--desde/--hasta`. `--force` vuelve a consultar la selección; `--keep-raw` conserva las respuestas comprimidas; `--no-second-sync` omite la comprobación adicional de caché. `--load-only` exige capturas íntegras bajo el contrato actual y no puede combinarse con `--force`.

## Almacén y reportes

Los datasets tienen almacenes y manifiestos separados bajo `data/sources/peru/sbs`; `--data-root PATH` cambia la raíz. La versión 3 reconsulta los períodos guardados con contratos anteriores y los sustituye al validar el nuevo esquema; no requiere borrar el almacén. Un fallo de estructura conserva la captura válida previa. Para completar la migración del histórico hay que sincronizar todos sus períodos.

La CLI comprueba la cobertura de la selección y, salvo `--no-second-sync`, exige que la segunda sincronización reutilice todas las capturas. No exporta una selección fallida o incompleta. Los tipos históricos sin correspondencia en el selector actual mantienen su etiqueta y un código ausente; no se establecen equivalencias legales por nombre.

El reporte es `outputs/clasificaciones_riesgo/clasificaciones_informes.xlsx`, con una hoja por dataset, encabezados en español, tablas con filtros y primera fila fija. `--output-dir PATH` cambia la carpeta. Desde la versión 0.14.0 sustituye las salidas antiguas CSV/JSON de validación de esta CLI; los manifiestos del almacén siguen disponibles para auditoría. No se generan archivos de ejemplo ni datos de prueba en el repositorio.

## API

```python
from fuentes_financieras import source

ratings = source('pe.sbs.clasificaciones_riesgo')
ratings.sync(periodos=['202602', '202601'])
history = ratings.load(periodos=['202602', '202601'], tipo_entidad='B')

reports = source('pe.sbs.informes_riesgo')
reports.sync(periodos=['202601'])
inventory = reports.load(periodos=['202601'], entidad='BANCO')
```

Ambos proveedores permiten filtros por tipo de entidad, entidad y clasificadora. Los nombres originales se mantienen. Los datasets no establecen elegibilidad ni rankings de calidad crediticia.

## Validación realizada

El 9 de octubre de 2026 se validaron con el código del repositorio los 30 períodos del resumen publicados desde marzo de 2012 hasta septiembre de 2026: 4.215 clasificaciones con referencias completas y sin avisos de enlace o símbolo desconocido. Se comprobó la reutilización de caché del histórico.

El inventario independiente se sincronizó para septiembre de 2025, marzo de 2026 y septiembre de 2026: 420 asociaciones, con exportación Excel y lectura sin red. Esta comprobación no certifica la disponibilidad de cada documento enlazado. Las pruebas cubren estructura ambigua, duplicados, período de enlace incorrecto, versiones, escala temporal, respuestas parciales, migración de esquema y conservación de capturas tras un fallo.
