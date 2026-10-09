# Documentos de clasificación con evidencia

Dataset independiente: `pe.sbs.documentos_riesgo`.

El proveedor recibe URLs oficiales explícitas. Descarga PDF completos, conserva sus versiones y extrae campos de la portada mediante reglas para formatos comprobados. La CLI consume expresamente `pe.sbs.informes_riesgo`; los proveedores no se sincronizan entre sí automáticamente.

## Uso

```bash
# Sincronizar el inventario del período y descargar los informes de una entidad.
python scripts/sync_documentos_riesgo.py --periodos 202601 --entidad ALFIN

# Limitar también por clasificadora.
python scripts/sync_documentos_riesgo.py --periodos 202601 --entidad BANBIF --clasificadora Moodys

# Exportar inventario, campos y PDF ya validados sin acceder a la red.
python scripts/sync_documentos_riesgo.py --periodos 202601 --entidad ALFIN --load-only

# Volver a analizar PDF locales íntegros.
python scripts/sync_documentos_riesgo.py --periodos 202601 --entidad ALFIN --force

# Comprobar de nuevo los archivos de los enlaces seleccionados.
python scripts/sync_documentos_riesgo.py --periodos 202601 --entidad ALFIN --redownload
```

Sin filtros por entidad/clasificadora/tipo, la CLI procesa todos los enlaces de los períodos indicados. `--tipo-entidad` acepta el código o etiqueta del inventario. `--data-root` cambia la raíz de datos y `--output-dir` la carpeta del reporte. Los filtros por nombre seleccionan asociaciones publicadas; no determinan identidad legal.

El lanzador verifica e instala solo dependencias ausentes o incompatibles, incluido `pypdf>=6,<7`. Para usar el comando instalado `fuentes-sync-risk-documents`, instale previamente el extra PDF con `python -m pip install -e ".[pdf]"`.

```python
from fuentes_financieras import source

reports = source('pe.sbs.informes_riesgo')
reports.sync(periodos=['202601'])
refs = reports.load(periodos=['202601'], entidad='ALFIN')

documents = source('pe.sbs.documentos_riesgo')
result = documents.sync(urls=refs.report_url.drop_duplicates().tolist())
fields = documents.load(report_ids=refs.report_id.tolist())
```

`fetch(url=...)` devuelve campos sin escribir la captura canónica, aunque conserva el PDF validado. `sync(urls=...)` valida y registra las capturas; URLs repetidas del mismo identificador se descargan una sola vez. Una nueva versión tiene otro identificador y otro archivo. `sync(force=True)` vuelve a analizar archivos íntegros existentes; `sync(redownload=True)` vuelve a solicitarlos.

## Evidencia y alcance

Cada registro conserva `report_id`, URL y cuatro identificadores SBS, `page_number` de base uno, `evidence_text`, `value_raw`, `normalized_value`, `temporal_role`, `extraction_status` y `pdf_sha256`. Las fechas completas válidas se normalizan a ISO; las escalas de rating y sus prefijos locales se conservan sin convertirlas a una escala común.

| Formato comprobado | Campos de portada reconocidos |
| --- | --- |
| Apoyo y Asociados, código 000408 | Tabla Actual/Anterior de fortaleza y depósitos con etiquetas reconocidas; perspectiva separada; fechas plurales de comité conservadas sin asignación temporal automática |
| JCR Latino America, código 001196 | Tabla Actual/Anterior de fortaleza y depósitos; perspectiva; fecha de comité asociada a la nota de información actual |
| PCR, código 000409 | Tarjetas de fortaleza y depósitos antes de las definiciones; perspectiva separada; fecha de comité rotulada |
| Moodys Local PE, código 000406 | Tabla de clasificaciones actuales de entidad, emisor y depósitos de corto plazo; perspectivas de entidad/emisor cuando aparecen; fechas de comité y publicación rotuladas |

Estas reglas describen estructuras concretas, no garantizan cobertura de todos los informes de cada clasificadora. La extracción se limita a la portada, aunque se comprueba la legibilidad y presencia de texto de todas las páginas. No extrae ratings desde párrafos narrativos, escalas explicativas o tablas históricas. La ausencia de un campo no implica que el producto carezca de rating.

`temporal_role=current/previous` se refiere a la posición en el bloque del informe, no a su vigencia hoy. `period_code` procede del enlace al resumen SBS y no reemplaza la fecha de comité ni la de publicación. Las fechas plurales sin correspondencia explícita mantienen `temporal_role=unspecified` y `needs_review`; no se asignan ordenándolas cronológicamente.

Las perspectivas separadas de Apoyo/JCR/PCR se conservan como perspectivas del informe; no se propagan automáticamente a cada producto. En Moodys las perspectivas de entidad y emisor permanecen separadas, y un guion junto al rating de depósitos no se interpreta como perspectiva estable.

Las reglas solo aceptan valores de la estructura reconocida. Varias observaciones del mismo campo y rol se conservan y marcan para revisión. Una portada sin campos reconocidos produce `unsupported_cover`; un documento sin texto produce `needs_ocr`. No se realiza OCR automático ni se afirma que estos documentos hayan quedado estructurados.

## Descarga, caché y errores

Los PDF se guardan en `data/sources/peru/sbs/documentos_riesgo/documents`, usando los identificadores de clasificadora, período, número y versión. El almacén canónico y su manifiesto siguen separados del inventario.

La descarga utiliza TLS verificado y un plazo de 180 segundos. Se exige una URL oficial HTTPS y se rechazan redirecciones finales fuera del destino. Se comprueban la cabecera PDF, el marcador de cierre, la estructura mediante lectura estricta, el cifrado y los límites de 20 MB/200 páginas. Una respuesta HTML o una transferencia incompleta no se guarda como PDF válido.

El archivo se sustituye de forma atómica solo después de validarlo. La caché exige que la huella del PDF y los datos canónicos coincidan con el manifiesto. Un PDF ausente o corrupto se vuelve a solicitar durante `sync`; `--load-only` lo rechaza. Un fallo de red o estructura conserva la captura válida anterior. La lectura no convierte estos fallos en ausencia de rating.

## Reporte y códigos de salida

`outputs/documentos_riesgo/documentos_riesgo.xlsx` contiene campos y asociaciones publicadas con tablas, filtros y encabezados en español. La asociación de entidad/clasificadora se incorpora mediante una unión explícita por `report_id`, sin afirmar que se haya verificado la identidad legal dentro del documento.

- **0:** documentos seleccionados íntegros y todos los campos emitidos extraídos del formato reconocido; no certifica extracción exhaustiva del documento.
- **1:** descarga, inventario o caché incompletos/dañados; no se exporta un reporte nuevo.
- **2:** reporte exportado con campos o documentos pendientes de revisión/extracción.

## Validación realizada

El 9 de octubre de 2026 se descargaron y revisaron visualmente cuatro informes reales enlazados desde el resumen 202601, uno por formato descrito. La extracción produjo 26 registros: 24 de bloques reconocidos y dos fechas de comité con asignación temporal pendiente.

Se comprobó el proveedor con esos bytes reales, una descarga directa adicional a través del proveedor, reutilización de caché sin red y exportación de tres selecciones desde caché. Las pruebas cubren PDF incompletos, HTML, cifrado, ausencia de texto, formatos desconocidos, versiones, fechas inválidas, corrupción de caché y conservación de archivos tras una descarga fallida. Esto no valida todo el histórico de PDF ni todas las entidades. Descargas, imágenes y reportes de prueba permanecen fuera del repositorio.
