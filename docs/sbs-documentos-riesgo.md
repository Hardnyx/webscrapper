# Documentos de clasificación con evidencia

Dataset independiente: `pe.sbs.documentos_riesgo`.

El proveedor recibe URLs oficiales explícitas. Descarga PDF completos, conserva sus versiones y extrae campos de portada y concentración mediante reglas para formatos comprobados. Revisa el texto de todas las páginas para localizar evidencia cualitativa candidata. La CLI consume expresamente `pe.sbs.informes_riesgo`; los proveedores no se sincronizan entre sí automáticamente.

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
| MicroRate, código 000410 | Bloque de calificación crediticia y perspectiva; no se convierte en fortaleza financiera ni en rating de depósitos |

Estas reglas describen estructuras concretas, no garantizan cobertura de todos los informes de cada clasificadora. Los ratings se extraen únicamente de la portada. No extrae ratings desde párrafos narrativos, escalas explicativas o tablas históricas. La ausencia de un campo no implica que el producto carezca de rating.

### Concentración y evidencia del cuerpo

Las reglas de concentración reconocen estructuras fechadas comprobadas: el párrafo de JCR con los 10 y 20 principales depositantes y el paréntesis de Moodys con los principales depositantes. Cada cifra conserva `top_depositors`, `unit=percent`, `observation_period` (mes publicado), `denominator_basis`, página y pasaje. En JCR se reconoce el total de depósitos como denominador; en el paréntesis de Moodys permanece `unspecified_in_excerpt` porque no está explicitado allí. No se completa una fecha ausente ni se asocia un porcentaje cercano por aproximación.

Se localizan hasta dos pasajes candidatos por documento para cada tema: estrategia, accionistas y soporte, costo de fondeo, factores de riesgo, menciones de eventos y concentración. Conservan texto y página, con `needs_review` y valor normalizado vacío. Las menciones pueden contener negaciones o explicaciones generales: **no constituyen eventos confirmados ni alertas automáticas**.

La cobertura por tema informa cuántas cifras y candidatos se encontraron, páginas revisadas y páginas con texto. `not_found_in_text` significa que estas reglas no encontraron evidencia; no demuestra que el dato no exista. Un documento sin texto requiere OCR. No se realiza extracción cualitativa exhaustiva ni interpretación automática de los pasajes.

`temporal_role=current/previous` se refiere a la posición en el bloque del informe, no a su vigencia hoy. `period_code` procede del enlace al resumen SBS y no reemplaza la fecha de comité ni la de publicación. Las fechas plurales sin correspondencia explícita mantienen `temporal_role=unspecified` y `needs_review`; no se asignan ordenándolas cronológicamente.

Las perspectivas separadas de Apoyo/JCR/PCR se conservan como perspectivas del informe; no se propagan automáticamente a cada producto. En Moodys las perspectivas de entidad y emisor permanecen separadas, y un guion junto al rating de depósitos no se interpreta como perspectiva estable.

Las reglas solo aceptan valores de la estructura reconocida. Varias observaciones del mismo campo y rol se conservan y marcan para revisión. Una portada sin campos reconocidos produce `unsupported_cover`; un documento sin texto produce `needs_ocr`. No se realiza OCR automático ni se afirma que estos documentos hayan quedado estructurados.

## Descarga, caché y errores

Los PDF se guardan en `data/sources/peru/sbs/documentos_riesgo/documents`, usando los identificadores de clasificadora, período, número y versión. El almacén canónico y su manifiesto siguen separados del inventario.

La descarga utiliza TLS verificado y un plazo de 180 segundos. Se exige una URL oficial HTTPS y se rechazan redirecciones finales fuera del destino. Se comprueban la cabecera PDF, el marcador de cierre, la estructura mediante lectura estricta, el cifrado y los límites de 20 MB/200 páginas. Una respuesta HTML o una transferencia incompleta no se guarda como PDF válido.

El archivo se sustituye de forma atómica solo después de validarlo. La caché exige que la huella del PDF y los datos canónicos coincidan con el manifiesto. Un PDF ausente o corrupto se vuelve a solicitar durante `sync`; `--load-only` lo rechaza. El contrato 2 vuelve a analizar los PDF íntegros del contrato 1 sin descargarlos de nuevo; `--load-only` exige una captura actualizada mediante `sync`. Un fallo de red o estructura conserva la captura válida anterior. La lectura no convierte estos fallos en ausencia de rating.

## Reporte y códigos de salida

`outputs/documentos_riesgo/documentos_riesgo.xlsx` contiene las hojas `campos`, `referencias`, `cobertura` y `comparaciones` con tablas, filtros y encabezados en español. La asociación de entidad/clasificadora se incorpora mediante una unión explícita por `report_id`, sin afirmar que se haya verificado la identidad legal dentro del documento.

- **0:** documentos seleccionados íntegros y todos los campos emitidos extraídos del formato reconocido; no certifica extracción exhaustiva del documento.
- **1:** descarga, inventario o caché incompletos/dañados; no se exporta un reporte nuevo.
- **2:** reporte exportado con campos o documentos pendientes de revisión/extracción.

## Validación realizada

El 9 de octubre de 2026 se comprobó el código del proveedor con ocho PDF reales: cinco formatos de clasificadora, bancos, una financiera, una CMAC y una CRAC, incluyendo un informe histórico 202502. Las cuatro descargas adicionales se hicieron con el transporte del repositorio y TLS verificado. Los documentos e imágenes de prueba permanecen fuera del repositorio.

La extracción produjo 113 registros: 45 campos reconocidos (incluidas tres cifras de concentración) y 68 pendientes de revisión (66 pasajes candidatos y dos fechas de comité). Para Alfin/JCR se reconocieron 9,7% y 13,2% de los depósitos en los 10 y 20 principales depositantes a diciembre de 2025, página 16. Para Banbif/Moodys se reconoció 22,30% para los 20 principales a diciembre de 2025, página 2, sin completar el denominador del pasaje. MicroRate produjo calificación crediticia y perspectiva para CMAC del Santa y CRAC Los Andes.

Se comprobó la migración de los ocho PDF íntegros sin red, una segunda sincronización sin reprocesarlos y siete exportaciones desde caché. Las pruebas cubren porcentajes inválidos, fecha o denominador ausentes, comparativos no promovidos a cifra actual, negaciones en menciones de eventos, límites de candidatos, PDF incompletos, HTML, cifrado, OCR pendiente, versiones, corrupción de caché y conservación tras fallos. Esto no valida todo el histórico ni todas las entidades; la estructuración cualitativa y los eventos confirmados siguen pendientes.

## Comparación de bloques del mismo documento

Desde la versión 0.17.0, `comparaciones` conserva los valores de los bloques Actual/Anterior, con sus páginas y pasajes. Solo compara una pareja única reconocida del mismo campo, informe y huella PDF. Los campos pendientes de revisión, duplicados o sin pareja quedan explícitos; una perspectiva del informe no se propaga a los ratings de depósitos.

Una diferencia produce `rating_text_change` o `outlook_change`, según el campo. No se asigna mejora/deterioro ordenando letras, ni una fecha de evento desde el código SBS. Se excluyen los pasajes cualitativos, las menciones de eventos sin confirmar y las fechas de comité ambiguas. La API independiente es `document_changes(fields)` en `fuentes_financieras.eventos_riesgo`.

En los ocho PDF reales de validación se obtuvieron 27 comparaciones: tres cambios literales de rating, seis pares con texto igual y 18 campos sin bloque anterior reconocido. No se encontró un cambio de perspectiva en esa muestra; los cambios de perspectiva se verificaron mediante pruebas sintéticas, incluyendo pares ambiguos y perspectivas separadas por producto. Estos resultados no representan cobertura exhaustiva del histórico de informes.

### Ampliación de concentración por entidad

Se admite el párrafo histórico JCR cuyo denominador dice «total de depósitos del Banco», incluyendo el espacio interno «d el» observado en la extracción PDF. También se reconoce la oración de Moody’s que compara un porcentaje «al cierre de» un año con otro «al término de» un año anterior. Cada cierre anual se registra en diciembre del año explícito; ambos valores conservan el mismo pasaje y un denominador no especificado. Una cifra repetida en otra página permanece como evidencia independiente: los consumidores deben comparar entidad, documento, período y número de depositantes antes de agregar observaciones.

Estas reglas no asignan la fecha de portada a un porcentaje sin fecha. Las cifras de MicroRate de los 20 principales depositantes sin período en el pasaje continúan como candidatas para revisión. Tampoco se extraen los límites internos ni porcentajes comparativos de una oración vecina. La concentración sigue siendo cobertura parcial de informes, no una serie completa para todas las entidades.

Validación con el código del repositorio y los PDF completos previamente guardados: Alfin/JCR `001196:202502:3:1`, página 16, produjo 14,7% y 17,7% para los 10 y 20 principales a junio de 2025; BanBif/Moody’s `000406:202601:48:1`, página 4, produjo 22,30% a diciembre de 2025 y 20,72% a diciembre de 2024 para los 20 principales. El 22,30% de BanBif ya figuraba por separado en la página 2. Se bloqueó el transporte durante el reprocesamiento para comprobar el uso de PDF locales íntegros y se verificaron las exportaciones `--load-only`. Los archivos de validación permanecieron fuera del repositorio.

### Comparación histórica JCR de CMAC Huancayo

La regla adicional reconoce el formato comprobado de Huancayo que publica los porcentajes de los 10 y 20 principales depositantes y añade una comparación «respectivamente al cierre de diciembre 2024». Se asignan únicamente los dos porcentajes históricos a los grupos publicados, conservando el denominador total de depósitos, el mes explícito, la página y la oración completa. Los porcentajes actuales de esa oración carecen de fecha explícita y permanecen en la evidencia candidata. La regla exige el vínculo «respectivamente», orden creciente de grupos, fechas y porcentajes válidos; todavía no generaliza este formato a otras entidades.

Se descargaron cuatro PDF oficiales mediante `pe.sbs.documentos_riesgo`, sin modificar su inventario ni introducir descargas en el repositorio:

| Entidad / clasificadora | Referencia SBS | Resultado de concentración fechada |
| --- | --- | --- |
| CMAC Huancayo / JCR | `001196:202601:7:1` | Página 14: 10 principales, 2,4%; 20 principales, 3,2%; diciembre de 2024. |
| CMAC Arequipa / JCR | `001196:202601:5:1` | Pasajes candidatos; sin fecha explícita vinculada a las cifras reconocibles. |
| CMAC Cusco / JCR | `001196:202601:6:1` | Pasajes candidatos; sin fecha explícita vinculada a las cifras reconocibles. |
| Interbank / Apoyo | `000408:202601:43:1` | Pasaje candidato; no se asigna el año de otra métrica al porcentaje de concentración. |

Se reprocesaron los cuatro PDF desde caché con el transporte deshabilitado y se verificó la exportación Excel. Solo las dos observaciones históricas de Huancayo resultaron reconocidas por las reglas de concentración. Esto amplía la cobertura comprobada, sin convertir la ausencia de una extracción en ausencia del dato financiero.

### Accionistas y soporte publicados en formatos comprobados

Se estructuran dos bloques completos de texto, conservando página, pasaje y hash PDF:

- JCR: tabla de dos accionistas con encabezado `Accionistas Acciones Participación (%)`, fila `Total` y siguiente encabezado de directorio. Cada accionista conserva su nombre literal en la etiqueta, número de acciones (`shareholder_shares`) y participación publicada (`shareholder_participation`). Se exige total de acciones consistente, porcentajes que sumen 100 y concordancia entre cantidades y porcentajes dentro del redondeo de dos decimales. No se calculan participaciones ausentes ni se infiere voto o control. El bloque no tiene fecha propia: `observation_period` permanece vacío y la vigencia es no especificada.
- Moody’s: viñeta de respaldo del principal accionista reflejado en capitalización de utilidades, con porcentaje, año de capitalización y ejercicio de origen explícitos. `earnings_capitalization` conserva el porcentaje; el período tiene solo el año publicado (`annual_observation`) y el denominador identifica las utilidades del ejercicio publicado. No se añade mes ni importe monetario. La mención no constituye garantía jurídica, compromiso futuro ni calificación independiente de soporte.

La cobertura de `ownership_support` cuenta los campos reconocidos en estos bloques y conserva los pasajes candidatos para revisión. La extracción describe afirmaciones publicadas por la clasificadora, sin verificación registral ni resolución automática de identidades. Las tablas incompletas, repetidas en una página, con sumas inconsistentes o más de dos accionistas permanecen fuera del formato reconocido. Los documentos o páginas distintos no se reconcilian automáticamente.

Validación con PDF reales: CMAC Huancayo/JCR `001196:202601:7:1`, página 24, publicó Municipalidad Provincial de Huancayo con 81.421.830 acciones y 92,33%, y Corporación Interamericana de Inversiones (BID Invest) con 6.763.841 acciones y 7,67%. BanBif/Moody’s `000406:202601:48:1`, página 2, publicó capitalización del 70% en 2025 correspondiente a utilidades del ejercicio 2024. La mención de 99,9% en el apartado de grupo económico de Alfin no se convierte en participación del banco: el sujeto y su relación con otros porcentajes requieren revisión.
