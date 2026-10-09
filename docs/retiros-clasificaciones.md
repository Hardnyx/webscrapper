# Retiros explícitos de clasificaciones

Dataset independiente `pe.clasificadoras.retiros`, contrato 1, versión 0.21.0.
Recibe enlaces explícitos de comunicados PDF de Moody’s Local Perú y Apoyo & Asociados. La captura usa TLS verificado y conserva el PDF con SHA-256. No consulta automáticamente el inventario SBS, otros proveedores, los sitios completos ni las páginas HTML de noticias.

## Uso

```bash
python scripts/sync_retiros_clasificaciones.py --urls \
  https://moodyslocal.com.pe/wp-content/uploads/2026/07/ML-PE-PR-Financiera-Qapaq-S.A.pdf \
  https://moodyslocal.com.pe/wp-content/uploads/2026/03/MLPE_Comunicadodeprensa-Santander-Consumer-Bank-27032026.pdf
```

```python
from fuentes_financieras import source

provider = source("pe.clasificadoras.retiros")
urls = ["https://moodyslocal.com.pe/wp-content/uploads/2026/07/ML-PE-PR-Financiera-Qapaq-S.A.pdf"]
provider.sync(urls=urls)
data = provider.load(urls=urls)
```

La CLI instalada es `fuentes-sync-rating-withdrawals`. El lanzador verifica las dependencias base y el extra `pdf`, e instala únicamente las faltantes o incompatibles mediante el mismo intérprete. `--data-root` y `--output-dir` permiten guardar las comprobaciones fuera del repositorio.

Se admiten únicamente PDF HTTPS bajo `/wp-content/uploads/AAAA/MM/archivo.pdf` en `moodyslocal.com.pe` o `www.aai.com.pe`, sin parámetros ni fragmentos. La URL canónica define el identificador del documento mediante SHA-256, evitando colisiones entre archivos del mismo nombre. Una redirección a otro documento se rechaza. Los enlaces duplicados se descargan una sola vez.

## Alcance reconocido

| Formato comprobado | Extracción |
| --- | --- |
| Moody’s afirma y retira Entidad y Emisor | Entidad literal, fecha del bloque de acción y ambas clasificaciones retiradas; motivo vacío si no aparece en ese bloque |
| Moody’s afirma otras clasificaciones y retira un programa de certificados | Solo el programa expresamente retirado y su vencimiento; las clasificaciones afirmadas no se incluyen en el alcance retirado |
| Apoyo & Asociados anuncia retiro de institución e instrumentos por contrato | Nombre literal del comunicado, fecha, alcance que enumera institución, depósitos, certificados y bonos subordinados, con sus ratings y motivo explícito |

Cada documento genera una fila. `withdrawal_scope` conserva el alcance reconocido, sin convertirlo en un retiro de todos los ratings del emisor. Los valores se mantienen con sus escalas originales; no se ordenan ni homologan las escalas de distintas clasificadoras. El nombre puede ser un alias publicado, como `Caja Arequipa`, y no se sustituye automáticamente por una razón social o un identificador interno.

El evento `rating_withdrawal_announced` confirma únicamente lo indicado por la clasificadora en ese comunicado. Un retiro por vencimiento o resolución de contrato no se convierte en downgrade, intervención, insolvencia, pérdida de autorización ni exclusión del universo de depósitos. La fecha corresponde al bloque fechado del comunicado, sin inferir otra fecha efectiva ni el estado vigente de la entidad.

## Límites del parser

Para Moody’s se exige un único bloque `ACCIÓN DE CLASIFICACIÓN`, fechado en Lima, con la identidad de la clasificadora y cierre antes de la tabla de acciones. Para Apoyo & Asociados se exige su encabezado fechado y el párrafo anterior a `Contactos`. El texto completo del bloque debe ajustarse a un formato comprobado. Las tablas `RET` o `RETIRADA`, el título, los fundamentos históricos y las páginas posteriores no bastan por sí solos para crear un evento.

Negaciones, hipótesis, ratificaciones, formatos no reconocidos y bloques ambiguos conservan `needs_review`, sin evento, entidad ni alcance extraídos. La fecha y el párrafo candidato pueden conservarse para revisión. Una portada sin texto conserva `needs_ocr`. Estos estados no significan que no exista un retiro; describen el límite de extracción.

Se valida firma PDF, cierre EOF, tamaño máximo de 20 MB, lectura estricta, ausencia de cifrado y entre 1 y 200 páginas. No se aplica OCR ni se ejecuta contenido del documento.

## Caché y reportes

El almacén es `data/sources/peru/clasificadoras/retiros`. Los documentos se reutilizan cuando el PDF local, el contrato, el manifiesto y los datos canónicos son válidos. `--force` reprocesa el PDF local verificado; `--redownload` vuelve a obtenerlo desde la fuente. Si falta el PDF o cambia su hash local, una sincronización normal vuelve a descargarlo. Una descarga inválida no sustituye el PDF válido ni los datos canónicos anteriores. No se mantiene un archivo de todas las revisiones de una URL.

`--load-only` exporta sin red y exige integridad tanto del Parquet como del PDF. No se combina con `--force` ni `--redownload`.

`retiros_clasificaciones.xlsx` tiene hojas `comunicados` y `retiros`, encabezados en español, tablas `TableStyleLight9`, filtros, congelación `A2` y sin ajustes de anchos. La segunda hoja incluye solo retiros reconocidos y puede estar vacía. Códigos: `0` toda la selección reconocida, `2` documentos válidos con casos por revisar, `1` captura incompleta o caché inválida, sin reporte nuevo.

## Validación y pendientes

Formatos contrastados con los PDF oficiales:

- [Financiera Qapaq, julio de 2026](https://moodyslocal.com.pe/wp-content/uploads/2026/07/ML-PE-PR-Financiera-Qapaq-S.A.pdf): retiro de Entidad y Emisor; no se atribuye un motivo desde otras páginas.
- [Santander Consumer Bank, marzo de 2026](https://moodyslocal.com.pe/wp-content/uploads/2026/03/MLPE_Comunicadodeprensa-Santander-Consumer-Bank-27032026.pdf): retiro de un programa por vencimiento; se mantienen fuera del alcance retirado los ratings afirmados.
- [Caja Arequipa, septiembre de 2025](https://www.aai.com.pe/wp-content/uploads/2025/09/Retiro-Clasificaciones-Caja-Arequipa-privado.pdf): retiro explícito de institución e instrumentos por resolución de contrato.
- [Financiera Qapaq, marzo de 2026](https://moodyslocal.com.pe/wp-content/uploads/2026/03/MLPE_Comunicadodeprensa-FinancieraQapaq-24032026-1.pdf): control de ratificación, sin evento de retiro reconocido.

Las pruebas cubren alcance parcial, negaciones, hipótesis, texto histórico, ambigüedad, fechas inválidas, URL incorrecta, PDF corrupto, página sin texto, caché, reproceso sin descarga, redescarga, recuperación de PDF dañado y exportación sin red. Los PDF y reportes de validación permanecen fuera del repositorio.
La sincronización real mediante el lanzador del repositorio descargó cuatro PDF: tres retiros reconocidos y la ratificación sin evento de retiro. Se comprobaron la reutilización de las cuatro capturas, el reproceso forzado sin acceso a red y el reporte sin red con cuatro comunicados y tres retiros. Se verificaron las tablas, filtros, congelación y ausencia de cambios de ancho.

Quedan pendientes otros formatos de estas agencias, comunicados PCR, JCR y MicroRate, descubrimiento paginado de comunicados, otros instrumentos, extracción de motivos de páginas posteriores y envío de alertas. El comunicado corporativo de Apoyo & Asociados sobre Telefónica de diciembre de 2025 utiliza otro formato y no se marca como retiro reconocido por este parser. La cobertura no es exhaustiva.
