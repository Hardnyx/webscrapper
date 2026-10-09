# Resoluciones SBS de fusión por absorción

Dataset independiente `pe.elperuano.fusiones`, contrato 1, versión 0.20.0.
Recibe enlaces explícitos de dispositivos legales de El Peruano y solicita su visor HTML con TLS verificado. No consulta el índice SBS, el universo de entidades ni otro proveedor automáticamente.

## Uso

```bash
python scripts/sync_fusiones.py --urls \
  https://busquedas.elperuano.pe/dispositivo/NL/2009000-1 \
  https://busquedas.elperuano.pe/dispositivo/NL/2115232-1
```

```python
from fuentes_financieras import source

provider = source("pe.elperuano.fusiones")
urls = ["https://busquedas.elperuano.pe/dispositivo/NL/2009000-1"]
provider.sync(urls=urls, keep_raw=True)
data = provider.load(urls=urls)
```

También se admiten enlaces `https://busquedas.elperuano.pe/api/visor_html/ID`.
Ambas rutas comparten la misma clave de dispositivo; se descargan una sola vez por selección. No se admiten otros dominios, HTTP, parámetros, fragmentos ni rutas de cuadernillo completo.
La CLI instalada es `fuentes-sync-mergers`. `--data-root` y `--output-dir` permiten ubicar las capturas y el reporte fuera del repositorio. `--load-only` exporta sin red; `--force` recaptura. No se combinan ambas opciones.

## Qué se reconoce

| Campo o disposición | Regla |
| --- | --- |
| Fecha de resolución | Datación completa del encabezado, separada de la vigencia y de la fecha de publicación, que no se extrae |
| Fusión autorizada | Primer párrafo después de `RESUELVE:`, que autoriza expresamente la fusión por absorción entre dos entidades |
| Aclaración de vigencia | Artículo único que autoriza la aclaratoria de la minuta y estipula una fecha explícita; constituye una disposición distinta de la autorización original |
| Entidades | Nombres literales en el artículo y en la sumilla; sin resolver aliases ni asignar identificadores internos |
| Roles | Solo cuando el artículo declara que la segunda entidad se extingue sin liquidarse; el orden de los nombres no basta en los otros formatos |
| Condiciones | Texto restante del artículo de autorización, conservado para revisión; no ejecuta ni evalúa condiciones jurídicas |

Las disposiciones reconocidas usan `merger_authorized` o `merger_date_clarification`. La fecha estipulada en una aclaración **no acredita inscripción, ejecución, elegibilidad ni vigencia actual**. Una autorización sin fecha reconocida conserva `effective_date` vacío. No se usa la fecha del encabezado como fecha efectiva ni se unen automáticamente resoluciones anteriores y posteriores.

El parser valida el contenedor identificado `xID`, una sola historia, sumilla, encabezado de resolución SBS, fecha válida, bloque resolutivo y cierre con el mismo identificador. El `<title>` del visor puede pertenecer a otra norma del cuadernillo y se ignora. Solo el primer párrafo resolutivo es elegible. Solicitudes, considerandos, negaciones, reglamentos generales, citas y artículos posteriores no crean eventos.

## Evidencia, caché y reporte

Cada fila conserva número de resolución, identificador de dispositivo, sumilla, párrafo de evidencia, ubicación, URL canónica, SHA-256 del HTML decodificado en UTF-8 y fecha de extracción. El almacén es `data/sources/peru/elperuano/fusiones`; las páginas son mutables y pueden revisarse tras el intervalo predeterminado de 24 horas.

La CLI guarda el HTML comprimido de la captura válida más reciente. Una captura fallida o un documento sin estructura válida no sustituye datos canónicos previos. No se conserva un archivo de todas las revisiones. La exportación exige que toda la selección esté validada y que el contenido Parquet coincida con su manifiesto y contrato.

`fusiones.xlsx` contiene `resoluciones` y `disposiciones`, con encabezados en español, tablas `TableStyleLight9`, filtros, congelación `A2` y sin ajustes de anchos. La segunda hoja contiene únicamente disposiciones reconocidas; puede estar vacía. Salidas: `0` selección reconocida, `2` captura válida con disposiciones por revisar, `1` fallo o caché incompatible; no se genera un reporte nuevo cuando falta una captura válida.

## Validación y cobertura pendiente

Formatos contrastados con el HTML oficial:

- [03245-2021, Servicios Financieros TOTAL EDPYME / Factoring Total](https://busquedas.elperuano.pe/dispositivo/NL/2009000-1): autorización, sin fecha efectiva extraída ni roles inferidos del orden.
- [01724-2022, Mapfre Perú Vida / Mapfre Perú](https://busquedas.elperuano.pe/dispositivo/NL/2071181-1): autorización y roles explícitos por extinción de la segunda entidad; sin fecha efectiva extraída.
- [03119-2022, aclaratoria de TOTAL / Factoring Total](https://busquedas.elperuano.pe/dispositivo/NL/2115232-1): resolución del 12 de octubre de 2022 que estipula el 1 de enero de 2022 como fecha de vigencia, sin convertirla en una nueva autorización ni certificar ejecución.

Las pruebas cubren negaciones, citas históricas, documentos incompletos, identificadores contradictorios, fechas inválidas, roles desconocidos, redirecciones, conservación de datos tras fallos y exportación sin red.
La sincronización real descargó las tres resoluciones, reutilizó las tres capturas y exportó sin red. Como controles, la resolución de cambio de nombre de Banco Azteca a Alfin (`2006679-1`) conservó el estado de revisión sin crear una fusión, y la ley general de control de concentraciones (`1917847-1`) fue rechazada por no ser una resolución SBS. Los HTML, cachés y reportes de comprobación permanecen fuera del repositorio.

La cobertura es parcial: no descubre todo el archivo legal, no procesa PDF, OCR, fe de erratas, múltiples historias, fusiones por constitución, escisiones ni transferencias patrimoniales. Tampoco extrae aún todas las modificaciones de nombre, autorizaciones de funcionamiento o fechas condicionadas de otros formatos. El dispositivo antiguo `1530118-1` ofrece solo visor PDF y su endpoint HTML devolvió HTTP 502 en la comprobación; no se contabilizó como captura válida. La falta de HTML no se interpreta como ausencia de la resolución.
