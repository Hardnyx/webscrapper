# Anuncios regulatorios con evidencia

Dataset independiente: `pe.sbs.anuncios_regulatorios`, contrato 1, incorporado en la versión 0.18.0; descubrimiento explícito desde la versión 0.19.0.

El proveedor recibe enlaces explícitos de noticias oficiales de la SBS y captura su contenido con TLS verificado. El proveedor de anuncios no consulta automáticamente otros proveedores ni búsquedas externas. La CLI puede consumir expresamente el índice independiente mediante `--descubrir`. Se admite únicamente HTTPS en `www.sbs.gob.pe/noticia/detallenoticia/idnoticia/ID`, o la ruta oficial `DetalleNoticia?IdNoticia=ID`, con un parámetro de título opcional que se elimina al construir el enlace canónico. Los identificadores duplicados o contradictorios se rechazan. Varias URLs del mismo identificador se procesan una sola vez.

## Ejecución

```bash
python scripts/sync_anuncios_regulatorios.py --urls https://www.sbs.gob.pe/noticia/detallenoticia/idnoticia/3749

# Exportar las noticias seleccionadas ya validadas, sin red.
python scripts/sync_anuncios_regulatorios.py --urls https://www.sbs.gob.pe/noticia/detallenoticia/idnoticia/3749 --load-only
```

`--force` vuelve a consultar las páginas seleccionadas. `--data-root` y `--output-dir` cambian las carpetas de datos y reportes. El lanzador comprueba versiones e instala únicamente dependencias ausentes o incompatibles. También se expone `fuentes-sync-regulatory-announcements` para instalaciones del paquete.

```python
from fuentes_financieras import source
announcements = source('pe.sbs.anuncios_regulatorios')
announcements.sync(urls=urls, keep_raw=True)
data = announcements.load(urls=urls)
```

## Descubrir enlaces del índice

El dataset independiente `pe.sbs.indice_noticias` captura páginas seleccionadas del listado SBS. Conserva título completo, fecha publicada en la tarjeta, enlace canónico, número de página, última página indicada por el sitio y fecha de captura. No descarga artículos ni confirma eventos a partir del título.

```bash
# Revisar dos páginas del índice sin descargar sus artículos.
python scripts/sync_anuncios_regulatorios.py --descubrir --paginas 1 2 --solo-indice

# Consumir el índice y descargar las noticias seleccionadas por título y fecha.
python scripts/sync_anuncios_regulatorios.py --descubrir --paginas 18 --desde 2024-07-11 --hasta 2024-07-11 --palabras Sullana

# Repetir una revisión del índice desde capturas locales íntegras.
python scripts/sync_anuncios_regulatorios.py --descubrir --paginas 1 2 --solo-indice --load-only
```

`--descubrir` y `--urls` son excluyentes. Sin `--paginas`, se revisa únicamente la página 1. Los filtros de fechas son inclusivos y se aplican a la fecha del índice, que permanece separada de la fecha escrita en el anuncio. Las palabras son subcadenas literales del título completo: basta una coincidencia, sin distinguir mayúsculas, conservando tildes. Una coincidencia selecciona un artículo para lectura; no confirma un evento.

Las hojas `indice`, `seleccion` y `cobertura_indice` muestran todas las tarjetas capturadas, la selección única por identificador y el alcance de cada página. Si no hay coincidencias, se exporta la selección vacía con su cobertura, sin consultar artículos. `--solo-indice` tampoco accede a los artículos. Con `--load-only`, ambos datasets deben estar íntegros en caché para las referencias seleccionadas si se solicitan detalles.

Cada respuesta debe identificar la página solicitada, su módulo y la última página publicada; una respuesta que repita la página 1 para otra página se rechaza. Se validan tarjetas completas, fechas y enlaces oficiales; referencias contradictorias entre páginas requieren una captura nueva. No se detiene el recorrido por una fecha antigua, porque no se presupone un orden cronológico perfecto. Las páginas se consideran mutables y se vuelven a comprobar al vencer el intervalo de refresco o con `--force`.

La captura **solo cubre las páginas solicitadas**, incluso si el sitio muestra un número mayor de páginas. Las páginas pueden cambiar entre consultas; los tiempos de captura permanecen visibles. No se reconstruye una instantánea histórica del índice ni se deduplican actos legales por semejanza de títulos: diferentes identificadores pueden corresponder a publicaciones sobre el mismo hecho.

```python
index = source('pe.sbs.indice_noticias')
index.sync(paginas=[1, 2], keep_raw=True)
links = index.load(paginas=[1, 2])
announcements.sync(urls=links.article_url.drop_duplicates().tolist())
```

## Qué se reconoce

Se exige un título y un cuerpo únicos de la estructura comprobada de DetalleNoticia. Solo se analiza el primer párrafo no vacío del cuerpo. El párrafo debe comenzar con una fecha completa válida de Lima y contener una disposición directa, dentro de uno de los formatos comprobados:

- Intervención anunciada: motivo explícito, sujeto institucional SBS, acción realizada y nombre legal de la entidad.
- Disolución e inicio de liquidación anunciados: nombre legal de la entidad en intervención, resolución citada y disposición explícita de la SBS.

El título también debe corresponder al tipo de disposición. No se extraen eventos desde la navegación, el título aislado, párrafos históricos posteriores, negaciones o hipótesis. Si una noticia válida no cumple estas reglas, se conserva con `needs_review`, sin entidad ni evento confirmado. Este estado no significa que la noticia carezca de eventos; significa que el formato no fue reconocido.

Cada registro conserva identificador, título, URL canónica, párrafo fuente, ubicación de evidencia, huella SHA-256 del HTML decodificado y vuelto a codificar en UTF-8, y fecha de extracción. `announcement_date` es la fecha escrita en el anuncio. **No equivale necesariamente a la fecha de publicación de una resolución ni a su vigencia legal**. `effective_date` queda vacío; el número de resolución solo se completa cuando aparece en la estructura reconocida.

Los nombres se conservan tal como aparecen en el anuncio. No se unen automáticamente al universo actual, ni se afirma identidad legal por similitud. Una noticia puede describir una situación histórica que haya cambiado después. La selección no certifica el estado actual de una entidad.

## Caché y reporte

El almacén independiente está en `data/sources/peru/sbs/anuncios_regulatorios`, con identificadores de noticia como claves. Las páginas se consideran mutables y pueden volver a comprobarse después del intervalo predeterminado de 24 horas. La CLI conserva también el HTML comprimido de la captura válida más reciente. No mantiene un archivo de todas las revisiones históricas de la página.

Se rechazan HTML incompletos, contenido mayor de 2 MB, redirecciones a otra noticia y estructura ausente o ambigua. La CLI exige capturas canónicas íntegras y compatibles. Una descarga o análisis fallidos conservan la captura válida anterior y no generan un reporte nuevo a partir de una selección incompleta.

`anuncios_regulatorios.xlsx` contiene `anuncios`, con todas las noticias seleccionadas y su estado, y `eventos`, con disposiciones reconocidas. Ambas hojas tienen encabezados en español, filtros y primera fila fija; las hojas con datos usan tablas de estilo claro 9, sin ajustar anchos.

- **0:** todas las noticias seleccionadas tienen una disposición reconocida; no implica cobertura exhaustiva del contenido ni del universo regulatorio.
- **1:** selección ausente, incompatible, dañada o captura fallida; no se exporta un reporte nuevo.
- **2:** reporte exportado con noticias que requieren revisión.

## Validación y pendientes

Se descargaron cinco noticias reales desde este entorno mediante el proveedor del repositorio: intervenciones de CMAC Sullana, Financiera Credinka y CRAC Raíz; disolución e inicio de liquidación de CRAC Raíz; y una noticia general de regulación cooperativa. Se reconocieron cuatro disposiciones; la noticia general permaneció sin evento confirmado. Se comprobaron reutilización inmediata de caché y exportación sin red. HTML, cachés y reportes de validación están fuera del repositorio.

Se validaron además tres páginas reales del índice (1, 2 y 18), con 45 noticias y paginación comprobada, reutilización de caché y exportación sin red. La CLI se verificó seleccionando anuncios desde el índice y descargando sus detalles, sin proporcionar URLs manuales. Las pruebas también cubren páginas equivocadas, tarjetas incompletas, fechas inválidas, enlaces ambiguos, filtros literales y selecciones vacías.

El alcance sigue siendo parcial: no hay recorrido exhaustivo predeterminado de todo el índice, archivo exhaustivo de resoluciones, retiros de rating confirmados por clasificadoras ni envío de alertas. Las [resoluciones de fusión](elperuano-fusiones.md) se extraen mediante un proveedor independiente desde El Peruano, con URLs explícitas y formatos comprobados. Las noticias y transferencias de bloques patrimoniales no se convierten automáticamente en fusiones. Tampoco se acredita la ejecución registral a partir de una autorización.
