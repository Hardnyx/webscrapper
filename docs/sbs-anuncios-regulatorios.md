# Anuncios regulatorios con evidencia

Dataset independiente: `pe.sbs.anuncios_regulatorios`, contrato 1, incorporado en la versión 0.18.0.

El proveedor recibe enlaces explícitos de noticias oficiales de la SBS y captura su contenido con TLS verificado. No consulta automáticamente la sala de prensa, otros proveedores o búsquedas externas. Se admite únicamente HTTPS en `www.sbs.gob.pe/noticia/detallenoticia/idnoticia/ID`, con un parámetro de título opcional que se elimina al construir el enlace canónico. Varias URLs del mismo identificador se procesan una sola vez.

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

El alcance sigue siendo parcial: no hay descubrimiento automático de noticias, archivo exhaustivo de resoluciones, extracción de fusiones o autorizaciones, retiros de rating confirmados por clasificadoras ni envío de alertas. Las transferencias de bloques patrimoniales no se convierten automáticamente en fusiones. Esos formatos requieren fuentes y validación propias.
