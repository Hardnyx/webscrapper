# Series macroeconómicas y de mercado

La versión 0.30 incorpora catorce proveedores independientes de series y completa el adaptador de curva SBS que estaba pendiente. No depende del código, configuración, credenciales ni archivos de ningún tablero consumidor.

## Proveedores de series

| Dataset | Consulta de muestra | Frecuencia admitida | Acceso |
| --- | --- | --- | --- |
| `pe.bcrp.series` | `PN01206PM` | Diaria, mensual, según sufijo | API BCRPData |
| `br.bcb.series` | `432` | Declarada por consumidor | SGS público |
| `mx.banxico.series` | `SF61745` | Declarada por consumidor | SIE, `BANXICO_TOKEN` |
| `cl.bcch.series` | `F073.TPMTAS.D` | Declarada por consumidor | BDE, `BCCH_USER` y `BCCH_PASSWORD` |
| `eu.bce.series` | `EXR/D.USD.EUR.SP00.A` | Diaria, mensual, trimestral, anual | API SDMX, clave individual |
| `eu.eurostat.series` | `irt_lt_mcby_m?geo=FR` | Mensual, trimestral, anual | JSON-stat, dimensiones distintas del tiempo fijadas |
| `uk.ons.series` | `/economy/inflationandpriceindices/timeseries/d7g7/mm23` | Mensual, trimestral, anual | Generador CSV |
| `uk.boe.series` | `IUDBEDR` | Diaria | IADB CSV |
| `de.bundesbank.series` | `D.I.ZST.ZI.EUR.S1311.B.A604.R10XX.R.A.A._Z._Z.A` | Diaria | Flujo BBSIS, CSV |
| `us.fred.series` | `DGS10` | Declarada por consumidor | CSV público, sin API key |
| `us.tesoro.series` | `10 Yr` | Diaria | Curva par Treasury CSV |
| `global.yahoo.precios` | `^GDAXI` | Diaria | Servicio de gráficos, sin garantía contractual |
| `br.bcb.focus` | `IPCA12` | Diaria | Mediana IPCA 12 meses, `Suavizada='S'`, `baseCalculo=0` |
| `us.fed_filadelfia.spf` | `CPIA` | Trimestral | Columna individual de los XLSX publicados |

Las frecuencias de consulta usan `daily`, `monthly`, `quarterly`, `annual`. Las muestras no constituyen un catálogo de todos los códigos disponibles. Los proveedores rechazan selecciones múltiples o dimensiones libres cuando impedirían identificar una serie.

SPF admite columnas `CPI`, `DRGDP`, `TBILL`, `UNEMP` con sufijos `1` a `6` o `A`, `B`, `C`, siempre que esa columna exista en el archivo. El trimestre es el de la encuesta, no el período pronosticado ni una fecha exacta de publicación. Los horizontes no se fusionan y se conservan en el código de columna. Las series SPF del BCE pueden consultarse como claves individuales del proveedor BCE; no se incorpora la función consumidora que reúne varios horizontes en una tarjeta.

## Uso y caché

```python
from fuentes_financieras import source

series = source('pe.bcrp.series')
result = series.sync(codigo='PN01206PM', frecuencia='monthly',
                     desde='2026-08', hasta='2026-09', keep_raw=True)
data = series.load(codigo='PN01206PM', frecuencia='monthly',
                   desde='2026-08', hasta='2026-09')
```

Cada solicitud corresponde a un código, frecuencia y año completos. Se evita que pedir agosto y luego septiembre cree capturas solapadas de la misma observación. La carga filtra el rango deseado; un extremo `hasta='2026-02'` incluye todo febrero. El código y frecuencia se identifican en el período mediante un hash estable, con particiones por serie y año. Los años se consideran revisables: se comprueba antigüedad y hash antes de reutilizar la caché. `force=True` fuerza actualización. El máximo es 101 años por plan; la ejecución conjunta añade su propio límite total.

Una captura validada acredita la integridad de las observaciones recibidas, no un calendario completo de datos ni todos los períodos de la historia. No se rellenan fines de semana, festivos ni huecos; las filas ausentes en origen siguen ausentes y los marcadores explícitos quedan nulos con `source_value_missing`. No se descartan silenciosamente números malformados ni fechas duplicadas. Los errores de descarga o estructura conservan la captura canónica anterior y figuran como fallo.

Las respuestas se normalizan antes de persistir `raw`: en estos adaptadores ese archivo contiene las observaciones canónicas, no una copia byte a byte de la respuesta. Esto evita persistir consultas con credenciales. El archivo raw de la curva SBS sí conserva el cuerpo JSON original.

## Contrato y escala

`series_code`, `title`, `period`, `period_date`, `frequency`, `value`, `source_value`, `unit`, `source_unit_multiplier`, `observation_status`, `data_quality_flags`, `source`, `source_url`, `retrieved_at`.

`period` conserva la etiqueta original cuando la fuente la publica. Los meses y trimestres se anclan al primer día del período para permitir orden y filtros; esa fecha no se presenta como fecha real de publicación. No se convierten frecuencias ni se calcula un promedio mensual de datos diarios.

`frequency` es declarada por el consumidor. Se comprueba cuando el código o la estructura publicada permiten hacerlo, pero los proveedores no consultan automáticamente todos los catálogos de metadatos. `unit='source_native'` significa escala no identificada por este adaptador: no autoriza a interpretar el valor como porcentaje, decimal, soles o dólares. Los códigos de unidad y el exponente `UNIT_MULT` se conservan cuando vienen en la respuesta; no se aplican multiplicadores ni se reescalan por el tamaño del número. El consumidor debe validar la definición de su serie antes de combinarla con otras.

Treasury identifica `percent_par_yield`; Focus `percent_expected_inflation_12m`. Yahoo conserva cierres sin redondearlos ni reemplazarlos por precios ajustados: utiliza el huso de la bolsa y excluye la sesión en curso cuando no puede acreditar su cierre. El título y código de FRED identifican al agregador; no se infiere el organismo emisor original desde el identificador.

## Comando

```bash
python scripts/sync_series_macro.py --dataset pe.bcrp.series --codigo PN01206PM --frecuencia monthly --desde 2026-08 --hasta 2026-09 --data-root /ruta/datos --output-dir /ruta/reportes
```

También está disponible `fuentes-sync-series` al instalar el paquete. `--load-only` verifica la selección sin sincronizar; `--force` fuerza la captura y es incompatible con `--load-only`. El Excel incluye `serie` y `capturas`, tablas con filtros, estilo claro 9 y encabezado fijo; no cambia anchos de columnas. Una selección incompleta no exporta un reporte nuevo.

Los proveedores pueden usarse desde perfiles TOML de `sync_fuentes.py` y su registro de novedades:

```toml
version = 1
[[trabajos]]
nombre = "Tipo de cambio BCRP"
dataset = "pe.bcrp.series"
[trabajos.consulta]
codigo = "PN01206PM"
frecuencia = "monthly"
desde = "2026-08"
hasta = "2026-09"
```

No se leen archivos de credenciales de otros proyectos. Las credenciales de Banxico y Chile solo se obtienen del entorno local y nunca se escriben en el catálogo, perfiles, excepciones de transporte o URLs guardadas. No se han validado conexiones reales de esos dos servicios por falta de credenciales en este entorno. No deben enviarse credenciales a un repositorio público.

## Curvas SBS

`pe.sbs.curva_soberana` admite `CCPSS`, `CBCRPS`, `CCPVS`, `CCPEDS`, `CBCRS`, `CCSDF`, `CSBCRD`, `CCINFS`, `CCCLD`. Los códigos se tomaron del extractor revisado; esta integración no certifica disponibilidad de todos ellos en todas las fechas.

```python
curve = source('pe.sbs.curva_soberana')
curve.sync(desde='2026-10-22', hasta='2026-10-22', curvas=['CCPSS'])
points = curve.load(desde='2026-10-22', hasta='2026-10-22', curvas=['CCPSS'])
```

También admite `fetch(fecha='2026-10-22', curva='CCPSS')`. Las solicitudes son diarias y omiten fines de semana, sin un calendario de feriados. Se valida éxito explícito, fecha devuelta, tipo de curva cuando está presente, plazos enteros no negativos y ausencia de duplicados. Se conserva toda la precisión de `tasa`, sin interpolación, selección de plazos, inflación implícita ni spreads. Contrato: curva, fecha, plazo en días, tasa anual como la publica la fuente y trazabilidad. Contrato de almacenamiento 2; se sustituye el adaptador pendiente anterior.

El entorno recibió tanto una respuesta JSON real de 161 puntos para CCPSS del 22/10/2026 como rechazos de seguridad HTTP 200 de la SBS. La captura normal logró validar y guardar esos 161 puntos; una actualización forzada posterior fue rechazada. El acceso es intermitente y la captura válida se conserva ante ese fallo. Las peticiones de API no añaden la cabecera AJAX por defecto; las pantallas WebForms existentes conservan su configuración anterior. El detector permite comprobar el acceso desde otra red sin tratar un rechazo como éxito.

Para comprobar la página desde el PC:

```bash
python tools/diagnostico_curva_sbs.py --fecha 2026-10-22 --curva CCPSS
```

El detector es autónomo: verifica versiones e instala solo dependencias ausentes o incompatibles con `sys.executable`. Mantiene TLS verificado, guarda respuesta y diagnóstico fuera del repositorio y distingue bloqueo HTTP/WAF, error HTTP, conexión/TLS y estructura. Una respuesta HTTP 200 que diga «Solicitud Rechazada» es bloqueo, no datos válidos. Recibir puntos no equivale a certificar la metodología financiera.

## Validación real de esta integración

Se usó septiembre de 2026 para doce proveedores públicos de series: BCRP, SGS, BCE, Eurostat, ONS, BoE, Bundesbank, FRED, Treasury, Yahoo, Focus y SPF (esta última con trimestres 1–3). No es una comprobación de todo su histórico. Los resultados, capturas, reportes y scripts temporales se guardaron fuera de git.

Se adaptó Bundesbank a su CSV actual identificado por clave, con separador coma, conservando además el formato antiguo con punto y coma. En Focus se comprobó que el problema del filtro era la codificación de espacios: Olinda requiere `%20` en esta consulta y responde con error si se usa la codificación de formulario con `+`; el valor de `Suavizada` continúa siendo la cadena `S`. No se sustituyó por un booleano.

## Auditoría del conjunto de extractores revisado

| Módulo o bloque de origen | Aplicación en esta biblioteca | Estado de esta entrega |
| --- | --- | --- |
| `bcrp`, `bcb`, `banxico`, `bcch` | Series originales de bancos centrales | Integrados en proveedores nuevos |
| `fred`, `tesoro`, `yahoo` | Observaciones originales de mercado | Integrados; FRED utiliza CSV público |
| `ecb`, `eurostat`, `ons`, `bundesbank`, `boe` | Series estadísticas individuales | Integrados |
| `focus` | IPCA esperado a 12 meses y expectativas anuales | IPCA12 integrado; recurso anual pendiente |
| `spf` | Medianas por trimestre y columna de horizonte | Integrado; no se copia la tarjeta de última encuesta |
| `sbs_curva` | Puntos originales de curvas SBS | Adaptador integrado; bloqueo de red documentado |
| `sbs`, `sbs_tasas_parser`, `sbs_store` | Tasas originales y almacenamiento | Ya existen proveedores SBS; no se duplica su almacén columnar ni la envolvente de entidades |
| `sbs_indicadores` | Ratios publicados por entidad | Ya existen solvencia, liquidez, calidad y rentabilidad; equivalencia de todas las variantes históricas todavía no comprobada |
| `sbs_clasificaciones`, `sbs_clasificaciones_web` | Resumen original y consulta histórica | Ya existen; no se trasladan elegibilidad ni límites por rating |
| `smv_fondos` | Valores cuota y catálogo de fondos | Ya existe proveedor SMV; no se copian restricciones de series ni reglas de rentabilidad del tablero |
| `bcb_titulos` | Descarga mensual de negociaciones originales | La captura bruta es aplicable y queda pendiente; valoración de títulos, calendario de flujos, filtros de atípicos y ajuste Nelson–Siegel permanecen en el consumidor |
| `bcrp_encuesta` | Encuesta publicada en XLSX con fecha, grupo y año objetivo | Aplicable; adaptación pendiente para conservar dimensiones y verificar escala del archivo |
| `bcrp_reporte` | Cuadro de proyecciones publicado en PDF | Aplicable; adaptación pendiente como fuente documental con versiones y evidencia |
| `diferencial`, `inflacion_implicita`, `prima_depositos` | Cálculos sobre otras fuentes y selección de contrapartes | No son extractores; permanecen en el consumidor |
| `web`, `navegador`, inicialización de TLS | Transporte y recuperación de páginas | Se usa el transporte verificado existente; no se copia el navegador ni se desactiva TLS |

Esta es la primera integración del conjunto: los pendientes de la tabla no se presentan como completados. No se modifica el proyecto consumidor, no se instala su calendario de actualización y no se trasladan interfaz, configuración empresarial o datos descargados al repositorio.

Se comprobó además la exportación del comando de series y una ejecución conjunta offline de BCRP, SGS, BCE, Focus y SPF: cinco capturas verificadas y registradas como novedades, sin consultas de red en esa ejecución.
