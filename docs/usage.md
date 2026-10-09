# Comandos y aplicaciones

Ejecutar desde la raíz del repositorio, salvo que se indique otra ruta. Los
lanzadores resuelven sus archivos respecto del repositorio y también pueden
invocarse mediante una ruta absoluta desde otro directorio.

## Tasas pasivas

```powershell
python scripts/sync_tasas_pasivas.py --tipos B C F R --desde 2026-09-01 --hasta 2026-09-30 --excel outputs/tasas_pasivas.xlsx
```

B/F se consultan por día hábil; C/R por mes. `--load-only` lee únicamente el
almacén local. `--force` solicita revalidación explícita. Una sincronización
incompleta termina con código 1 y no exporta un Excel que parezca completo.
La exportación incluye tabla, filtros, estilo claro 9 y encabezado inmovilizado.

El aviso explícito SBS «No existe información para la fecha elegida» se
registra como período no disponible, sin confundirlo con un cambio de
estructura. Las respuestas incompletas sin ese aviso siguen siendo errores.
Los períodos recientes no disponibles se reconsultan cuando vence el intervalo
de refresco; `--force` permite revalidarlos antes. No se sustituye la fecha
solicitada por otro período con datos.

## Valores cuota SMV

```powershell
python scripts/sync_valores_cuota.py --help
```

El comando admite rangos de capturas, migración desde datos anteriores y
backfill histórico. Se reutilizan particiones válidas; el reporte consumidor
filtra los instrumentos después de la extracción.

## Clasificaciones históricas

```powershell
python scripts/sync_clasificaciones_riesgo.py --data-root data/sources --output-dir outputs/clasificaciones_riesgo
```

[Contrato y validación](sbs-clasificaciones-riesgo.md).

## Anuncios regulatorios

```powershell
python scripts/sync_anuncios_regulatorios.py --urls https://www.sbs.gob.pe/noticia/detallenoticia/idnoticia/3749
```

Procesa enlaces explícitos de noticias SBS; `--load-only` exporta una selección íntegra desde caché. Las noticias sin disposición reconocida permanecen visibles para revisión. [Alcance y validación](sbs-anuncios-regulatorios.md).

## Tipo de cambio contable

```python
from fuentes_financieras.providers.sbs.tipo_cambio_contable import get_accounting_exchange_rate

rate = get_accounting_exchange_rate(
    "2026-09-30", storage_dir="data/sources/peru/sbs/tipo_cambio_contable"
)
print(rate["date"], rate["usd_pen_accounting"])
```

La fecha efectiva puede ser anterior a la solicitada. Los imports anteriores
mediante `fuentes_financieras.sbs_tipo_cambio` siguen funcionando.

## Aplicaciones de escritorio

```powershell
python scripts/run_desktop.py --app average-exchange-rate
python scripts/run_desktop.py --app weighted-exchange-rate
```

Requieren un entorno gráfico, Tkinter y un navegador compatible con Selenium.
El lanzador instala las dependencias Python del extra `desktop` cuando faltan.

## Empaquetar

```powershell
python scripts/build_app.py --app passive-rates
python scripts/build_app.py --app average-exchange-rate
python scripts/build_app.py --app weighted-exchange-rate
```

Los binarios se guardan en `dist/`. Tasas pasivas usa consola; los dos tipos de
cambio conservan interfaz gráfica. El empaquetado debe validarse en Windows.

## Capturar un sitio

Editar la URL y opciones de `tools/site_capture/site_dump_config.json`:

```powershell
python scripts/capture_site.py
python scripts/capture_site.py --config C:/ruta/config.json
```

Sin `--config`, se utiliza la configuración incluida con la herramienta. Una
ruta de salida relativa se interpreta respecto del archivo de configuración.
El ejemplo apunta a `outputs/site_capture/`. El comando admite también el
lanzador Windows `tools/site_capture/run_site_dump.bat`.

## Utilidades de git

- `bash tools/git/dump_filetree.sh`: escribe el inventario de archivos versionados en `outputs/filetree.txt`.
- `bash tools/git/commit_push.sh "feat(sbs): add a data source"`: revisa el estado, agrega los cambios, crea un commit convencional y ejecuta push.

## Liquidez SBS

```powershell
python scripts/sync_liquidez.py --desde 2026-08
python scripts/sync_liquidez.py --datasets cobertura --desde 2026-06
python scripts/sync_liquidez.py --datasets financiacion --desde 2026-07
```

[Períodos, unidades y cobertura comprobada](sbs-liquidez.md). Los comandos exportan
`liquidez.xlsx`, `cobertura_liquidez.xlsx` y `financiacion_neta_estable.xlsx`, respectivamente.

## Calidad de cartera y provisiones SBS

```powershell
python scripts/sync_calidad_cartera.py --desde 2026-08
python scripts/sync_calidad_cartera.py --desde 2026-08 --datasets calidad categorias morosidad saldos --load-only
```

[Contrato, bases crediticias y unidades](sbs-calidad-cartera.md). El reporte
`outputs/calidad_cartera/calidad_cartera.xlsx` contiene una hoja por fuente seleccionada.

## Fondeo y castigos SBS

```powershell
python scripts/sync_fondeo.py --datasets personas escalas --desde 2026-08
python scripts/sync_fondeo.py --datasets plazos --desde 2026-08 --tipos B F
python scripts/sync_fondeo.py --datasets adeudos --desde 2026-08 --tipos B C R
python scripts/sync_fondeo.py --datasets adeudos --desde 2026-07 --tipos F
python scripts/sync_castigos.py --desde 2026-08
```

[Alcances, fechas originales, unidades y disponibilidad](sbs-fondeo-castigos.md).
Las escalas son agregados del sistema; los castigos son flujos mensuales.

## Rentabilidad y eficiencia SBS

```powershell
python scripts/sync_rentabilidad.py --desde 2026-08
python scripts/sync_rentabilidad.py --desde 2026-08 --datasets rentabilidad --tipos B F
python scripts/sync_rentabilidad.py --desde 2026-08 --datasets eficiencia --tipos C R
```

[Indicadores, anualización, unidades y límites de comparación](sbs-rentabilidad.md).
El reporte es `outputs/rentabilidad_eficiencia/rentabilidad_eficiencia.xlsx`.

## Tamaño y participación SBS

```powershell
python scripts/sync_participacion.py --desde 2026-08
```

[Alcance, unidades y rankings publicados](sbs-participacion.md).
