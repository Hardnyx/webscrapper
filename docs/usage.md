# Comandos y aplicaciones

Ejecutar desde la raíz del repositorio, salvo que se indique otra ruta. Los
lanzadores resuelven sus archivos respecto del repositorio y también pueden
invocarse mediante una ruta absoluta desde otro directorio.

## Tasas pasivas

```powershell
python scripts/sync_passive_rates.py --tipos B C F R --desde 2026-09-01 --hasta 2026-09-30 --excel outputs/tasas_pasivas.xlsx
```

B/F se consultan por día hábil; C/R por mes. `--load-only` lee únicamente el
almacén local. `--force` solicita revalidación explícita. Una sincronización
incompleta termina con código 1 y no exporta un Excel que parezca completo.
La exportación incluye tabla, filtros, estilo claro 9 y encabezado inmovilizado.

## Valores cuota SMV

```powershell
python scripts/sync_fund_values.py --help
```

El comando admite rangos de capturas, migración desde datos anteriores y
backfill histórico. Se reutilizan particiones válidas; el reporte consumidor
filtra los instrumentos después de la extracción.

## Clasificaciones históricas

```powershell
python scripts/sync_risk_ratings.py --data-root data/sources --output-dir outputs/risk_ratings
```

[Contrato y validación](sbs-risk-ratings.md).

## Tipo de cambio contable

```python
from fuentes_financieras.providers.sbs.accounting_exchange_rate import get_accounting_exchange_rate

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
