# webscrapper

Scrapers de fuentes financieras para FUENTES. El paquete compartido `fuentes_financieras`
se instala desde este repositorio; los scripts históricos en `SBS/` siguen disponibles.

## Instalación

Python 3.11 o posterior:

```powershell
python -m pip install -e .
```

Para pruebas locales:

```powershell
python -m pip install -e ".[dev]"
python -m pytest
```

## Fuentes integradas

| Fuente | Acceso | Estado |
| --- | --- | --- |
| SBS tasas pasivas B/C/F/R | `source("pe.sbs.tasas_pasivas")` | Implementado; B/F diario, C/R mensual |
| SMV valores cuota, todos los fondos y series | `source("pe.smv.fondos_mutuos.valores_cuota")` | Implementado; capturas e histórico EVCP |
| SBS tipo de cambio contable USD/PEN | `fuentes_financieras.sbs_tipo_cambio` | Implementado; funciones propias |
| SBS curva soberana | `source("pe.sbs.curva_soberana")` | Adaptador pendiente; las consultas lanzan un error explícito |

La presencia en el catálogo no certifica una descarga reciente. Las pruebas
locales verifican integración y contratos; el acceso real depende de SBS/SMV.

## Uso desde Automatizaciones

Instalar este repositorio en el mismo entorno Python de los notebooks permite:

```python
from fuentes_financieras import source

rates = source("pe.sbs.tasas_pasivas")
funds = source("pe.smv.fondos_mutuos.valores_cuota")
```

La interfaz común es `fetch/sync/load`. Los notebooks consumen datos y
seleccionan instrumentos; la extracción permanece aquí.

`FINANCIAL_SOURCES_DATA_ROOT` permite reutilizar el almacén existente. También se
conserva la detección de `estructura.json` y de los almacenes de Automatizaciones.
Fuera de ese proyecto, los datos se guardan en `data/sources` de este repositorio.
Los datos y las respuestas descargadas están excluidos de git.

## SMV

Consultar las opciones antes de definir el período de descarga:

```powershell
python -m fuentes_financieras.cli_smv --help
python -m fuentes_financieras.cli_smv --desde 2026-09-01 --hasta 2026-09-30
```

El histórico completo se ejecuta primero localmente. Este repositorio todavía
no configura descargas programadas ni publica datos mediante GitHub Actions.

## Tipo de cambio contable

```python
from fuentes_financieras.sbs_tipo_cambio import get_accounting_exchange_rate

rate = get_accounting_exchange_rate(
    "2026-09-30", storage_dir="data/sources/peru/sbs/tipo_cambio_contable"
)
print(rate["date"], rate["usd_pen_accounting"])
```

La fecha efectiva devuelta puede ser anterior a la fecha solicitada.
