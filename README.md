# webscrapper

Biblioteca personal de fuentes financieras, con extracción, caché y almacenamiento compartidos
mediante el paquete Python `fuentes_financieras`.

## Organización

| Carpeta | Contenido |
| --- | --- |
| [`src/fuentes_financieras/`](src/fuentes_financieras/) | API, proveedores SBS/SMV, almacenamiento y comandos instalables |
| [`scripts/`](scripts/) | Ejecutores con instalación automática de dependencias y empaquetador único |
| [`apps/sbs/`](apps/sbs/) | Aplicaciones de escritorio de tipo de cambio promedio y ponderado |
| [`tools/`](tools/) | Captura de sitios y utilidades de git |
| [`tests/`](tests/) | Pruebas locales de contratos, caché, comandos y estructura |
| [`docs/`](docs/) | Arquitectura, uso y migración de rutas |

`data/`, `outputs/` y `dist/` se generan localmente y están excluidos de git.

## Empezar

Se requiere Python 3.11 o posterior. Los ejecutores de `scripts/` comprueban las
versiones e instalan solo dependencias faltantes o incompatibles con el mismo
intérprete; no requieren instalar previamente el proyecto.

Desde la raíz del repositorio:

```powershell
python scripts/sync_tasas_pasivas.py --tipos B C F R --desde 2026-09-01 --hasta 2026-09-30 --excel outputs/tasas_pasivas.xlsx
python scripts/sync_valores_cuota.py --desde 2026-09-01 --hasta 2026-09-30
python scripts/sync_estados_financieros.py --desde 2026-08
python scripts/sync_solvencia.py --desde 2026-07
python scripts/sync_liquidez.py --desde 2026-08
python scripts/sync_calidad_cartera.py --desde 2026-08
python scripts/sync_fondeo.py --datasets personas escalas --desde 2026-08
python scripts/sync_castigos.py --desde 2026-08
python scripts/sync_rentabilidad.py --desde 2026-08
python scripts/sync_participacion.py --desde 2026-08
python scripts/sync_riesgo_cambiario.py --desde 2026-07
python scripts/sync_riesgo_cambiario.py --datasets capital --desde 2026-08 --tipos C R
python scripts/sync_clasificaciones_riesgo.py
python scripts/sync_universo_depositos.py
python scripts/sync_referencias_tasas_pasivas.py --desde 2026-10-05 --hasta 2026-10-06
```

Las fechas son parámetros de ejemplo. Use `--help` para consultar las opciones
antes de iniciar una descarga histórica.

Para consumir el paquete desde notebooks o instalar los comandos:

```powershell
python -m pip install -e .
fuentes-sync-rates --help
fuentes-sync-smv --help
fuentes-sync-ratings --help
fuentes-sync-universe --help
```

## Fuentes

| Fuente | Acceso | Estado |
| --- | --- | --- |
| SBS riesgo cambiario | `source("pe.sbs.posicion_cambiaria")` / `source("pe.sbs.posicion_cambiaria_capital")` | Posición B/F/C/R en miles de soles; ratio C/R con patrimonio del mes anterior |
| SBS tamaño y participación | `source("pe.sbs.participacion")` | Rankings mensuales de créditos, depósitos y patrimonio B/F/C/R |
| SBS rentabilidad y eficiencia | `source("pe.sbs.rentabilidad")` y `source("pe.sbs.eficiencia")` | ROA/ROE, ratios de gastos y productividad B/F/C/R; períodos y denominadores diferenciados |
| SBS fondeo | `source("pe.sbs.depositos_persona")`, escalas, plazos y adeudos | Personas y adeudos B/F/C/R; escalas agregadas del sistema; plazos B/F |
| SBS castigos | `source("pe.sbs.castigos")` | Flujos mensuales por tipo de crédito B/F/C/R |
| SBS calidad y provisiones | `source("pe.sbs.calidad_cartera")` y tres datasets complementarios | Ratios, categorías, morosidad por días y saldos B/F/C/R |
| SBS liquidez MN/ME | `source("pe.sbs.liquidez")` | Mensual; importes y ratios B/F/C/R |
| SBS cobertura de liquidez | `source("pe.sbs.cobertura_liquidez")` | Promedios trimestrales; período declarado separado del índice |
| SBS financiación neta estable | `source("pe.sbs.financiacion_neta_estable")` | Mensual; importes ponderados y ratio B/F/C/R |
| SBS tasas pasivas B/C/F/R | `source("pe.sbs.tasas_pasivas")` | Implementado; B/F diario, C/R mensual |
| SBS referencias de tasas pasivas | `source("pe.sbs.tasas_pasivas_mercado")` | TIPMN/TIPMEX y FTIPMN/FTIPMEX; promedios por producto desde tasas locales |
| SBS universo de depósitos | `source("pe.sbs.universo_depositos")` | Implementado; capturas fechadas y correspondencias locales |
| SMV valores cuota | `source("pe.smv.fondos_mutuos.valores_cuota")` | Implementado; capturas e histórico EVCP, todos los fondos y series |
| SBS clasificaciones históricas | `source("pe.sbs.clasificaciones_riesgo")` | Implementado; HTML semestral, sin PDF/XLS |
| SBS tipo de cambio contable USD/PEN | `providers.sbs.tipo_cambio_contable` | Implementado; funciones propias |
| SBS curva soberana | `source("pe.sbs.curva_soberana")` | Pendiente de migración; consultas bloqueadas con error explícito |

La implementación disponible y la presencia en el catálogo no certifican una
descarga reciente. Las pruebas locales no consultan SBS ni SMV.

El [universo de depósitos](docs/sbs-universo-depositos.md) incluye un catálogo de
identidades internas y un reporte Excel de correspondencias con tasas y
clasificaciones. No determina elegibilidad regulatoria.

Las [referencias de tasas pasivas](docs/sbs-referencias-tasas-pasivas.md) conservan
separadas las bases sobre saldos y flujos y las ventanas diarias y mensuales.

## Uso desde Automatizaciones

```python
from fuentes_financieras import source

rates = source("pe.sbs.tasas_pasivas")
funds = source("pe.smv.fondos_mutuos.valores_cuota")
ratings = source("pe.sbs.clasificaciones_riesgo")
```

La interfaz compartida es `fetch/sync/load`. Los notebooks seleccionan y analizan
los datos; los proveedores extraen y almacenan.

`FINANCIAL_SOURCES_DATA_ROOT` fija el almacén. Se conserva la detección de
`estructura.json` y de los almacenes de Automatizaciones. En una instalación
local del repositorio, el valor predeterminado es `data/sources`.

## Escritorio y empaquetado

```powershell
python scripts/run_desktop.py --app average-exchange-rate
python scripts/run_desktop.py --app weighted-exchange-rate
python scripts/build_app.py --app passive-rates
```

El empaquetador acepta las tres aplicaciones y genera sus ejecutables en
`dist/`, usando un entorno temporal aislado. Su ejecución no se ha validado en
Windows. Las aplicaciones de escritorio mantienen sus extractores anteriores;
son fuentes distintas del tipo de cambio contable.

## Pruebas

```powershell
python -m pip install -e ".[dev,browser]"
python -m pytest
python -m compileall -q src scripts apps tools
```

## Documentación

- [Arquitectura y almacenamiento](docs/architecture.md)
- [Comandos y aplicaciones](docs/usage.md)
- [Clasificaciones históricas SBS](docs/sbs-clasificaciones-riesgo.md)
- [Migración desde las rutas anteriores](docs/migration.md)

El histórico completo se ejecuta primero localmente. Este repositorio no
programa descargas ni publica datos mediante GitHub Actions.

Los [estados financieros mensuales](docs/sbs-estados-financieros.md) incorporan balance y resultados B/F/C/R, con unidades y cobertura explícitas. Los archivos financieros usan nombres en español; consulta los comandos anteriores actualizados.

La [solvencia mensual](docs/sbs-solvencia.md) incorpora requerimientos, APR, ratios de capital y composición del patrimonio efectivo B/F/C/R. Las unidades ambiguas y las inconsistencias publicadas quedan marcadas.

La [liquidez SBS](docs/sbs-liquidez.md) incorpora liquidez MN/ME, cobertura y financiación neta estable como fuentes independientes, con fechas y escalas verificadas.

La [calidad de cartera](docs/sbs-calidad-cartera.md) incluye provisiones como cobertura porcentual y como saldos del balance, con categorías y umbrales de atraso separados.

El [fondeo y los castigos](docs/sbs-fondeo-castigos.md) incorporan depósitos, escalas de montos del sistema, plazos B/F, adeudos y flujos mensuales de créditos castigados. Las escalas no ofrecen concentración individual por contraparte.

La [rentabilidad y eficiencia](docs/sbs-rentabilidad.md) conserva ratios de utilidad anualizada, bases de comparación y unidades por persona u oficina, sin recalcular ni homogeneizar indicadores distintos.

El [tamaño y participación](docs/sbs-participacion.md) conserva importes, posiciones, porcentajes individuales y acumulados publicados, con su universo de comparación.

El [riesgo cambiario](docs/sbs-riesgo-cambiario.md) conserva balance, derivados, delta de opciones y posición global, además del ratio publicado de cajas con el mes del patrimonio efectivo explícito.
