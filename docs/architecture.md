# Arquitectura

## Responsabilidades

`src/fuentes_financieras/` es la única implementación compartida de fuentes financieras.

| Ubicación | Responsabilidad |
| --- | --- |
| `api.py`, `catalog.py`, `registry.py` | API pública, catálogo y resolución de proveedores |
| `providers/sbs/` | Universo de depósitos, tasas pasivas, clasificaciones, tipo de cambio contable y adaptador de curva soberana |
| `providers/smv/` | Valores cuota, descubrimiento, capturas e histórico |
| `transports/` | Clientes HTTP reutilizables |
| `provider.py`, `storage.py`, `runtime.py` | Contrato, caché, manifiestos y ubicación de los datos |
| `cli/` | Comandos instalables; sin instalación automática de dependencias |
| `sbs/` | Fachadas públicas de consulta para consumidores existentes |

Los módulos `cli_rates.py`, `cli_smv.py` y `sbs_tipo_cambio.py` son puntos de
compatibilidad: delegan en la implementación organizada, sin duplicarla.

`scripts/_bootstrap.py` comparte la comprobación de dependencias de los
lanzadores. `scripts/build_app.py` concentra el empaquetado de las tres
aplicaciones SBS. `apps/` contiene interfaces de escritorio; `tools/` contiene
utilidades independientes de la API financiera.

## Almacenamiento

La prioridad y selección entre almacenes existentes de Automatizaciones se
mantiene en `runtime.py`. `FINANCIAL_SOURCES_DATA_ROOT` permite escoger una ruta
explícita. El marcador `.fuentes_financieras_root` identifica este repositorio cuando
se utiliza fuera de Automatizaciones.

| Directorio local | Contenido |
| --- | --- |
| `data/sources/` | Particiones canónicas, respuestas y manifiestos de fuentes |
| `outputs/clasificaciones_riesgo/` | Resumen y reporte de validación del histórico de clasificaciones |
| `outputs/universo_depositos/` | Universo, identidades y correspondencias con datasets locales |
| `outputs/referencias_tasas_pasivas/` | Referencias generales y promedios publicados por producto |
| `outputs/site_capture/` | Capturas y recursos de sitios |
| `outputs/` | Exportaciones y otros resultados |
| `dist/` | Ejecutables generados |

Las particiones se verifican contra sus hashes antes de reutilizarse. Los
manifiestos conservan rutas relativas para permitir mover los almacenes. Una
migración de contrato de clasificaciones puede actualizar el esquema sin
eliminar los controles de integridad de la base V5.

`entities.py` resuelve correspondencias sobre datos proporcionados por el
consumidor. Los proveedores no se llaman entre sí. El comando de universo lee
los caches locales de tasas y clasificaciones y exporta un reporte sin reglas
de elegibilidad. `reference/entity_catalog.json` conserva identidades internas
dentro del almacén y los alias explícitos tienen alcance de dataset y fechas.

`referencias_tasas.py` extrae promedios publicados de un DataFrame de tasas entregado
por el consumidor. `pe.sbs.tasas_pasivas_mercado` consulta su propia página;
el comando combina sus referencias generales con los promedios locales.
Conserva bases sobre saldos/flujos, grupos y ventanas, sin calcular spreads.

## Alcance de validación

Las pruebas son locales y no generan descargas históricas. La compilación de
Python comprueba sintaxis, pero no certifica las interfaces gráficas ni la
compilación de ejecutables en Windows. La curva soberana conserva un error
explícito hasta completar su migración.

`pe.sbs.estados_financieros` publica cuentas mensuales en formato largo y conserva sus originales Excel y notas. No consulta otros proveedores ni calcula ratios. El reporte se escribe en `outputs/estados_financieros/`.

`_excel_mensual.py` concentra descubrimiento de enlaces oficiales, identificación del formato Excel, períodos, transporte y caché. Los parsers de estados financieros y solvencia siguen independientes. El comando de solvencia solicita sus dos datasets explícitamente y escribe en `outputs/solvencia/`.
