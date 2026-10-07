# Arquitectura

## Responsabilidades

`src/fuentes_financieras/` es la única implementación compartida de fuentes financieras.

| Ubicación | Responsabilidad |
| --- | --- |
| `api.py`, `catalog.py`, `registry.py` | API pública, catálogo y resolución de proveedores |
| `providers/sbs/` | Tasas pasivas, clasificaciones, tipo de cambio contable y adaptador de curva soberana |
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
| `outputs/risk_ratings/` | Resumen y reporte de validación del histórico de clasificaciones |
| `outputs/site_capture/` | Capturas y recursos de sitios |
| `outputs/` | Exportaciones y otros resultados |
| `dist/` | Ejecutables generados |

Las particiones se verifican contra sus hashes antes de reutilizarse. Los
manifiestos conservan rutas relativas para permitir mover los almacenes. Una
migración de contrato de clasificaciones puede actualizar el esquema sin
eliminar los controles de integridad de la base V5.

## Alcance de validación

Las pruebas son locales y no generan descargas históricas. La compilación de
Python comprueba sintaxis, pero no certifica las interfaces gráficas ni la
compilación de ejecutables en Windows. La curva soberana conserva un error
explícito hasta completar su migración.
