# Migración de rutas

Reorganización del 7 de octubre de 2026.

| Ruta anterior | Ruta actual |
| --- | --- |
| `SBS/Tasa pasiva/script.py` | `scripts/sync_tasas_pasivas.py` |
| `descargar_historico_clasificaciones.py` | `scripts/sync_clasificaciones_riesgo.py` |
| `SBS/Tipo de Cambio Promedio/script.py` | `apps/sbs/tipo_cambio_promedio.py`, mediante `scripts/run_desktop.py` |
| `SBS/Tipo de Cambio Ponderado/script.py` | `apps/sbs/tipo_cambio_ponderado.py`, mediante `scripts/run_desktop.py` |
| Los tres `empaquetar.py` y sus lanzadores | `scripts/build_app.py --app ...` |
| `core/site_capture/` | `tools/site_capture/`, mediante `scripts/capture_site.py` |
| `git/autopush.sh` | `tools/git/commit_push.sh` |
| `git/dump_filetree.sh` | `tools/git/dump_filetree.sh` |
| `src/fuentes_financieras/README.md` | `docs/architecture.md` y README principal |
| `fuentes_financieras.cli_rates` | `fuentes_financieras.cli.tasas_pasivas`; import anterior compatible |
| `fuentes_financieras.cli_smv` | `fuentes_financieras.cli.valores_cuota`; import anterior compatible |
| `fuentes_financieras.sbs_tipo_cambio` | `fuentes_financieras.providers.sbs.tipo_cambio_contable`; import anterior compatible |

La API `source/fetch/sync/load/describe/list_datasets`, los IDs de datasets y los
almacenes canónicos se mantienen. No se mueven ni eliminan datos locales.

Si se utilizaba `datos_historico`, indicar esa carpeta expresamente:

```powershell
python scripts/sync_clasificaciones_riesgo.py --data-root datos_historico
```

El inventario antiguo `git/filetree.txt` se retira porque ya no reflejaba el
repositorio. La utilidad genera un inventario nuevo dentro de `outputs/`.

Los paquetes originales recuperados para la integración fueron
`fuentes_financieras_codigo_CORREGIDO_V5.zip` y
`fuentes_financieras_clasificaciones_historico_v2.zip`.

Los módulos y lanzadores financieros usan ahora nombres españoles: `tipo_cambio_ponderado.py`, `tipo_cambio_promedio.py`, `tipo_cambio_contable.py`, `universo_depositos.py`, `tasas_pasivas_mercado.py`, `referencias_tasas.py`, y los comandos `sync_tasas_pasivas.py`, `sync_clasificaciones_riesgo.py`, `sync_valores_cuota.py`, `sync_universo_depositos.py`, `sync_referencias_tasas_pasivas.py`. Los imports directos y rutas anteriores de estos archivos deben actualizarse; los IDs de dataset y los nombres de comandos instalados se mantienen.
