# Clasificaciones históricas SBS

Dataset: `pe.sbs.clasificaciones_riesgo`.

Fuente: https://www.sbs.gob.pe/app/iece/paginas/MostrarResumenClasificaciones.aspx

## Ejecutar

```powershell
python scripts/sync_risk_ratings.py
```

El lanzador verifica dependencias y utiliza el comando instalado
`fuentes_financieras.cli.risk_ratings`. El proveedor usa `curl_cffi`, WebForms y HTML;
no utiliza navegador ni descarga PDF o XLS.

1. Descubre todos los períodos publicados.
2. Reutiliza el HTML del período inicial y consulta los períodos faltantes.
3. Extrae todas las entidades y clasificaciones.
4. Guarda Parquet por año y mantiene `state/manifest.json`.
5. Repite `sync()` y exige que no vuelva a descargar períodos.
6. Carga el histórico local y comprueba cobertura y duplicados.

## Almacén y salidas

Los datos usan el almacén resuelto por `fuentes_financieras`, predeterminado
`data/sources`. Para cambiarlo: `--data-root PATH`.

Los reportes se guardan en `outputs/risk_ratings` respecto del directorio de
trabajo. Para cambiarlo: `--output-dir PATH`.

Se mantienen los nombres `resultado_historico_clasificaciones.json` y
`resumen_historico_clasificaciones.csv` como salidas internas de validación.

`--force` reconsulta períodos; `--keep-raw` conserva respuestas para auditoría;
`--no-second-sync` omite la segunda validación de caché.

## API

```python
from fuentes_financieras import source

ratings = source("pe.sbs.clasificaciones_riesgo")
result = ratings.fetch(anio=2026, semestre=1)
ratings.sync()
history = ratings.load(desde="2024-03", hasta="2026-09", tipo_entidad="S")
```

## Contrato canónico

`period_code`, `period`, `period_date`, `year`, `semester`, `entity_type_code`,
`entity_type`, `entity_name`, `rating_agency`, `rating`, `trend`, `source`,
`source_url`, `retrieved_at`.

Las columnas de texto utilizan `string` y las de año/semestre `Int64`.
El contrato v2 permite migrar un almacén parcial de v1 sin borrarlo manualmente.
Los enlaces a informes PDF no forman parte del dataset.
