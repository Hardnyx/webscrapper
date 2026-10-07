# SBS Clasificaciones de Riesgo — histórico completo

Dataset:

```text
pe.sbs.clasificaciones_riesgo
```

Fuente:

```text
https://www.sbs.gob.pe/app/iece/paginas/MostrarResumenClasificaciones.aspx
```

## Ejecutar

```powershell
python descargar_historico_clasificaciones.py
```

No usa Playwright ni Selenium.

El extractor trabaja únicamente con:

```text
curl_cffi + ASP.NET WebForms + HTML/updatePanel
```

No descarga:

```text
PDF
resumen.xls
```

## Qué hace

1. GET inicial a SBS.
2. Descubre automáticamente todos los períodos publicados.
3. Usa directamente el HTML del período que ya viene en el GET.
4. Hace un POST por cada período histórico restante, con `Tipo Entidad=""`.
5. Extrae todas las entidades y clasificaciones del HTML devuelto.
6. Guarda el histórico en Parquet particionado por año.
7. Mantiene `state/manifest.json` para evitar redescargas.
8. Repite el mismo `sync()` y exige `downloaded=0`.
9. Ejecuta `load()` local y comprueba cobertura completa.

## Salidas

```text
datos_historico/
└── peru/
    └── sbs/
        └── clasificaciones_riesgo/
            ├── state/
            │   └── manifest.json
            └── canonical/
                ├── year=2012/
                │   └── data.parquet
                ├── year=2013/
                └── ...

resultado_historico_clasificaciones.json
resumen_historico_clasificaciones.csv
```

## Uso posterior

```python
from fuentes_financieras import source

clasif = source("pe.sbs.clasificaciones_riesgo")
```

Consulta puntual:

```python
resultado = clasif.fetch(periodo="202601")
df = resultado.data
```

También:

```python
resultado = clasif.fetch(anio=2026, semestre=1)
```

Actualizar todo lo disponible:

```python
clasif.sync()
```

Cargar exclusivamente desde almacenamiento local:

```python
df = clasif.load()
```

Filtrar localmente:

```python
df = clasif.load(
    desde="2024-03",
    hasta="2026-09",
    tipo_entidad="S",
)
```

## Esquema canónico

```text
period_code
period
period_date
year
semester
entity_type_code
entity_type
entity_name
rating_agency
rating
trend
source
source_url
retrieved_at
```

Los enlaces a informes PDF no forman parte de esta versión.


## Corrección v2 — estabilidad de esquema

La primera versión calculaba `schema_hash` usando los dtypes inferidos por
pandas. En pandas 3 una columna nullable puede inferirse de forma distinta entre
períodos aunque el contrato lógico sea idéntico.

La v2:

- fija explícitamente los dtypes canónicos;
- usa `string` para columnas textuales/nullable;
- usa `Int64` para `year` y `semester`;
- asocia `last_schema_hash` a `contract_version`;
- migra automáticamente un `datos_historico` parcial generado por v1;
- no requiere borrar manualmente los datos ya descargados.

Puede ejecutarse sobre la misma carpeta:

```powershell
python descargar_historico_clasificaciones.py
```
