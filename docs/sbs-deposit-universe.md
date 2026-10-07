# Universo de entidades y correspondencias

Dataset: `pe.sbs.universo_depositos`.

Fuente: https://www.sbs.gob.pe/app/pp/empresasweb/Paginas/EmpCaptarDep.aspx

Se adapta el parser del capturador v4. Extrae B/F/C/R y excluye las cooperativas.
Valida las cuatro categorías sin imponer sus conteos actuales. Una respuesta
incompleta o un bloqueo produce un error y no reemplaza los datos válidos.

## Ejecutar

```powershell
python scripts/sync_deposit_universe.py
python scripts/sync_deposit_universe.py --load-only
python scripts/sync_deposit_universe.py --html captura.html --observed-on 2026-10-07
```

El lanzador comprueba dependencias. `--data-root PATH` selecciona el almacén;
`--output-dir PATH` selecciona las salidas. La importación requiere la fecha real
de captura, no la fecha de procesamiento. La fuente solo publica el universo
vigente: `fetch()` y `sync()` no aceptan fechas históricas. Las consultas locales
admiten `desde`, `hasta`, `tipo` y `latest=True`.

El caché conserva una observación por fecha de Lima y particiones por año,
HTML y manifiestos. Dentro de la misma fecha, `--force` reemplaza la observación
previa; no es un registro de todas las versiones intradiarias.

```python
from fuentes_financieras import source
universe = source('pe.sbs.universo_depositos')
universe.sync(keep_raw=True)
current = universe.load(latest=True)
```

Contrato: `period`, `period_date`, `entity_type_code`, `entity_type`,
`entity_name`, `normalized_name`, `source_order`, `source`, `source_url`,
`retrieved_at`. La fecha identifica cuándo se observó la lista, no la fecha
legal de autorización.

## Identidades y equivalencias

El catálogo se guarda en `reference/entity_catalog.json` dentro del almacén.
Los identificadores `internal-…` son internos, no códigos oficiales SBS.
Se conservan entre ejecuciones. Por defecto se reconoce el mismo tipo y nombre
normalizado, ignorando acentos, puntuación y espacios. No se eliminan palabras
del nombre ni se realizan asignaciones por similitud.

Los alias revisados se pasan con `--aliases equivalencias.json`:

```json
[
  {
    "dataset": "pe.sbs.clasificaciones_riesgo",
    "entity_type_code": "B",
    "alias": "NOMBRE PUBLICADO EN CLASIFICACIONES",
    "entity_id": "COPIAR EL IDENTIFICADOR DE LA HOJA IDENTIDADES",
    "valid_from": "2026-01-01",
    "valid_to": null,
    "evidence": "URL o referencia documental que acredita la equivalencia"
  }
]
```

La equivalencia requiere evidencia y fecha de vigencia. Puede aplicarse a un
nombre cambiado en el universo usando `dataset: pe.sbs.universo_depositos`.
Un cambio de tipo también requiere una equivalencia explícita. Dos identidades
posibles producen `ambiguous`; ninguna produce `unmatched`. El cruce con datos
históricos reconoce identidad, sin inferir autorización en la fecha del dato.
Los identificadores no deben usarse como llave universal entre catálogos
independientes sin compartir el catálogo y sus equivalencias.

## Reporte

`outputs/deposit_universe/universo_correspondencias.xlsx` incluye:

- Universo más reciente e historial de capturas.
- Identidades internas y fechas de primera/última observación.
- Correspondencias por dataset, entidad y fecha, conservando no resueltos.
- Apariciones/desapariciones entre listas observadas, sin calificarlas como
  eventos legales de autorización o intervención.
- Cobertura local: un dataset sin datos queda expresamente sin evaluar.

Los cruces leen únicamente el caché de tasas y clasificaciones, sin consultar
esos sitios ni descargar su histórico. Todas las hojas tienen filtros y las
hojas con datos usan `TableStyleLight9`. No se calcula elegibilidad regulatoria,
score ni rating conservador: corresponde al proyecto consumidor.
