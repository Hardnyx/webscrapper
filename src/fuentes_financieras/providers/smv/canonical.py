from __future__ import annotations

import re
import unicodedata

import pandas as pd

KNOWN_RENAMES = {
    "Fondo Mutuo": "fund_name",
    "Administradora": "administrator",
    "Fec. Inicio Operación": "operation_start_date",
    "Cuota Moneda": "unit_currency",
    "Cuota Valor": "unit_value",
    "Patrimonio S/.": "assets_pen",
    "Partícipes N": "participants",
    "Fecha_VC": "value_date",
    "Fecha Consulta": "query_date",
    "Fuente": "source",
    "Método Consulta": "transport",
}
CANONICAL_TO_LEGACY = {value: key for key, value in KNOWN_RENAMES.items()}


def _slug(value: object) -> str:
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = text.casefold().replace("%", " pct ")
    text = re.sub(r"[^a-z0-9]+", "_", text).strip("_")
    return text or "column"


def canonicalize_mutual_fund_values(frame: pd.DataFrame) -> pd.DataFrame:
    """Convierte un snapshot SMV a un esquema estable sin perder columnas.

    Las columnas de negocio principales reciben nombres estables. Cualquier
    columna adicional publicada por SMV se conserva con prefijo ``source__``.
    El esquema canónico es el contrato persistente del warehouse; los proyectos
    no deben depender del HTML ni de nombres variables del portal.
    """
    result = frame.copy()
    rename: dict[str, str] = {}
    used: set[str] = set()

    for column in result.columns:
        # Evita volver a canonicalizar una tabla que ya está en el contrato.
        if column in CANONICAL_TO_LEGACY or column == "country_code" or str(column).startswith("source__"):
            target = str(column)
        elif column in KNOWN_RENAMES:
            target = KNOWN_RENAMES[column]
        else:
            target = f"source__{_slug(column)}"
        base = target
        suffix = 2
        while target in used:
            target = f"{base}_{suffix}"
            suffix += 1
        used.add(target)
        rename[column] = target

    result = result.rename(columns=rename)
    result["country_code"] = "PE"

    for column in ("operation_start_date", "value_date", "query_date"):
        if column in result.columns:
            result[column] = pd.to_datetime(
                result[column], errors="coerce", format="mixed", dayfirst=True
            ).dt.normalize()

    for column in ("unit_value", "assets_pen", "participants"):
        if column in result.columns:
            result[column] = pd.to_numeric(result[column], errors="coerce")

    preferred = [
        "country_code",
        "query_date",
        "value_date",
        "fund_name",
        "administrator",
        "operation_start_date",
        "unit_currency",
        "unit_value",
        "assets_pen",
        "participants",
        "source",
        "transport",
    ]
    ordered = [column for column in preferred if column in result.columns]
    ordered += [column for column in result.columns if column not in ordered]
    return result[ordered]


def to_legacy_mutual_fund_schema(frame: pd.DataFrame) -> pd.DataFrame:
    """Proyecta el canónico al contrato histórico usado por el heatmap.

    Es una capa de compatibilidad de lectura, no un segundo almacenamiento. Así
    los consumidores existentes pueden seguir usando ``Fondo Mutuo`` y
    ``Cuota Valor`` mientras la única copia persistente vive en ``canonical/``.
    Columnas canónicas no conocidas se preservan para no perder información.
    """
    result = frame.copy()
    rename = {column: CANONICAL_TO_LEGACY[column] for column in result.columns if column in CANONICAL_TO_LEGACY}
    result = result.rename(columns=rename)

    if "country_code" in result.columns and "País" not in result.columns:
        # Se preserva country_code; no se traduce ni se elimina porque es parte
        # retained as warehouse lineage metadata for generic consumers.
        pass

    for column in ("Fec. Inicio Operación", "Fecha_VC", "Fecha Consulta"):
        if column in result.columns:
            result[column] = pd.to_datetime(result[column], errors="coerce").dt.normalize()
    for column in ("Cuota Valor", "Patrimonio S/.", "Partícipes N"):
        if column in result.columns:
            result[column] = pd.to_numeric(result[column], errors="coerce")
    return result
