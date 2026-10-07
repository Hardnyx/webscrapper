from __future__ import annotations

from io import StringIO
import re
import unicodedata

import pandas as pd

from .errors import SMVParseError


def _normalized_text(value: object) -> str:
    text = "" if pd.isna(value) else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", text).strip().lower()


def _parse_smv_dates(values: pd.Series) -> pd.Series:
    """Parsea fechas conocidas de SMV sin recurrir a inferencia ambigua."""
    if pd.api.types.is_datetime64_any_dtype(values):
        return pd.to_datetime(values, errors="coerce")

    text = values.astype("string").str.strip()
    result = pd.Series(pd.NaT, index=values.index, dtype="datetime64[ns]")

    # La página histórica usa normalmente dd/mm/YYYY. Los formatos adicionales
    # cubren datos históricos almacenados o pequeñas variaciones conocidas sin activar dateutil.
    for date_format in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        pending = result.isna() & text.notna() & text.ne("")
        if not pending.any():
            break
        parsed = pd.to_datetime(text.loc[pending], format=date_format, errors="coerce")
        result.loc[pending] = parsed

    return result


def _build_headers(table: pd.DataFrame) -> list[str]:
    if len(table) < 2:
        raise SMVParseError("La tabla de encabezados de SMV está incompleta.")

    headers: list[str] = []
    for column in table.columns:
        first = table.iloc[0, column]
        second = table.iloc[1, column]

        first_text = "" if pd.isna(first) else str(first).strip()
        second_text = "" if pd.isna(second) else str(second).strip()

        if first_text == second_text or not second_text:
            header = first_text
        elif not first_text:
            header = second_text
        else:
            header = f"{first_text} {second_text}"

        headers.append(re.sub(r"\s+", " ", header).strip())

    return headers


def _find_header_table(tables: list[pd.DataFrame]) -> tuple[int, list[str]]:
    for index, table in enumerate(tables):
        if table.empty or len(table) < 2:
            continue

        sample = " ".join(
            _normalized_text(value) for value in table.iloc[:2].to_numpy().ravel()
        )
        if "fondo mutuo" in sample and "administradora" in sample and "cuota" in sample:
            return index, _build_headers(table)

    raise SMVParseError(
        "No se encontró la tabla de encabezados esperada en la respuesta de SMV."
    )


def resolve_fecha_vc(result: pd.DataFrame) -> pd.DataFrame:
    """Identifica la fecha efectiva del valor cuota usando contenido y encabezado.

    La SMV ha presentado variaciones de encabezado donde ``Inf. Atrasada`` queda
    asociado a una columna porcentual y la fecha efectiva aparece en la columna
    contigua. Por ello no se confía únicamente en el nombre del encabezado.
    """
    excluded = {"Fec. Inicio Operación", "Fecha Consulta"}
    current = (
        _parse_smv_dates(result["Fecha_VC"])
        if "Fecha_VC" in result.columns
        else pd.Series(pd.NaT, index=result.index, dtype="datetime64[ns]")
    )
    current_rate = current.notna().mean() if len(current) else 0.0

    if current_rate >= 0.5:
        result["Fecha_VC"] = current
        return result

    candidates: list[tuple[float, str, pd.Series]] = []
    for column in result.columns:
        if column in excluded or column == "Fecha_VC":
            continue

        parsed = _parse_smv_dates(result[column])
        rate = parsed.notna().mean() if len(parsed) else 0.0
        if rate >= 0.5:
            candidates.append((rate, column, parsed))

    if not candidates:
        if "Fecha_VC" in result.columns:
            result["Fecha_VC"] = current
        return result

    candidates.sort(key=lambda item: item[0], reverse=True)
    _, _, parsed = candidates[0]
    result["Fecha_VC"] = parsed
    return result


def _rename_key_columns(frame: pd.DataFrame) -> pd.DataFrame:
    rename: dict[str, str] = {}

    for column in frame.columns:
        normalized = _normalized_text(column)

        if "fondo mutuo" in normalized:
            rename[column] = "Fondo Mutuo"
        elif normalized == "administradora" or "administradora" in normalized:
            rename[column] = "Administradora"
        elif "fec" in normalized and "inicio" in normalized and "operacion" in normalized:
            rename[column] = "Fec. Inicio Operación"
        elif "cuota" in normalized and "moneda" in normalized:
            rename[column] = "Cuota Moneda"
        elif "cuota" in normalized and "valor" in normalized:
            rename[column] = "Cuota Valor"
        elif "variacion" in normalized and "inicio" in normalized:
            rename[column] = "Rentabilidad(A)"
        elif "inf" in normalized and "atrasada" in normalized:
            rename[column] = "Fecha_VC"
        elif "fecha" in normalized and "atrasada" in normalized:
            rename[column] = "Fecha_VC"

    result = frame.rename(columns=rename)

    required = {"Fondo Mutuo", "Cuota Valor"}
    missing = required - set(result.columns)
    if missing:
        raise SMVParseError(
            "La estructura de SMV cambió o no pudo identificarse. "
            f"Columnas requeridas ausentes: {sorted(missing)}"
        )

    return resolve_fecha_vc(result)

def parse_smv_tables(tables: list[pd.DataFrame]) -> pd.DataFrame:
    """Convierte las tablas HTML de SMV en un único DataFrame normalizado."""
    header_index, headers = _find_header_table(tables)
    candidates: list[pd.DataFrame] = []

    for index, table in enumerate(tables):
        if index == header_index or table.empty:
            continue
        category_table = (
            index > header_index + 1
            and len(tables[index - 1].columns) == 1
            and len(tables[index - 1]) <= 2
        )
        if category_table and len(table.columns) != len(headers):
            raise SMVParseError(
                f"La tabla de fondos {index} tiene {len(table.columns)} columnas; "
                f"se esperaban {len(headers)}."
            )
        if len(table.columns) != len(headers):
            continue

        candidate = table.copy()
        candidate.columns = headers

        # Ignora encabezados repetidos que pueden aparecer dentro del HTML.
        first_row = " ".join(
            _normalized_text(value) for value in candidate.iloc[0].tolist()
        )
        if "fondo mutuo" in first_row and "administradora" in first_row:
            continue

        candidates.append(candidate)

    if not candidates:
        raise SMVParseError("SMV no devolvió tablas de fondos con la estructura esperada.")

    frame = pd.concat(candidates, ignore_index=True)
    frame = frame.dropna(how="all").reset_index(drop=True)
    frame = _rename_key_columns(frame)

    frame["Cuota Valor"] = pd.to_numeric(frame["Cuota Valor"], errors="coerce")
    if "Fecha_VC" in frame.columns:
        frame["Fecha_VC"] = _parse_smv_dates(frame["Fecha_VC"])

    return frame


def parse_smv_html(html: str) -> pd.DataFrame:
    """Lee la respuesta HTML de SMV y devuelve las filas de fondos mutuos."""
    if not html.strip():
        raise SMVParseError("El HTML de SMV está vacío.")

    try:
        tables = pd.read_html(StringIO(html), decimal=".", thousands=",")
    except ValueError as error:
        raise SMVParseError("SMV no devolvió tablas HTML legibles.") from error

    return parse_smv_tables(tables)
