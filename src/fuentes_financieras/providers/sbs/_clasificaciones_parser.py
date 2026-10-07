from __future__ import annotations

import re
from datetime import date

import pandas as pd
from bs4 import BeautifulSoup, FeatureNotFound


PERIOD_FIELD = "ctl00$MainContent$DdlPeriodos"
ENTITY_FIELD = "ctl00$MainContent$DdlTiposEntidad"


def _soup(html: str):
    try:
        return BeautifulSoup(html, "lxml")
    except FeatureNotFound:
        return BeautifulSoup(html, "html.parser")


def period_parts(period_code: str) -> tuple[int, int, int, str]:
    code = str(period_code).strip()
    m = re.fullmatch(r"(\d{4})(0[12])", code)
    if not m:
        raise ValueError(f"Código de período SBS inválido: {period_code!r}")

    year = int(m.group(1))
    semester = int(m.group(2))
    month = 3 if semester == 1 else 9
    period = f"{year:04d}-{month:02d}"
    return year, semester, month, period


def period_date(period_code: str) -> str:
    year, semester, month, _ = period_parts(period_code)
    day = 31 if month == 3 else 30
    return date(year, month, day).isoformat()


def select_options(html: str, select_name: str) -> list[dict]:
    s = _soup(html)
    sel = s.find("select", attrs={"name": select_name})
    if sel is None:
        raise ValueError(f"No se encontró select {select_name!r}")

    out = []
    for index, option in enumerate(sel.find_all("option")):
        out.append({
            "index": index,
            "value": option.get("value", ""),
            "text": option.get_text(" ", strip=True),
            "selected": option.has_attr("selected"),
        })
    return out


def selected_period_code(html: str) -> str | None:
    options = select_options(html, PERIOD_FIELD)
    selected = next((x for x in options if x["selected"]), None)
    if selected and selected["value"]:
        return str(selected["value"])
    first = next((x for x in options if x["value"]), None)
    return str(first["value"]) if first else None


def entity_type_map(html: str) -> dict[str, str]:
    options = select_options(html, ENTITY_FIELD)
    return {
        str(x["text"]).strip(): str(x["value"]).strip()
        for x in options
        if str(x["value"]).strip() and str(x["text"]).strip()
    }


def available_periods(html: str) -> list[dict]:
    out = []
    for item in select_options(html, PERIOD_FIELD):
        code = str(item["value"]).strip()
        if not re.fullmatch(r"\d{6}", code):
            continue
        year, semester, month, period = period_parts(code)
        out.append({
            "period_code": code,
            "period": period,
            "period_date": period_date(code),
            "year": year,
            "semester": semester,
            "label": item["text"],
            "selected": bool(item["selected"]),
        })
    return out


def _result_table(html: str):
    s = _soup(html)
    candidates = []

    for table in s.find_all("table"):
        rows = table.find_all("tr", recursive=False)
        if not rows:
            continue

        header_cells = rows[0].find_all(["th", "td"], recursive=False)
        headers = [c.get_text(" ", strip=True) for c in header_cells]

        if (
            len(headers) >= 3
            and "Tipo de Entidad" in headers
            and "Entidad" in headers
        ):
            candidates.append((len(headers), len(rows), table))

    if not candidates:
        raise ValueError("No se encontró la tabla exterior de clasificaciones.")

    candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return candidates[0][2]


def _trend(cell) -> str | None:
    img = cell.find("img")
    if img is None:
        return None

    text = " ".join([
        str(img.get("title") or ""),
        str(img.get("alt") or ""),
        str(img.get("src") or ""),
    ]).lower()

    if any(token in text for token in ("sub", "mejor", "up")):
        return "up"
    if any(token in text for token in ("baj", "deterior", "down")):
        return "down"
    return None


def to_long_form(
    html: str,
    *,
    period_code: str,
    type_code_by_label: dict[str, str],
    source_url: str,
    retrieved_at: str,
) -> pd.DataFrame:
    table = _result_table(html)
    rows = table.find_all("tr", recursive=False)
    if not rows:
        return pd.DataFrame()

    header_cells = rows[0].find_all(["th", "td"], recursive=False)
    headers = [c.get_text(" ", strip=True) for c in header_cells]
    agencies = headers[2:]

    year, semester, month, period = period_parts(period_code)
    pdate = period_date(period_code)

    records: list[dict] = []

    for tr in rows[1:]:
        cells = tr.find_all(["td", "th"], recursive=False)
        if len(cells) < len(headers) or len(cells) < 2:
            continue

        entity_type = cells[0].get_text(" ", strip=True)
        entity_name = cells[1].get_text(" ", strip=True)
        if not entity_type or not entity_name:
            continue

        entity_type_code = type_code_by_label.get(entity_type)

        for idx, agency in enumerate(agencies, start=2):
            if idx >= len(cells):
                break

            cell = cells[idx]
            link = cell.find("a")
            if link is None:
                continue

            rating = link.get_text(" ", strip=True).replace("\xa0", " ").strip()
            if not rating:
                continue

            records.append({
                "period_code": str(period_code),
                "period": period,
                "period_date": pdate,
                "year": year,
                "semester": semester,
                "entity_type_code": entity_type_code,
                "entity_type": entity_type,
                "entity_name": entity_name,
                "rating_agency": str(agency).strip(),
                "rating": rating,
                "trend": _trend(cell),
                "source": "SBS",
                "source_url": source_url,
                "retrieved_at": retrieved_at,
            })

    columns = [
        "period_code",
        "period",
        "period_date",
        "year",
        "semester",
        "entity_type_code",
        "entity_type",
        "entity_name",
        "rating_agency",
        "rating",
        "trend",
        "source",
        "source_url",
        "retrieved_at",
    ]

    df = pd.DataFrame(records, columns=columns)
    if df.empty:
        return df

    # Canonical dtypes are part of the dataset contract.
    # Do not let pandas infer a different dtype merely because a nullable
    # column happens to contain only nulls in one historical period.
    string_columns = [
        "period_code",
        "period",
        "period_date",
        "entity_type_code",
        "entity_type",
        "entity_name",
        "rating_agency",
        "rating",
        "trend",
        "source",
        "source_url",
        "retrieved_at",
    ]
    integer_columns = [
        "year",
        "semester",
    ]

    for column in string_columns:
        df[column] = df[column].astype("string")

    for column in integer_columns:
        df[column] = df[column].astype("Int64")

    # One logical rating per period/entity/rating agency.
    key = [
        "period_code",
        "entity_type",
        "entity_name",
        "rating_agency",
    ]
    df = df.drop_duplicates(subset=key, keep="first")

    return df.sort_values(
        ["period_code", "entity_type", "entity_name", "rating_agency"],
        kind="mergesort",
    ).reset_index(drop=True)
