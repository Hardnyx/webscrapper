from __future__ import annotations

import re
from datetime import date
from urllib.parse import urljoin, urlsplit, parse_qs

from fuentes_financieras.exceptions import SchemaChangedError

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

    if len(candidates) != 1:
        raise SchemaChangedError('Tabla exterior de clasificaciones ambigua.')
    return candidates[0][2]


def _trend(cell, source_url):
    images = cell.find_all('img')
    if not images:
        return None, '', '', '', ''
    if len(images) != 1:
        raise SchemaChangedError('Múltiples símbolos de cambio en una clasificación.')
    img = images[0]
    title, alt, src = (str(img.get(k) or '').strip() for k in ('title', 'alt', 'src'))
    filename = urlsplit(src).path.rsplit('/', 1)[-1].casefold()
    known = {'subio.png': ('up', 'Subió en relación a la clasificación anterior'),
             'bajo.png': ('down', 'Bajó en relación a la clasificación anterior')}
    trend = None
    flag = 'source_change_symbol_unrecognized'
    if filename in known:
        trend, expected = known[filename]
        if title and title != expected:
            raise SchemaChangedError('Símbolo y texto de cambio contradictorios.')
        flag = ''
    return trend, urljoin(source_url, src) if src else '', title, alt, flag


def report_reference(href, period_code, source_url):
    if not href:
        return dict(report_url='', report_agency_code='', report_period_code='',
                    report_file_number='', report_version='', report_id=''), 'source_report_link_missing'
    url = urljoin(source_url, href)
    parts = urlsplit(url)
    if (parts.scheme != 'https' or parts.netloc != 'extranet.sbs.gob.pe'
            or parts.path != '/iece/descargar' or parts.fragment):
        raise SchemaChangedError('Destino del informe SBS no reconocido.')
    params = parse_qs(parts.query, keep_blank_values=True)
    keys = ('codClasificadora', 'codPeriodo', 'numArchivo', 'numVersion')
    if set(params) != set(keys) or any(len(params[k]) != 1 or not re.fullmatch(r'\d+', params[k][0]) for k in keys):
        raise SchemaChangedError('Identificadores de informe SBS inválidos o ambiguos.')
    agency, period, number, version = (params[k][0] for k in keys)
    if period != period_code or any(int(v) <= 0 for v in (agency, number, version)):
        raise SchemaChangedError('Informe de otro período o identificador no positivo.')
    return dict(report_url=url, report_agency_code=agency, report_period_code=period,
                report_file_number=number, report_version=version,
                report_id=':'.join((agency, period, number, version))), ''


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
    if headers[:2] != ['Tipo de Entidad', 'Entidad'] or any(not h for h in headers[2:]) or len(set(headers[2:])) != len(headers[2:]):
        raise SchemaChangedError('Encabezado de clasificadoras cambiado o duplicado.')
    agencies = headers[2:]

    year, semester, month, period = period_parts(period_code)
    pdate = period_date(period_code)

    records: list[dict] = []

    for tr in rows[1:]:
        cells = tr.find_all(["td", "th"], recursive=False)
        if len(cells) == 1 and cells[0].get('colspan') == str(len(headers)) and not cells[0].get_text(strip=True) and not cells[0].find(['a', 'img']):
            continue
        if len(cells) != len(headers):
            raise SchemaChangedError('Fila de clasificaciones incompleta o con columnas nuevas.')

        entity_type = cells[0].get_text(" ", strip=True)
        entity_name = cells[1].get_text(" ", strip=True)
        if not entity_type or not entity_name:
            raise SchemaChangedError('Clasificación sin tipo de entidad o nombre.')

        entity_type_code = type_code_by_label.get(entity_type)

        for idx, agency in enumerate(agencies, start=2):
            if idx >= len(cells):
                break

            cell = cells[idx]
            links = cell.find_all('a')
            if not links:
                if cell.get_text(strip=True) or cell.find('img'):
                    raise SchemaChangedError('Contenido de clasificación sin enlace reconocible.')
                continue
            if len(links) != 1:
                raise SchemaChangedError('Múltiples clasificaciones en una misma celda.')
            link = links[0]

            rating = link.get_text(" ", strip=True).replace("\xa0", " ").strip()
            if not rating:
                raise SchemaChangedError('Enlace de clasificación sin texto.')
            reference, report_flag = report_reference(link.get('href'), str(period_code), source_url)
            trend, icon_url, icon_title, icon_alt, trend_flag = _trend(cell, source_url)

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
                "trend": trend,
                "rating_kind": "institutional_summary",
                "trend_basis": "published_change_vs_previous_classification",
                "source_change_icon_url": icon_url,
                "source_change_title": icon_title,
                "source_change_alt": icon_alt,
                "data_quality_flags": ';'.join(filter(None, (report_flag, trend_flag))),
                **reference,
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
        'rating_kind',
        'trend_basis',
        'source_change_icon_url',
        'source_change_title',
        'source_change_alt',
        'data_quality_flags',
        'report_url',
        'report_agency_code',
        'report_period_code',
        'report_file_number',
        'report_version',
        'report_id',
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
        'rating_kind',
        'trend_basis',
        'source_change_icon_url',
        'source_change_title',
        'source_change_alt',
        'data_quality_flags',
        'report_url',
        'report_agency_code',
        'report_period_code',
        'report_file_number',
        'report_version',
        'report_id',
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
    if df.duplicated(subset=key).any():
        raise SchemaChangedError('Clasificaciones duplicadas para período, entidad y clasificadora.')

    return df.sort_values(
        ["period_code", "entity_type", "entity_name", "rating_agency"],
        kind="mergesort",
    ).reset_index(drop=True)
