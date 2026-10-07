from __future__ import annotations

import re
import unicodedata
import numpy as np
import pandas as pd
from bs4 import BeautifulSoup, Tag, FeatureNotFound

from fuentes_financieras.exceptions import SchemaChangedError

KNOWN_TABLE_IDS = {
    ("MN", "general"): "ctl00_cphContent_rpgActualPrimTablaMn_OT",
    ("MN", "persona"): "ctl00_cphContent_rpgActualMn_OT",
    ("ME", "general"): "ctl00_cphContent_rpgActualPrimTablaMex_OT",
    ("ME", "persona"): "ctl00_cphContent_rpgActualMex_OT",
}

DESIRED_COLS = [
    "Banco",
    "Depósitos de Ahorro",
    "Hasta 30 días",
    "31-90 días",
    "91-180 días",
    "181-360 días",
    "Más de 360 días",
    "Depósitos a Plazo",
    "Depósitos CTS",
]

PERSON_TERM_COLS = [
    "Hasta 30 días",
    "31-90 días",
    "91-180 días",
    "181-360 días",
    "Más de 360 días",
]


def _make_soup(html: str):
    try:
        return BeautifulSoup(html, "lxml")
    except FeatureNotFound:
        return BeautifulSoup(html, "html.parser")


def _norm(s: str) -> str:
    import unicodedata
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

def build_header_with_spans(table: Tag):
    thead = table.find("thead")
    if not thead:
        return None
    trs = thead.find_all("tr", recursive=False) or thead.find_all("tr")
    if not trs:
        return None

    total_cols = 0
    for cell in trs[0].find_all(["th", "td"], recursive=False):
        try:
            total_cols += int(cell.get("colspan", 1))
        except Exception:
            total_cols += 1
    if total_cols <= 0:
        return None

    grid = [["" for _ in range(total_cols)] for _ in range(len(trs))]

    for r_idx, tr in enumerate(trs):
        c = 0
        for cell in tr.find_all(["th", "td"], recursive=False):
            while c < total_cols and grid[r_idx][c]:
                c += 1
            if c >= total_cols:
                break
            text = cell.get_text(" ", strip=True)
            try:
                colspan = int(cell.get("colspan", 1))
                rowspan = int(cell.get("rowspan", 1))
            except Exception:
                colspan = rowspan = 1
            for rr in range(rowspan):
                for cc in range(colspan):
                    if r_idx + rr < len(trs) and c + cc < total_cols:
                        if not grid[r_idx + rr][c + cc]:
                            grid[r_idx + rr][c + cc] = text
            c += colspan

    names = []
    for col in range(total_cols):
        parts = []
        for row in grid:
            v = row[col].strip()
            if v and (not parts or parts[-1] != v):
                parts.append(v)
        names.append(" - ".join(parts) if parts else f"col_{col}")
    return names

def extract_inner_data(table: Tag):
    tbody = table.find("tbody")
    rows = tbody.find_all("tr", recursive=False) if tbody else table.find_all("tr")
    data = []
    for tr in rows:
        tds = tr.find_all("td", recursive=False)
        if not tds:
            continue
        cells = [td.get_text("\n", strip=True) for td in tds]
        if any(c.strip() for c in cells):
            data.append(cells)
    return data

def extract_banks(main_table: Tag):
    bank_cells = main_table.find_all("td", class_=lambda c: c and "rpgRowHeaderField" in c)
    banks = [td.get_text(" ", strip=True) for td in bank_cells if td.get_text(" ", strip=True)]
    if banks:
        return banks

    # Fallback si Telerik cambia las clases: filas externas de una sola celda textual.
    out = []
    for tr in main_table.find_all("tr"):
        tds = tr.find_all("td", recursive=False)
        if len(tds) == 1:
            txt = tds[0].get_text(" ", strip=True)
            if txt and not re.fullmatch(r"[-\d., %]+", txt):
                out.append(txt)
    return out

def clean_num(x):
    if x is None:
        return np.nan
    s = str(x).strip()
    if s in ("", "-", "--", "NA", "N/A"):
        return np.nan
    # Las tasas de SBS usan punto decimal; la coma, si aparece, se toma como miles.
    s2 = s.replace(",", "").replace(" ", "")
    try:
        return float(s2)
    except Exception:
        return s

def _largest_inner_table(main_tbl: Tag) -> Tag:
    nested = main_tbl.find_all("table")
    if not nested:
        return main_tbl
    return max(nested, key=lambda t: sum(len(tr.find_all("td")) for tr in t.find_all("tr")))

def _table_signature(table: Tag) -> str:
    head = table.find("thead")
    return _norm(head.get_text(" ", strip=True) if head else table.get_text(" ", strip=True)[:1500])

def _currency_from_context(table: Tag) -> str | None:
    node: Tag | None = table
    for _ in range(8):
        if node is None:
            break
        ident = _norm(" ".join(filter(None, [node.get("id", ""), " ".join(node.get("class", []))])))
        if "mex" in ident or "extranj" in ident:
            return "ME"
        if "nacional" in ident or "mn" in ident:
            return "MN"
        node = node.parent if isinstance(node.parent, Tag) else None
    return None

def _kind_from_headers(table: Tag) -> str | None:
    sig = _table_signature(table)
    if ("persona natural" in sig or "personas naturales" in sig) and (
        "persona juridica" in sig or "personas juridicas" in sig
    ):
        return "persona"
    if "depositos de ahorro" in sig and "depositos cts" in sig and "depositos a plazo" in sig:
        return "general"

    # Fallback estructural: la tabla general SBS tiene 8 valores por entidad.
    # Especially useful in ME when Telerik leaves upper headers empty.
    try:
        inner = _largest_inner_table(table)
        rows = extract_inner_data(inner)
        widths = [len(r) for r in rows if r]
        width = max(widths, default=0)
        term_hits = sum(k in sig for k in ("hasta 30", "31-90", "91-180", "181-360", "mas de 360"))
        if width == 8 and term_hits >= 3:
            return "general"
        if width == 10 and ("natural" in sig or "jurid" in sig):
            return "persona"
    except Exception:
        pass
    return None

def discover_main_tables(html_text: str) -> dict[tuple[str, str], Tag]:
    soup = _make_soup(html_text)
    found: dict[tuple[str, str], Tag] = {}

    # 1) IDs conocidos: rápidos y exactos.
    for key, table_id in KNOWN_TABLE_IDS.items():
        t = soup.find("table", id=table_id)
        if t is not None:
            found[key] = t

    if len(found) == 4:
        return found

    # 2) Fallback semántico. Examina tablas principales con encabezados esperados.
    candidates = []
    for t in soup.find_all("table"):
        kind = _kind_from_headers(t)
        if kind is None:
            continue
        currency = _currency_from_context(t)
        if currency is None:
            # Buscar texto/IDs en una ventana superior más amplia.
            parent_text = _norm(t.parent.get_text(" ", strip=True)[:2000]) if isinstance(t.parent, Tag) else ""
            if "moneda extranjera" in parent_text:
                currency = "ME"
            elif "moneda nacional" in parent_text:
                currency = "MN"
        if currency:
            candidates.append((currency, kind, t))

    # Preferir la tabla externa que contiene encabezados de bancos.
    for currency, kind, t in candidates:
        key = (currency, kind)
        if key in found:
            continue
        parent = t
        best = t
        for _ in range(10):
            if not isinstance(parent.parent, Tag):
                break
            parent = parent.parent
            if parent.name == "table" and extract_banks(parent):
                best = parent
                break
        found[key] = best

    return found

def parse_main_table(main_tbl: Tag) -> pd.DataFrame:
    inner = _largest_inner_table(main_tbl)
    headers = build_header_with_spans(inner)
    rows = extract_inner_data(inner)
    banks = extract_banks(main_tbl)

    if not headers:
        maxcols = max((len(r) for r in rows), default=0)
        headers = [f"col_{i}" for i in range(maxcols)]

    rows_clean = []
    for row in rows:
        rows_clean.append([c.split("\n")[0].strip() if isinstance(c, str) else c for c in row])

    n_header = len(headers)
    max_len = max((len(r) for r in rows_clean), default=0)
    # Telerik a veces incluye una cabecera lógica adicional sin celda de datos.
    if max_len == n_header - 1 and headers:
        h0 = _norm(headers[0])
        if not h0 or any(k in h0 for k in ("empresa", "entidad", "sistema")):
            headers = headers[1:]
            n_header -= 1
    elif max_len == n_header + 1:
        headers = headers[1:]
        rows_clean = [r[1:] for r in rows_clean]
        n_header -= 1

    fixed = []
    for r in rows_clean:
        rr = list(r[:n_header])
        rr += [""] * max(0, n_header - len(rr))
        fixed.append(rr)

    df = pd.DataFrame(fixed, columns=headers)
    if banks:
        # Evitar incluir "Promedio" si la tabla de datos no lo trae o viceversa.
        if len(banks) == len(df):
            df.insert(0, "Banco", banks)
        elif len(banks) == len(df) + 1 and _norm(banks[-1]) == "promedio":
            df.insert(0, "Banco", banks[:-1])
        else:
            raise SchemaChangedError(
                f"No se pudo alinear entidades ({len(banks)}) con filas ({len(df)})."
            )
    else:
        raise SchemaChangedError("No se pudieron identificar los bancos/entidades de la tabla.")

    for c in df.columns:
        if c != "Banco":
            df[c] = df[c].map(clean_num)
    return df.reset_index(drop=True)

def _find_col(existing: list[str], aliases: list[str]) -> str | None:
    norm_cols = [(_norm(c), c) for c in existing]
    # Primero coincidencia exacta normalizada, luego contiene.
    for alias in aliases:
        na = _norm(alias)
        for nc, original in norm_cols:
            if nc == na:
                return original
    for alias in aliases:
        na = _norm(alias)
        for nc, original in norm_cols:
            if na in nc:
                return original
    return None

def harmonize_general(df: pd.DataFrame) -> pd.DataFrame:
    existing = list(df.columns)
    data_cols = [c for c in existing if c != "Banco"]

    # Estructura contractual del cuadro SBS: 8 valores por entidad.
    # En ME, Telerik deja vacíos varios encabezados del primer nivel; por eso la
    # posición es más estable que el texto visible para estas 8 columnas.
    if "Banco" in existing and len(data_cols) == 8:
        canonical = DESIRED_COLS[1:]
        out = pd.DataFrame({"Banco": df["Banco"].astype(str).values})
        for dest, src in zip(canonical, data_cols):
            out[dest] = df[src].values
        return out[DESIRED_COLS]

    # Fallback semántico si SBS agrega/reordena columnas en el futuro.
    mapping = {
        "Banco": _find_col(existing, ["Banco"]),
        "Depósitos de Ahorro": _find_col(existing, ["Depósitos de Ahorro", "Ahorro"]),
        "Hasta 30 días": _find_col(existing, ["Hasta 30 días", "0-30"]),
        "31-90 días": _find_col(existing, ["31-90 días", "31 - 90"]),
        "91-180 días": _find_col(existing, ["91-180 días", "91 - 180"]),
        "181-360 días": _find_col(existing, ["181-360 días", "181 - 360"]),
        "Más de 360 días": _find_col(existing, ["Más de 360 días", "Mas de 360 dias"]),
        "Depósitos CTS": _find_col(existing, ["Depósitos CTS", "CTS"]),
    }
    mapping["Depósitos a Plazo"] = next(
        (c for c in existing if _norm(c) == "depositos a plazo"), None
    )
    missing = [c for c in DESIRED_COLS if not mapping.get(c)]
    if missing:
        raise SchemaChangedError(f"Columnas generales no reconocidas: {missing}")
    return pd.DataFrame({dest: df[src].values for dest, src in mapping.items()})[DESIRED_COLS]

def split_person_table(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    if "Banco" not in df.columns:
        raise SchemaChangedError("La tabla por persona no contiene Banco.")
    cols = [c for c in df.columns if c != "Banco"]
    nat_cols = [c for c in cols if "natural" in _norm(c)]
    jur_cols = [c for c in cols if "jurid" in _norm(c)]

    # Fallback por posición solo cuando hay exactamente 10 tramos esperados.
    if not nat_cols and not jur_cols and len(cols) == 10:
        nat_cols, jur_cols = cols[:5], cols[5:]
    if len(nat_cols) != 5 or len(jur_cols) != 5:
        raise SchemaChangedError(
            f"Estructura por persona inesperada: naturales={len(nat_cols)}, jurídicas={len(jur_cols)}"
        )

    def build(subcols: list[str]) -> pd.DataFrame:
        out = pd.DataFrame({"Banco": df["Banco"].astype(str).values})
        for dest in PERSON_TERM_COLS:
            src = _find_col(subcols, [dest])
            if src is None:
                # Si los encabezados se duplican/abrevian, usar orden como último fallback.
                idx = PERSON_TERM_COLS.index(dest)
                src = subcols[idx]
            out[dest] = df[src].values
        return out

    return build(nat_cols), build(jur_cols)

def extract_all_tables(html_text: str):
    tables = discover_main_tables(html_text)
    missing = [key for key in KNOWN_TABLE_IDS if key not in tables]
    if missing:
        raise SchemaChangedError(f"Faltan tablas SBS: {missing}")

    parsed = {key: parse_main_table(tbl) for key, tbl in tables.items()}
    general_mn = harmonize_general(parsed[("MN", "general")])
    general_me = harmonize_general(parsed[("ME", "general")])
    mn_nat, mn_jur = split_person_table(parsed[("MN", "persona")])
    me_nat, me_jur = split_person_table(parsed[("ME", "persona")])
    return general_mn, general_me, mn_nat, mn_jur, me_nat, me_jur

def to_long_form(
    html_text: str,
    *,
    entity_type: str,
    period: str,
    period_date: str,
    frequency: str,
    source_url: str,
    retrieved_at: str,
) -> pd.DataFrame:
    """
    Convierte las 4 tablas SBS a un esquema canónico largo.

    No interpreta en exceso las columnas: `metric` conserva el nombre
    contractual de SBS (Ahorro, tramos, CTS, etc.).
    """
    (
        general_mn,
        general_me,
        mn_nat,
        mn_jur,
        me_nat,
        me_jur,
    ) = extract_all_tables(html_text)

    blocks = []

    def melt_block(
        frame: pd.DataFrame,
        *,
        currency: str,
        table_kind: str,
        person_type: str,
    ):
        if frame.empty:
            return

        value_cols = [c for c in frame.columns if c != "Banco"]

        long = frame.melt(
            id_vars=["Banco"],
            value_vars=value_cols,
            var_name="metric",
            value_name="rate",
        ).rename(columns={"Banco": "entity_name"})

        long["currency"] = currency
        long["table_kind"] = table_kind
        long["person_type"] = person_type
        blocks.append(long)

    melt_block(
        general_mn,
        currency="MN",
        table_kind="general",
        person_type="ALL",
    )
    melt_block(
        general_me,
        currency="ME",
        table_kind="general",
        person_type="ALL",
    )
    melt_block(
        mn_nat,
        currency="MN",
        table_kind="persona",
        person_type="Natural",
    )
    melt_block(
        mn_jur,
        currency="MN",
        table_kind="persona",
        person_type="Jurídica",
    )
    melt_block(
        me_nat,
        currency="ME",
        table_kind="persona",
        person_type="Natural",
    )
    melt_block(
        me_jur,
        currency="ME",
        table_kind="persona",
        person_type="Jurídica",
    )

    out = pd.concat(
        blocks,
        ignore_index=True,
        sort=False,
    )

    out.insert(0, "entity_type", entity_type)
    out.insert(1, "frequency", frequency)
    out.insert(2, "period", period)
    out.insert(3, "period_date", period_date)

    out["source"] = "SBS"
    out["source_url"] = source_url
    out["retrieved_at"] = retrieved_at

    out["rate"] = pd.to_numeric(
        out["rate"],
        errors="coerce",
    )

    cols = [
        "entity_type",
        "frequency",
        "period",
        "period_date",
        "currency",
        "table_kind",
        "person_type",
        "entity_name",
        "metric",
        "rate",
        "source",
        "source_url",
        "retrieved_at",
    ]

    return out[cols].reset_index(drop=True)
