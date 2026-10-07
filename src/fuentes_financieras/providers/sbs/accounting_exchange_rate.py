from __future__ import annotations

import io
import json
import re
from datetime import datetime, timezone
from html.parser import HTMLParser
from pathlib import Path
from typing import Iterable

import pandas as pd
import requests

SOURCE_URL = "https://www.sbs.gob.pe/app/stats/seriesH_TCC_res_excel.asp"
SOURCE_PAGE_URL = "https://www.sbs.gob.pe/app/stats/seriesH-TC-Contable.asp"
USD_CURRENCY_CODE = "02"
USD_CURRENCY_NAME = "Dólar de N.A."
DEFAULT_HISTORY_START = pd.Timestamp("2001-01-31")


class _TableRowsParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.rows: list[list[str]] = []
        self._row: list[str] | None = None
        self._cell: list[str] | None = None

    def handle_starttag(self, tag: str, attrs) -> None:
        name = tag.casefold()
        if name == "tr":
            self._row = []
        elif name in {"td", "th"} and self._row is not None:
            self._cell = []

    def handle_data(self, data: str) -> None:
        if self._cell is not None:
            self._cell.append(data)

    def handle_endtag(self, tag: str) -> None:
        name = tag.casefold()
        if name in {"td", "th"} and self._row is not None and self._cell is not None:
            text = re.sub(r"\s+", " ", " ".join(self._cell)).strip()
            self._row.append(text)
            self._cell = None
        elif name == "tr" and self._row is not None:
            if any(cell for cell in self._row):
                self.rows.append(self._row)
            self._row = None
            self._cell = None


def _as_timestamp(value: object) -> pd.Timestamp:
    ts = pd.Timestamp(value).normalize()
    if pd.isna(ts):
        raise ValueError(f"Fecha inválida: {value}")
    return ts


def _fmt_sbs_date(value: object) -> str:
    return _as_timestamp(value).strftime("%d/%m/%Y")


def _numeric_value(text: object) -> float | None:
    raw = str(text).strip().replace("\xa0", " ")
    if not raw:
        return None
    raw = re.sub(r"[^0-9,\.\-]", "", raw)
    if not raw or raw in {"-", ".", ","}:
        return None
    if "," in raw and "." in raw:
        # En la página SBS el punto es decimal; esta rama también tolera formatos
        # con separador de miles en caso cambie la salida.
        if raw.rfind(".") > raw.rfind(","):
            raw = raw.replace(",", "")
        else:
            raw = raw.replace(".", "").replace(",", ".")
    elif "," in raw:
        raw = raw.replace(",", ".")
    try:
        value = float(raw)
    except ValueError:
        return None
    return value if value > 0 else None


def _normalize_frame(frame: pd.DataFrame, *, currency_code: str = USD_CURRENCY_CODE) -> pd.DataFrame:
    if frame.empty:
        return pd.DataFrame(columns=[
            "date", "usd_pen_accounting", "currency_code", "currency", "source", "source_url", "retrieved_at"
        ])
    out = frame.copy()
    raw_dates = out["date"]
    if pd.api.types.is_datetime64_any_dtype(raw_dates):
        parsed_dates = pd.to_datetime(raw_dates, errors="coerce")
    else:
        parsed_dates = pd.to_datetime(raw_dates, errors="coerce", format="%Y-%m-%d")
        missing = parsed_dates.isna()
        if missing.any():
            parsed_dates.loc[missing] = pd.to_datetime(
                raw_dates.loc[missing], errors="coerce", format="%d/%m/%Y"
            )
    out["date"] = parsed_dates.dt.normalize()
    out["usd_pen_accounting"] = pd.to_numeric(out["usd_pen_accounting"], errors="coerce")
    out = out.dropna(subset=["date", "usd_pen_accounting"])
    out = out[out["usd_pen_accounting"] > 0]
    out["currency_code"] = str(currency_code)
    out["currency"] = "USD"
    out["source"] = "SBS_TIPO_CAMBIO_CONTABLE"
    out["source_url"] = SOURCE_URL
    if "retrieved_at" not in out.columns:
        out["retrieved_at"] = datetime.now(timezone.utc).isoformat()
    cols = ["date", "usd_pen_accounting", "currency_code", "currency", "source", "source_url", "retrieved_at"]
    out = out[cols].sort_values("date").drop_duplicates("date", keep="last").reset_index(drop=True)
    return out


def _parse_html_table(text: str, *, currency_code: str = USD_CURRENCY_CODE) -> pd.DataFrame:
    parser = _TableRowsParser()
    parser.feed(text)
    rows: list[dict] = []
    date_pattern = re.compile(r"\b(\d{1,2}/\d{1,2}/\d{4})\b")

    for cells in parser.rows:
        date_value: str | None = None
        date_idx: int | None = None
        for i, cell in enumerate(cells):
            m = date_pattern.search(cell)
            if m:
                date_value = m.group(1)
                date_idx = i
                break
        if date_value is None:
            continue

        candidates: list[float] = []
        for i, cell in enumerate(cells):
            if i == date_idx:
                continue
            # Ignora códigos enteros pequeños y toma preferentemente valores decimales de TC.
            for token in re.findall(r"-?\d+(?:[\.,]\d+)?", cell):
                val = _numeric_value(token)
                if val is not None and 0.1 <= val <= 100.0:
                    candidates.append(val)
        if not candidates:
            continue
        # El tipo de cambio suele ser el último valor numérico útil de la fila.
        rows.append({"date": date_value, "usd_pen_accounting": candidates[-1]})

    if not rows:
        # Fallback para una salida no tabular pero todavía HTML/texto.
        cleaned = re.sub(r"<[^>]+>", " ", text)
        for match in re.finditer(
            r"(\d{1,2}/\d{1,2}/\d{4}).{0,120}?([0-9]+(?:[\.,][0-9]{2,6}))",
            cleaned,
            flags=re.S,
        ):
            value = _numeric_value(match.group(2))
            if value is not None:
                rows.append({"date": match.group(1), "usd_pen_accounting": value})

    return _normalize_frame(pd.DataFrame(rows), currency_code=currency_code)


def parse_accounting_exchange_rate_response(
    content: bytes,
    *,
    content_type: str = "",
    encoding: str | None = None,
    currency_code: str = USD_CURRENCY_CODE,
) -> pd.DataFrame:
    """Convierte la salida del generador SBS a una serie canónica diaria.

    El recurso histórico de la SBS suele devolver HTML con apariencia de Excel.
    También se toleran XLSX/XLS reales si el entorno tiene el motor de lectura.
    """
    if not content:
        return _normalize_frame(pd.DataFrame(), currency_code=currency_code)

    # XLSX real.
    if content.startswith(b"PK"):
        try:
            raw = pd.read_excel(io.BytesIO(content), header=None)
            text = raw.to_html(index=False, header=False)
            return _parse_html_table(text, currency_code=currency_code)
        except Exception:
            pass

    # Real XLS/OLE input. pandas support depends on xlrd availability.
    # intentando como texto porque la SBS históricamente entrega HTML etiquetado como Excel.
    if content.startswith(b"\xd0\xcf\x11\xe0"):
        try:
            raw = pd.read_excel(io.BytesIO(content), header=None)
            text = raw.to_html(index=False, header=False)
            return _parse_html_table(text, currency_code=currency_code)
        except Exception:
            pass

    encodings = [encoding, "utf-8-sig", "utf-8", "cp1252", "latin-1"]
    text = None
    for enc in [x for x in encodings if x]:
        try:
            text = content.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = content.decode("latin-1", errors="replace")
    return _parse_html_table(text, currency_code=currency_code)


def fetch_accounting_exchange_rate(
    start_date: object,
    end_date: object,
    *,
    currency_code: str = USD_CURRENCY_CODE,
    timeout: int = 60,
    session: requests.Session | None = None,
) -> tuple[pd.DataFrame, bytes, str]:
    """Consulta directamente el generador histórico SBS, sin automatizar el navegador."""
    start = _as_timestamp(start_date)
    end = _as_timestamp(end_date)
    if start > end:
        raise ValueError("Rango inválido: start_date posterior a end_date.")

    params = {
        "fec_ini": _fmt_sbs_date(start),
        "fec_fin": _fmt_sbs_date(end),
        "chk": "",
        "moneda": str(currency_code),
    }
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; Financial-Sources/1.0)",
        "Referer": SOURCE_PAGE_URL,
        "Accept": "text/html,application/xhtml+xml,application/vnd.ms-excel,*/*;q=0.8",
    }
    client = session or requests.Session()

    def _request_once():
        response = client.get(SOURCE_URL, params=params, headers=headers, timeout=timeout)
        response.raise_for_status()
        content_type = response.headers.get("Content-Type", "")
        frame = parse_accounting_exchange_rate_response(
            response.content,
            content_type=content_type,
            encoding=response.encoding,
            currency_code=currency_code,
        )
        return response, frame, content_type

    response, frame, content_type = _request_once()
    if frame.empty:
        # Algunas aplicaciones ASP de la SBS inicializan cookies/estado al abrir primero
        # la página de la serie. No automatizamos clicks: solo arrancamos la sesión y
        # repetimos el GET directo al generador histórico.
        try:
            client.get(SOURCE_PAGE_URL, headers=headers, timeout=timeout).raise_for_status()
        except Exception:
            pass
        response, frame, content_type = _request_once()

    if frame.empty:
        raise RuntimeError(
            "La SBS respondió, pero no se pudo extraer ningún Tipo de Cambio Contable "
            f"entre {_fmt_sbs_date(start)} y {_fmt_sbs_date(end)}."
        )
    return frame, response.content, content_type


def _load_canonical(path: Path) -> pd.DataFrame:
    if not path.is_file():
        return _normalize_frame(pd.DataFrame())
    try:
        frame = pd.read_csv(path, encoding="utf-8-sig")
    except Exception:
        frame = pd.read_csv(path)
    return _normalize_frame(frame)


def _write_canonical(frame: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    out = _normalize_frame(frame)
    out.to_csv(path, index=False, encoding="utf-8-sig", date_format="%Y-%m-%d")


def sync_accounting_exchange_rate(
    storage_dir: str | Path,
    *,
    as_of: object | None = None,
    currency_code: str = USD_CURRENCY_CODE,
    initial_lookback_days: int = 45,
    overlap_days: int = 7,
    full_history: bool = False,
    force: bool = False,
    timeout: int = 60,
    verbose: bool = False,
) -> pd.DataFrame:
    """Sincroniza Tipo de Cambio Contable SBS en raw/ y canonical/.

    - Si no existe histórico local, el uso normal descarga solo una ventana reciente.
    - `full_history=True` permite hacer el backfill 31/01/2001 -> fecha objetivo una sola vez.
    - Si ya existe información, consulta únicamente un solapamiento corto desde la última fecha.
    - Evita repetir la misma consulta varias veces el mismo día/fecha objetivo.
    """
    base = Path(storage_dir)
    raw_dir = base / "raw"
    canonical_dir = base / "canonical"
    canonical_path = canonical_dir / "tipo_cambio_contable_usd_pen.csv"
    metadata_path = canonical_dir / "metadata.json"
    raw_dir.mkdir(parents=True, exist_ok=True)
    canonical_dir.mkdir(parents=True, exist_ok=True)

    target = _as_timestamp(as_of if as_of is not None else pd.Timestamp.today())
    existing = _load_canonical(canonical_path)

    metadata: dict = {}
    if metadata_path.is_file():
        try:
            metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
        except Exception:
            metadata = {}

    if not force and not full_history and not existing.empty:
        min_date = existing["date"].min()
        max_date = existing["date"].max()
        if min_date <= target <= max_date:
            return existing
        if metadata.get("last_checked_target") == target.strftime("%Y-%m-%d"):
            return existing

    if full_history:
        start = DEFAULT_HISTORY_START
    elif existing.empty:
        start = target - pd.Timedelta(days=max(1, int(initial_lookback_days)))
    else:
        min_date = existing["date"].min()
        max_date = existing["date"].max()
        if target < min_date:
            start = target - pd.Timedelta(days=max(1, int(initial_lookback_days)))
        else:
            start = max_date - pd.Timedelta(days=max(0, int(overlap_days)))

    if start > target:
        start = target

    if verbose:
        print(
            "SBS Tipo de Cambio Contable: "
            f"consultando {_fmt_sbs_date(start)} -> {_fmt_sbs_date(target)} (moneda={currency_code})."
        )

    try:
        fetched, raw_bytes, content_type = fetch_accounting_exchange_rate(
            start,
            target,
            currency_code=currency_code,
            timeout=timeout,
        )
        stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
        suffix = ".xlsx" if raw_bytes.startswith(b"PK") else ".xls" if raw_bytes.startswith(b"\xd0\xcf\x11\xe0") else ".html"
        raw_path = raw_dir / f"tcc_usd_{start:%Y%m%d}_{target:%Y%m%d}_{stamp}{suffix}"
        raw_path.write_bytes(raw_bytes)

        merged = pd.concat([existing, fetched], ignore_index=True, sort=False)
        merged = _normalize_frame(merged, currency_code=currency_code)
        _write_canonical(merged, canonical_path)
        metadata = {
            "source": "SBS Tipo de Cambio Contable",
            "source_url": SOURCE_URL,
            "currency_code": str(currency_code),
            "last_checked_target": target.strftime("%Y-%m-%d"),
            "last_checked_at_utc": datetime.now(timezone.utc).isoformat(),
            "last_content_type": content_type,
            "canonical_file": canonical_path.name,
            "first_date": merged["date"].min().strftime("%Y-%m-%d") if not merged.empty else None,
            "last_date": merged["date"].max().strftime("%Y-%m-%d") if not merged.empty else None,
        }
        metadata_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
        return merged
    except Exception:
        # Existing valid observations before the target date remain eligible for
        # seguir operando con el último dato almacenado. Si no existe nada, propagamos el error.
        usable = existing[existing["date"] <= target] if not existing.empty else existing
        if not usable.empty:
            if verbose:
                print("SBS no respondió; se utilizará el último Tipo de Cambio Contable almacenado.")
            return existing
        raise


def get_accounting_exchange_rate(
    as_of: object | None,
    *,
    storage_dir: str | Path,
    refresh: bool = True,
    force: bool = False,
    verbose: bool = False,
) -> dict:
    """Devuelve el último Tipo de Cambio Contable USD/PEN disponible hasta `as_of`."""
    target = _as_timestamp(as_of if as_of is not None else pd.Timestamp.today())
    base = Path(storage_dir)
    if refresh:
        frame = sync_accounting_exchange_rate(
            base,
            as_of=target,
            force=force,
            verbose=verbose,
        )
    else:
        frame = _load_canonical(base / "canonical" / "tipo_cambio_contable_usd_pen.csv")
    usable = frame[frame["date"] <= target].sort_values("date")
    if usable.empty:
        raise LookupError(f"No existe Tipo de Cambio Contable SBS disponible hasta {target:%Y-%m-%d}.")
    row = usable.iloc[-1]
    return {
        "date": pd.Timestamp(row["date"]),
        "usd_pen_accounting": float(row["usd_pen_accounting"]),
        "currency_code": str(row["currency_code"]),
        "source": str(row["source"]),
    }


# Alias en español para uso interactivo.
sincronizar_tipo_cambio_contable = sync_accounting_exchange_rate
obtener_tipo_cambio_contable = get_accounting_exchange_rate
