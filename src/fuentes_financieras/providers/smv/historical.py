"""Generic EVCP discovery and yearly exports for every SMV mutual fund."""
from __future__ import annotations

from datetime import date
from hashlib import sha256
from http.cookiejar import CookieJar
from io import BytesIO
import json
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import HTTPCookieProcessor, Request, build_opener

from lxml import html as lxml_html
import pandas as pd

from fuentes_financieras.exceptions import SourceUnavailableError
from fuentes_financieras.storage import content_hash, schema_hash, utc_now_iso

URL = "https://www.smv.gob.pe/SIMV/Frm_EVCP?data=5A959494701B26421F184C081CACF55BFA328E8EBC"


def _fields(page: bytes) -> dict[str, str]:
    form = lxml_html.fromstring(page).xpath('//form[@id="frmMain"]')
    if not form:
        raise SourceUnavailableError("La SMV cambió el formulario EVCP.")
    return {node.get("name"): node.get("value", "") for node in
            form[0].xpath('.//input[@name]') if node.get("type") in ("hidden", "text")}


def _options(page: bytes, element_id: str) -> list[dict[str, str]]:
    options = lxml_html.fromstring(page).xpath(f'//select[@id="{element_id}"]/option')
    return [{"id": item.get("value", ""), "name": item.text_content().strip()}
            for item in options if item.get("value", "").isdigit()]


def _session(timeout: float):
    opener = build_opener(HTTPCookieProcessor(CookieJar()))

    def get() -> bytes:
        with opener.open(URL, timeout=timeout) as response:
            return response.read()

    def post(fields: dict[str, str]) -> bytes:
        request = Request(URL, data=urlencode(fields).encode(), headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "User-Agent": "Mozilla/5.0",
        })
        with opener.open(request, timeout=timeout) as response:
            return response.read()

    return get, post


def discover_funds(provider, *, refresh: bool = False, timeout: float = 90) -> list[dict[str, str]]:
    """Discover the company/fund IDs, caching the full universe independently of any report."""
    catalog = provider.storage.state_root / "smv_funds.json"
    if catalog.is_file() and not refresh:
        return json.loads(catalog.read_text(encoding="utf-8"))["funds"]
    get, post = _session(timeout)
    page = get()
    companies = _options(page, "MainContent_cboDenominacionSocial")
    if not companies:
        raise SourceUnavailableError("EVCP no devolvió administradoras de fondos.")
    funds: list[dict[str, str]] = []
    for company in companies:
        fields = _fields(page)
        fields.update({
            "__EVENTTARGET": "ctl00$MainContent$cboDenominacionSocial",
            "__EVENTARGUMENT": "",
            "ctl00$MainContent$cboDenominacionSocial": company["id"],
            "ctl00$MainContent$TextBox1": company["name"],
            "ctl00$MainContent$cboFondo": "--SELECCIONE FONDO--",
        })
        selected = post(fields)
        for fund in _options(selected, "MainContent_cboFondo"):
            funds.append({"company_id": company["id"], "company_name": company["name"],
                          "fund_id": fund["id"], "fund_name": fund["name"]})
    if not funds:
        raise SourceUnavailableError("EVCP no devolvió fondos para ninguna administradora.")
    catalog.parent.mkdir(parents=True, exist_ok=True)
    temporary = catalog.with_suffix(".json.tmp")
    temporary.write_text(json.dumps({"source_url": URL, "checked_at": utc_now_iso(),
                                     "funds": funds}, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(catalog)
    return funds


def fetch_export(fund: dict[str, str], start: pd.Timestamp, end: pd.Timestamp,
                 *, timeout: float = 180) -> bytes | None:
    """Export one company/fund interval using the site's native Excel action."""
    get, post = _session(timeout)
    page = get()
    fields = _fields(page)
    fields.update({
        "__EVENTTARGET": "ctl00$MainContent$cboDenominacionSocial",
        "__EVENTARGUMENT": "",
        "ctl00$MainContent$cboDenominacionSocial": fund["company_id"],
        "ctl00$MainContent$TextBox1": fund["company_name"],
        "ctl00$MainContent$cboFondo": "--SELECCIONE FONDO--",
    })
    page = post(fields)
    if fund["fund_id"] not in {x["id"] for x in _options(page, "MainContent_cboFondo")}:
        raise SourceUnavailableError(f"EVCP ya no ofrece {fund['fund_name']}.")
    fields = _fields(page)
    fields.update({
        "__EVENTTARGET": "", "__EVENTARGUMENT": "",
        "ctl00$MainContent$cboDenominacionSocial": fund["company_id"],
        "ctl00$MainContent$TextBox1": fund["company_name"],
        "ctl00$MainContent$cboFondo": fund["fund_id"],
        "ctl00$MainContent$txtFechDesde": start.strftime("%d/%m/%Y"),
        "ctl00$MainContent$txtFechHasta": end.strftime("%d/%m/%Y"),
        "ctl00$MainContent$btnBuscar": "Buscar",
    })
    page = post(fields)
    if not lxml_html.fromstring(page).xpath('//*[@id="MainContent_imgexcel"]'):
        return None
    fields = _fields(page)
    fields.update({
        "__EVENTTARGET": "ctl00$MainContent$imgexcel", "__EVENTARGUMENT": "",
        "ctl00$MainContent$cboDenominacionSocial": fund["company_id"],
        "ctl00$MainContent$TextBox1": fund["company_name"],
        "ctl00$MainContent$cboFondo": fund["fund_id"],
        "ctl00$MainContent$txtFechDesde": start.strftime("%d/%m/%Y"),
        "ctl00$MainContent$txtFechHasta": end.strftime("%d/%m/%Y"),
    })
    return post(fields)


def parse_export(raw: bytes, fund: dict[str, str], start: pd.Timestamp,
                 end: pd.Timestamp) -> pd.DataFrame:
    tables = pd.read_html(BytesIO(raw), decimal=".", thousands=",")
    if len(tables) != 1 or not {"Fecha Información", "Valor Cuota", "Fondo"}.issubset(tables[0].columns):
        raise SourceUnavailableError("EVCP: estructura de exportación inesperada.")
    original = tables[0]
    dates = pd.to_datetime(original["Fecha Información"].astype(str), format="%d/%m/%Y", errors="coerce")
    values = pd.to_numeric(original["Valor Cuota"], errors="coerce")
    if (dates.isna().any() or dates.duplicated().any() or values.isna().any()
            or values.lt(0).any() or dates.lt(start).any() or dates.gt(end).any()):
        raise SourceUnavailableError("EVCP: fechas duplicadas/fuera de rango o cuotas ilegibles.")
    return pd.DataFrame({
        "country_code": "PE",
        "query_date": dates,
        "value_date": dates,
        "fund_name": original["Fondo"].astype(str),
        "administrator": fund["company_name"],
        "unit_value": values,
        "source": "SMV",
        "transport": "EVCP_EXPORT",
        "source_url": URL,
        "source__company_id": fund["company_id"],
        "source__fund_id": fund["fund_id"],
        "source__fund_label": fund["fund_name"],
    }).sort_values("value_date").reset_index(drop=True)


def sync_historical(provider, *, desde: str | date, hasta: str | date,
                    refresh_catalog: bool = False, force: bool = False,
                    funds: list[dict[str, str]] | None = None,
                    timeout: float = 180) -> dict[str, int]:
    """Yearly incremental backfill; default scope is every discovered fund."""
    start, end = pd.Timestamp(desde).normalize(), pd.Timestamp(hasta).normalize()
    if end < start:
        raise ValueError("Rango inválido: hasta anterior a desde.")
    universe = funds if funds is not None else discover_funds(provider, refresh=refresh_catalog)
    result = {"funds": len(universe), "periods_reused": 0, "periods_new": 0,
              "periods_empty": 0, "rows_new": 0}
    for fund in universe:
        for year in range(start.year, end.year + 1):
            lo, hi = max(start, pd.Timestamp(year, 1, 1)), min(end, pd.Timestamp(year, 12, 31))
            key = f"evcp:{fund['company_id']}:{fund['fund_id']}:{year}"
            partition = f"evcp/year={year}/company={fund['company_id']}/fund={fund['fund_id']}"
            entry = provider.storage.manifest.get(key)
            same_range = (entry and entry.get("start") == lo.strftime("%Y-%m-%d")
                          and entry.get("end") == hi.strftime("%Y-%m-%d"))
            if not force and same_range and entry.get("status") == "unavailable":
                result["periods_reused"] += 1
                continue
            if (not force and entry and entry.get("status") == "validated"
                    and same_range
                    and provider.storage.period_matches(partition, key, entry.get("content_hash"))):
                result["periods_reused"] += 1
                continue
            raw = fetch_export(fund, lo, hi, timeout=timeout)
            if raw is None:
                provider.storage.manifest.set(key, {
                    "status": "unavailable", "period_key": key, "partition_key": partition,
                    "start": lo.strftime("%Y-%m-%d"), "end": hi.strftime("%Y-%m-%d"),
                    "checked_at": utc_now_iso(), "metadata": fund,
                })
                provider.storage.manifest.save()
                result["periods_empty"] += 1
                continue
            frame = parse_export(raw, fund, lo, hi)
            if frame.empty:
                result["periods_empty"] += 1
                continue
            path = provider.storage.upsert_period(partition, key, frame)
            raw_path = provider.storage.write_raw(key, raw, suffix=".xls")
            provider.storage.manifest.set(key, {
                "status": "validated", "period_key": key, "partition_key": partition,
                "parser_version": "evcp-1", "contract_version": provider.contract_version,
                "start": lo.strftime("%Y-%m-%d"), "end": hi.strftime("%Y-%m-%d"),
                "content_hash": content_hash(frame), "schema_hash": schema_hash(frame),
                "raw_sha256": sha256(raw).hexdigest(), "rows": len(frame),
                "checked_at": utc_now_iso(),
                "canonical_path": provider.storage.relative_path(path),
                "raw_path": provider.storage.relative_path(raw_path), "metadata": fund,
            })
            provider.storage.manifest.save()
            result["periods_new"] += 1
            result["rows_new"] += len(frame)
    return result
