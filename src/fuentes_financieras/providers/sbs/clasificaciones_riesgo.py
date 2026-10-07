from __future__ import annotations

import re
from calendar import monthrange
from datetime import date, datetime
from typing import Iterable

import pandas as pd

from fuentes_financieras.exceptions import (
    InvalidQueryError,
    PeriodUnavailableError,
    SourceUnavailableError,
)
from fuentes_financieras.models import FetchResult, PeriodRequest
from fuentes_financieras.provider import DatasetProvider
from fuentes_financieras.storage import utc_now_iso
from fuentes_financieras.transports.curl_chrome import CurlChromeTransport

from ._clasificaciones_parser import (
    ENTITY_FIELD,
    PERIOD_FIELD,
    available_periods,
    entity_type_map,
    period_parts,
    selected_period_code,
    to_long_form,
)
from ._webforms import successful_controls


URL = (
    "https://www.sbs.gob.pe/app/iece/paginas/"
    "MostrarResumenClasificaciones.aspx"
)

BUTTON_FIELD = "ctl00$MainContent$BtnConsultar"
SCRIPT_MANAGER_FIELD = "ctl00$ScriptManager"
BUTTON_VALUE = "Consultar"
SCRIPT_MANAGER_VALUE = "ctl00$ScriptManager|ctl00$MainContent$BtnConsultar"


_HIDDEN_RE = re.compile(
    r"\|(\d+)\|hiddenField\|([^|]+)\|([^|]*)"
)


def _parse_date(value) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        text = value.strip()
        for fmt in ("%Y-%m-%d", "%Y-%m", "%Y%m", "%Y%m%d"):
            try:
                parsed = datetime.strptime(text, fmt)
                return parsed.date()
            except ValueError:
                pass
    raise InvalidQueryError(
        f"Fecha inválida: {value!r}. Use YYYY-MM-DD o YYYY-MM."
    )




def _parse_bound(value, *, end: bool = False) -> date:
    if isinstance(value, str) and re.fullmatch(r"\d{4}-\d{2}", value.strip()):
        year, month = map(int, value.strip().split("-"))
        day = monthrange(year, month)[1] if end else 1
        return date(year, month, day)
    return _parse_date(value)

def _period_from_year_semester(year: int, semester: int) -> str:
    year = int(year)
    semester = int(semester)
    if semester not in {1, 2}:
        raise InvalidQueryError("semestre debe ser 1 o 2.")
    return f"{year:04d}0{semester}"


def _period_code_from_value(value) -> str:
    code = str(value).strip()
    try:
        period_parts(code)
    except ValueError as exc:
        raise InvalidQueryError(str(exc)) from exc
    return code


def _parse_delta_response(text: str) -> tuple[str, dict[str, str]]:
    marker = "|updatePanel|"
    pos = text.find(marker)
    if pos < 0:
        raise SourceUnavailableError(
            "La respuesta SBS no contiene updatePanel ASP.NET."
        )

    panel_name_end = text.find("|", pos + len(marker))
    if panel_name_end < 0:
        raise SourceUnavailableError("Respuesta ASP.NET delta malformada.")

    content_start = panel_name_end + 1

    boundary = re.search(
        r"\|\d+\|hiddenField\|__EVENTTARGET\|",
        text[content_start:],
    )
    if boundary is None:
        boundary = re.search(
            r"\|\d+\|hiddenField\|__VIEWSTATE\|",
            text[content_start:],
        )
    if boundary is None:
        raise SourceUnavailableError(
            "No se encontró el límite del updatePanel ASP.NET."
        )

    content_end = content_start + boundary.start()
    fragment = text[content_start:content_end]

    hidden: dict[str, str] = {}
    for match in _HIDDEN_RE.finditer(text[content_end:]):
        hidden[match.group(2)] = match.group(3)

    return fragment, hidden


def _merge_hidden_state(
    previous: dict[str, str],
    payload: dict[str, str],
    hidden: dict[str, str],
) -> dict[str, str]:
    out = dict(previous)
    out[PERIOD_FIELD] = payload.get(PERIOD_FIELD, "")
    out[ENTITY_FIELD] = payload.get(ENTITY_FIELD, "")
    out.update(hidden)
    return out


class RiskRatingsClient:
    """Stateful client for the SBS WebForms ratings page."""

    def __init__(self, *, state_dir):
        self.transport = CurlChromeTransport(
            state_dir=state_dir,
            normal_marker="Clasificaciones e Informes Semestrales",
        )
        self.form_state: dict[str, str] = {}
        self.initial_html = ""
        self.current_html = ""
        self.current_period: str | None = None
        self.period_catalog: list[dict] = []
        self.type_code_by_label: dict[str, str] = {}
        self.opened = False

    def reset(self):
        self.transport.reset()
        self.form_state = {}
        self.initial_html = ""
        self.current_html = ""
        self.current_period = None
        self.period_catalog = []
        self.type_code_by_label = {}
        self.opened = False

    def open(self):
        if self.opened:
            return

        response = self.transport.request("GET", URL)
        html = response.text

        if "Clasificaciones e Informes Semestrales" not in html:
            raise SourceUnavailableError(
                "GET SBS no devolvió la página de clasificaciones esperada."
            )

        self.form_state = successful_controls(html)
        self.initial_html = html
        self.current_html = html
        self.current_period = selected_period_code(html)
        self.period_catalog = available_periods(html)
        self.type_code_by_label = entity_type_map(html)
        self.opened = True

    def list_periods(self) -> list[dict]:
        self.open()
        return [dict(x) for x in self.period_catalog]

    def _payload(self, period_code: str) -> dict[str, str]:
        data = dict(self.form_state)
        data[PERIOD_FIELD] = period_code
        data[ENTITY_FIELD] = ""
        data["__EVENTTARGET"] = ""
        data["__EVENTARGUMENT"] = ""
        data[SCRIPT_MANAGER_FIELD] = SCRIPT_MANAGER_VALUE
        data["__ASYNCPOST"] = "true"
        data[BUTTON_FIELD] = BUTTON_VALUE
        return data

    def fetch_period(self, period_code: str) -> tuple[str, str]:
        """
        Return (html_fragment, raw_response).

        The period already present in the initial GET requires no POST.
        Any other period requires one POST and updates the WebForms state.
        """
        self.open()
        period_code = _period_code_from_value(period_code)

        available = {x["period_code"] for x in self.period_catalog}
        if period_code not in available:
            raise PeriodUnavailableError(
                f"Período {period_code} no publicado por SBS."
            )

        if self.current_period == period_code and self.current_html:
            return self.current_html, self.current_html

        payload = self._payload(period_code)

        response = self.transport.request(
            "POST",
            URL,
            data=payload,
            headers={
                "X-MicrosoftAjax": "Delta=true",
                "Origin": "https://www.sbs.gob.pe",
                "Referer": URL,
                "Accept": "*/*",
                "Content-Type": (
                    "application/x-www-form-urlencoded; charset=UTF-8"
                ),
            },
        )

        fragment, hidden = _parse_delta_response(response.text)

        required = {
            "__EVENTTARGET",
            "__EVENTARGUMENT",
            "__VIEWSTATE",
            "__VIEWSTATEGENERATOR",
            "__PREVIOUSPAGE",
            "__EVENTVALIDATION",
        }
        missing = sorted(required - set(hidden))
        if missing:
            raise SourceUnavailableError(
                "Respuesta SBS incompleta; hidden fields faltantes: "
                + ", ".join(missing)
            )

        self.form_state = _merge_hidden_state(
            self.form_state,
            payload,
            hidden,
        )
        self.current_html = fragment
        self.current_period = period_code

        return fragment, response.text


class RiskRatingsProvider(DatasetProvider):
    parser_version = "2026-10-06.2"
    contract_version = "2"

    def __init__(self, spec):
        super().__init__(spec)
        self.client = RiskRatingsClient(
            state_dir=self.storage.state_root,
        )

    def available_periods(self) -> list[dict]:
        return self.client.list_periods()

    def single_request(
        self,
        *,
        periodo=None,
        anio=None,
        semestre=None,
        **_,
    ) -> PeriodRequest:
        if periodo is None:
            if anio is None or semestre is None:
                raise InvalidQueryError(
                    "Use periodo='YYYY01/YYYY02' o anio=YYYY, semestre=1/2."
                )
            periodo = _period_from_year_semester(anio, semestre)

        code = _period_code_from_value(periodo)
        year, sem, month, period = period_parts(code)

        catalog = self.available_periods()
        available = {x["period_code"] for x in catalog}
        if code not in available:
            raise PeriodUnavailableError(
                f"Período {code} no publicado por SBS."
            )

        latest = catalog[0]["period_code"] if catalog else None

        return PeriodRequest(
            period_key=code,
            partition_key=f"year={year:04d}",
            params={
                "periodo": code,
                "year": year,
                "semester": sem,
            },
            mutable=(code == latest),
        )

    def plan_sync(
        self,
        *,
        desde=None,
        hasta=None,
        periodos=None,
        **_,
    ) -> Iterable[PeriodRequest]:
        catalog = self.available_periods()
        if not catalog:
            return

        by_code = {x["period_code"]: x for x in catalog}

        if periodos is not None:
            if isinstance(periodos, str):
                periodos = [periodos]
            wanted = [_period_code_from_value(x) for x in periodos]
            missing = [x for x in wanted if x not in by_code]
            if missing:
                raise PeriodUnavailableError(
                    f"Períodos no publicados por SBS: {missing}"
                )
            selected = [by_code[x] for x in wanted]
        else:
            selected = list(catalog)

            if desde is not None:
                start = _parse_bound(desde, end=False)
                selected = [
                    x for x in selected
                    if pd.Timestamp(x["period_date"]) >= pd.Timestamp(start)
                ]

            if hasta is not None:
                end = _parse_bound(hasta, end=True)
                selected = [
                    x for x in selected
                    if pd.Timestamp(x["period_date"]) <= pd.Timestamp(end)
                ]

        # Keep the GET-selected period first so it can be parsed with zero POSTs;
        # then process the remaining periods in SBS publication order.
        current = self.client.current_period
        selected.sort(
            key=lambda x: (
                0 if x["period_code"] == current else 1,
                -int(x["period_code"]),
            )
        )

        latest = catalog[0]["period_code"]

        for item in selected:
            yield PeriodRequest(
                period_key=item["period_code"],
                partition_key=f"year={item['year']:04d}",
                params={
                    "periodo": item["period_code"],
                    "year": item["year"],
                    "semester": item["semester"],
                },
                mutable=(item["period_code"] == latest),
            )

    def _fetch_period(self, request: PeriodRequest) -> FetchResult:
        code = request.params["periodo"]

        try:
            html, raw = self.client.fetch_period(code)
        except Exception:
            # Discard an ambiguous WebForms state and retry once from a clean GET.
            self.client.reset()
            try:
                html, raw = self.client.fetch_period(code)
            except Exception:
                self.client.reset()
                raise

        retrieved_at = utc_now_iso()

        data = to_long_form(
            html,
            period_code=code,
            type_code_by_label=self.client.type_code_by_label,
            source_url=URL,
            retrieved_at=retrieved_at,
        )

        if data.empty:
            raise PeriodUnavailableError(
                f"SBS devolvió período {code} sin clasificaciones."
            )

        unknown_types = sorted(
            set(data.loc[data["entity_type_code"].isna(), "entity_type"])
        )

        return FetchResult(
            dataset_id=self.spec.dataset_id,
            data=data,
            raw=raw,
            metadata={
                "period_code": code,
                "rows": int(len(data)),
                "entities": int(
                    data[["entity_type", "entity_name"]]
                    .drop_duplicates()
                    .shape[0]
                ),
                "entity_types": sorted(
                    data["entity_type_code"].dropna().unique().tolist()
                ),
                "unmapped_entity_type_labels": unknown_types,
                "url": URL,
                "transport": "curl_cffi/chrome + ASP.NET WebForms",
                "parser_version": self.parser_version,
            },
        )

    def filter_loaded(
        self,
        data: pd.DataFrame,
        *,
        desde=None,
        hasta=None,
        periodo=None,
        periodos=None,
        tipo_entidad=None,
        clasificadora=None,
        entidad=None,
        **_,
    ) -> pd.DataFrame:
        out = data.copy()

        if periodo is not None:
            periodos = [periodo]

        if periodos is not None:
            if isinstance(periodos, str):
                periodos = [periodos]
            wanted = {_period_code_from_value(x) for x in periodos}
            out = out[out["period_code"].astype(str).isin(wanted)]

        if "period_date" in out.columns:
            dates = pd.to_datetime(out["period_date"], errors="coerce")
            if desde is not None:
                out = out[dates >= pd.Timestamp(_parse_bound(desde, end=False))]
                dates = pd.to_datetime(out["period_date"], errors="coerce")
            if hasta is not None:
                out = out[dates <= pd.Timestamp(_parse_bound(hasta, end=True))]

        if tipo_entidad is not None:
            wanted = str(tipo_entidad).strip()
            mask = (
                out["entity_type_code"].astype(str).eq(wanted)
                | out["entity_type"].astype(str).str.casefold().eq(wanted.casefold())
            )
            out = out[mask]

        if clasificadora is not None:
            wanted = str(clasificadora).strip().casefold()
            out = out[
                out["rating_agency"].astype(str).str.casefold().str.contains(
                    re.escape(wanted), regex=True, na=False
                )
            ]

        if entidad is not None:
            wanted = str(entidad).strip().casefold()
            out = out[
                out["entity_name"].astype(str).str.casefold().str.contains(
                    re.escape(wanted), regex=True, na=False
                )
            ]

        return out.drop(columns=["_period_key"], errors="ignore")
