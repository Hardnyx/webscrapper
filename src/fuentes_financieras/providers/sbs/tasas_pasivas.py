from __future__ import annotations

import json
import re
from calendar import monthrange
from datetime import date, datetime, timedelta, timezone
from typing import Iterable

import pandas as pd
from bs4 import BeautifulSoup, FeatureNotFound

from fuentes_financieras.exceptions import (
    InvalidQueryError,
    PeriodUnavailableError,
    SchemaChangedError,
    SourceUnavailableError,
)
from fuentes_financieras.models import FetchResult, PeriodRequest
from fuentes_financieras.provider import DatasetProvider
from fuentes_financieras.storage import utc_now_iso
from fuentes_financieras.transports.curl_chrome import CurlChromeTransport

from ._tasas_parser import KNOWN_TABLE_IDS, to_long_form
from ._webforms import (
    merge_hidden_state,
    parse_delta_response,
    set_suffix,
    successful_controls,
)


DAILY_TYPES = {"B", "F"}
MONTHLY_TYPES = {"C", "R"}
ALL_TYPES = DAILY_TYPES | MONTHLY_TYPES

MONTH_NAMES = [
    "Enero",
    "Febrero",
    "Marzo",
    "Abril",
    "Mayo",
    "Junio",
    "Julio",
    "Agosto",
    "Setiembre",
    "Octubre",
    "Noviembre",
    "Diciembre",
]

MONTH_TO_NUM = {
    name.lower(): i + 1
    for i, name in enumerate(MONTH_NAMES)
}

BASE_URL = (
    "https://www.sbs.gob.pe/app/pp/EstadisticasSAEEPortal/"
    "Paginas/TIPasivaDepositoEmpresa.aspx?tip={tipo}"
)

def _make_soup(html: str):
    try:
        return BeautifulSoup(html, "lxml")
    except FeatureNotFound:
        return BeautifulSoup(html, "html.parser")



def _parse_date(value) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                pass
    raise InvalidQueryError(
        f"Fecha inválida: {value!r}. Formatos admitidos: YYYY-MM-DD o DD/MM/YYYY."
    )


def _month_start(d: date) -> date:
    return d.replace(day=1)


def _add_months(d: date, months: int) -> date:
    index = d.year * 12 + (d.month - 1) + months
    year, month0 = divmod(index, 12)
    return date(year, month0 + 1, 1)


def _iter_months(start: date, end: date):
    cur = _month_start(start)
    stop = _month_start(end)
    while cur <= stop:
        yield cur
        cur = _add_months(cur, 1)


def _effective_daily_date(html: str) -> str | None:
    m = re.search(
        r'id="ctl00_cphContent_lblMensajeFecha"[^>]*>'
        r'.*?al\s+(\d{2}/\d{2}/\d{4})',
        html,
        flags=re.I | re.S,
    )
    if m:
        return m.group(1)

    m = re.search(
        r'id="ctl00_cphContent_rdpDate_dateInput"[^>]*'
        r'value="(\d{2}/\d{2}/\d{4})"',
        html,
        flags=re.I | re.S,
    )
    return m.group(1) if m else None


def _effective_month(html: str) -> tuple[int, int] | None:
    m = re.search(
        r'\ba\s+(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|'
        r'Setiembre|Octubre|Noviembre|Diciembre)\s+del\s+(\d{4})',
        html,
        flags=re.I | re.S,
    )
    if not m:
        return None

    month = MONTH_TO_NUM[m.group(1).lower()]
    year = int(m.group(2))
    return year, month


def _has_four_tables(html: str) -> bool:
    return all(
        table_id in html
        for table_id in KNOWN_TABLE_IDS.values()
    )


def _validate_published_tables(html: str):
    if _has_four_tables(html):
        return
    text = _make_soup(html).get_text(' ', strip=True)
    if re.search(r'No\s+existe\s+informaci[oó]n\s+para\s+la\s+fecha\s+elegida', text, re.I):
        raise PeriodUnavailableError('SBS: No existe información para la fecha elegida.')
    raise SchemaChangedError('Respuesta sin las cuatro tablas SBS.')


def _dropdown_state(
    *,
    selected_index: int,
    selected_text: str,
    selected_value: str,
) -> str:
    return json.dumps(
        {
            "enabled": True,
            "logEntries": [],
            "selectedIndex": selected_index,
            "selectedText": selected_text,
            "selectedValue": selected_value,
        },
        separators=(",", ":"),
        ensure_ascii=False,
    )


class PassiveRatesClient:
    """
    Cliente de una sola `tip`.

    Mantiene el estado WebForms entre consultas para evitar GET innecesarios.
    """

    def __init__(
        self,
        *,
        tipo: str,
        state_dir,
    ):
        tipo = tipo.upper().strip()

        if tipo not in ALL_TYPES:
            raise InvalidQueryError(
                f"Tipo inválido: {tipo!r}. Valores admitidos: B/F/C/R."
            )

        self.tipo = tipo
        self.url = BASE_URL.format(tipo=tipo)

        marker = (
            "ctl00_cphContent_btnConsultar"
            if tipo in DAILY_TYPES
            else "ctl00_cphContent_btnConsultaMensual"
        )

        self.transport = CurlChromeTransport(
            state_dir=state_dir,
            normal_marker=marker,
        )

        self.form_state: dict[str, str] = {}
        self.last_html = ""
        self.opened = False

        self.selected_year: int | None = None
        self.selected_month: int | None = None
        self.year_options: list[str] = []

    def reset(self):
        self.transport.reset()
        self.form_state = {}
        self.last_html = ""
        self.opened = False
        self.selected_year = None
        self.selected_month = None
        self.year_options = []

    def open(self):
        response = self.transport.request(
            "GET",
            self.url,
        )

        html = response.text

        if "__VIEWSTATE" not in html:
            raise SourceUnavailableError(
                "GET SBS sin __VIEWSTATE."
            )

        self.form_state = successful_controls(html)
        self.last_html = html
        self.opened = True

        if self.tipo in MONTHLY_TYPES:
            self._read_monthly_selection(html)

        return self

    def _read_monthly_selection(self, html: str):
        soup = _make_soup(html)

        year_root = soup.find(
            id="ctl00_cphContent_rAnio"
        )
        month_root = soup.find(
            id="ctl00_cphContent_rMes"
        )

        if year_root:
            fake = year_root.find(
                class_="rddlFakeInput"
            )
            if fake:
                txt = fake.get_text(" ", strip=True)
                if txt.isdigit():
                    self.selected_year = int(txt)

        if month_root:
            fake = month_root.find(
                class_="rddlFakeInput"
            )
            if fake:
                txt = fake.get_text(" ", strip=True)
                self.selected_month = MONTH_TO_NUM.get(
                    txt.lower()
                )

        dropdown = soup.find(
            id="ctl00_cphContent_rAnio_DropDown"
        )
        if dropdown:
            self.year_options = [
                li.get_text(" ", strip=True)
                for li in dropdown.select("li.rddlItem")
                if li.get_text(" ", strip=True)
            ]

    def _daily_payload(self, target: date):
        data = dict(self.form_state)

        iso = target.strftime("%Y-%m-%d")
        display = target.strftime("%d/%m/%Y")
        iso_state = f"{iso}-00-00-00"

        set_suffix(
            data,
            "$rdpDate",
            iso,
            "ctl00$cphContent$rdpDate",
        )
        set_suffix(
            data,
            "$rdpDate$dateInput",
            display,
            "ctl00$cphContent$rdpDate$dateInput",
        )

        input_state = {
            "enabled": True,
            "emptyMessage": "",
            "validationText": iso_state,
            "valueAsString": iso_state,
            "minDateStr": "1000-01-01-00-00-00",
            "maxDateStr": "2099-12-31-00-00-00",
            "lastSetTextBoxValue": display,
        }

        set_suffix(
            data,
            "rdpDate_dateInput_ClientState",
            json.dumps(
                input_state,
                separators=(",", ":"),
            ),
            "ctl00_cphContent_rdpDate_dateInput_ClientState",
        )

        set_suffix(
            data,
            "$hdTipoEntidad",
            self.tipo,
            "ctl00$cphContent$hdTipoEntidad",
        )

        data["__EVENTTARGET"] = ""
        data["__EVENTARGUMENT"] = ""

        data["ctl00$MainScriptManager"] = (
            "ctl00$cphContent$updConsulta|"
            "ctl00$cphContent$btnConsultar"
        )

        data["__ASYNCPOST"] = "true"
        data[
            "ctl00$cphContent$btnConsultar"
        ] = "Consultar"

        # El diagnóstico demostró que el TSM generado por JS no es necesario.
        data.pop(
            "ctl00_MainScriptManager_TSM",
            None,
        )

        return data

    def _monthly_payload(
        self,
        *,
        year: int,
        month: int,
    ):
        data = dict(self.form_state)

        if not (1 <= month <= 12):
            raise InvalidQueryError(
                f"Mes inválido: {month}"
            )

        if self.year_options:
            if str(year) not in self.year_options:
                raise PeriodUnavailableError(
                    f"SBS no ofrece el año {year} para tip={self.tipo}."
                )
            year_index = self.year_options.index(str(year))
        else:
            # Fallback actual: lista descendente.
            if self.selected_year is None:
                raise SchemaChangedError(
                    "No se pudo identificar lista de años Telerik."
                )
            year_index = self.selected_year - year

        month_index = month - 1
        month_name = MONTH_NAMES[month_index]

        year_changed = (
            self.selected_year is None
            or year != self.selected_year
        )
        month_changed = (
            self.selected_month is None
            or month != self.selected_month
        )

        data[
            "ctl00_cphContent_rAnio_ClientState"
        ] = (
            _dropdown_state(
                selected_index=year_index,
                selected_text=str(year),
                selected_value=str(year),
            )
            if year_changed
            else ""
        )

        data[
            "ctl00_cphContent_rMes_ClientState"
        ] = (
            _dropdown_state(
                selected_index=month_index,
                selected_text=month_name,
                selected_value=f"{month:02d}",
            )
            if month_changed
            else ""
        )

        set_suffix(
            data,
            "$hdTipoEntidad",
            self.tipo,
            "ctl00$cphContent$hdTipoEntidad",
        )

        data["__EVENTTARGET"] = ""
        data["__EVENTARGUMENT"] = ""

        data["ctl00$MainScriptManager"] = (
            "ctl00$cphContent$updConsulta|"
            "ctl00$cphContent$btnConsultaMensual"
        )

        data["__ASYNCPOST"] = "true"
        data[
            "ctl00$cphContent$btnConsultaMensual"
        ] = "Consultar"

        data.pop(
            "ctl00_MainScriptManager_TSM",
            None,
        )

        return data

    def _post(self, payload: dict[str, str]):
        response = self.transport.request(
            "POST",
            self.url,
            data=payload,
        )

        raw = response.text

        try:
            html, hidden = parse_delta_response(raw)
        except ValueError:
            html, hidden = raw, {}

        self.form_state = merge_hidden_state(
            payload,
            html,
            hidden,
        )
        self.last_html = html

        return html, raw

    def fetch_daily(self, target: date):
        if self.tipo not in DAILY_TYPES:
            raise InvalidQueryError(
                f"tip={self.tipo} no es diario."
            )

        if not self.opened:
            self.open()

        payload = self._daily_payload(target)
        html, raw = self._post(payload)

        expected = target.strftime("%d/%m/%Y")
        effective = _effective_daily_date(html)

        if effective != expected:
            raise PeriodUnavailableError(
                f"Solicitado {expected}; SBS devolvió {effective!r}."
            )

        _validate_published_tables(html)

        return html, raw

    def fetch_monthly(
        self,
        *,
        year: int,
        month: int,
    ):
        if self.tipo not in MONTHLY_TYPES:
            raise InvalidQueryError(
                f"tip={self.tipo} no es mensual."
            )

        if not self.opened:
            self.open()

        payload = self._monthly_payload(
            year=year,
            month=month,
        )

        html, raw = self._post(payload)

        effective = _effective_month(html)

        if effective != (year, month):
            raise PeriodUnavailableError(
                f"Solicitado {year}-{month:02d}; SBS devolvió {effective!r}."
            )

        _validate_published_tables(html)

        self.selected_year = year
        self.selected_month = month

        return html, raw


class PassiveRatesProvider(DatasetProvider):
    parser_version = "2026-10-07.1"
    contract_version = "1"

    def __init__(self, spec):
        super().__init__(spec)
        self._clients: dict[str, PassiveRatesClient] = {}

    def _client(self, tipo: str):
        tipo = tipo.upper()
        if tipo not in self._clients:
            self._clients[tipo] = PassiveRatesClient(
                tipo=tipo,
                state_dir=self.storage.state_root,
            )
        return self._clients[tipo]

    def single_request(
        self,
        *,
        tipo: str,
        fecha=None,
        anio: int | None = None,
        mes: int | None = None,
        **_,
    ) -> PeriodRequest:
        tipo = tipo.upper().strip()

        if tipo not in ALL_TYPES:
            raise InvalidQueryError(
                "Tipo fuera de dominio: B, F, C o R."
            )

        today = date.today()

        if tipo in DAILY_TYPES:
            if fecha is None:
                raise InvalidQueryError(
                    f"tip={tipo} requiere fecha."
                )

            d = _parse_date(fecha)

            return PeriodRequest(
                period_key=f"{tipo}:{d.isoformat()}",
                partition_key=f"entity_type={tipo}/year={d.year}",
                params={
                    "tipo": tipo,
                    "fecha": d.isoformat(),
                },
                mutable=d >= today - timedelta(days=7),
            )

        if anio is None or mes is None:
            raise InvalidQueryError(
                f"tip={tipo} requiere anio y mes."
            )

        d = date(int(anio), int(mes), 1)
        mutable_floor = _add_months(
            _month_start(today),
            -2,
        )

        return PeriodRequest(
            period_key=f"{tipo}:{d:%Y-%m}",
            partition_key=f"entity_type={tipo}/year={d.year}",
            params={
                "tipo": tipo,
                "anio": d.year,
                "mes": d.month,
            },
            mutable=d >= mutable_floor,
        )

    def plan_sync(
        self,
        *,
        tipos,
        desde,
        hasta=None,
        **_,
    ):
        if isinstance(tipos, str):
            tipos = [tipos]

        tipos = [
            str(t).upper().strip()
            for t in tipos
        ]

        invalid = sorted(set(tipos) - ALL_TYPES)
        if invalid:
            raise InvalidQueryError(
                f"Tipos inválidos: {invalid}"
            )

        start = _parse_date(desde)
        end = _parse_date(hasta or date.today())

        if end < start:
            raise InvalidQueryError(
                "Rango inválido: hasta anterior a desde."
            )

        for tipo in tipos:
            if tipo in DAILY_TYPES:
                cur = start
                while cur <= end:
                    # Evitar llamadas obvias sin publicación.
                    if cur.weekday() < 5:
                        yield self.single_request(
                            tipo=tipo,
                            fecha=cur,
                        )
                    cur += timedelta(days=1)

            else:
                for month_date in _iter_months(
                    start,
                    end,
                ):
                    yield self.single_request(
                        tipo=tipo,
                        anio=month_date.year,
                        mes=month_date.month,
                    )

    def _fetch_period(
        self,
        request: PeriodRequest,
    ) -> FetchResult:
        tipo = request.params["tipo"]
        client = self._client(tipo)

        try:
            if tipo in DAILY_TYPES:
                target = _parse_date(
                    request.params["fecha"]
                )
                html, raw = client.fetch_daily(target)

                period = target.isoformat()
                period_date = target.isoformat()
                frequency = "daily"

            else:
                year = int(request.params["anio"])
                month = int(request.params["mes"])

                html, raw = client.fetch_monthly(
                    year=year,
                    month=month,
                )

                period = f"{year:04d}-{month:02d}"
                period_date = f"{year:04d}-{month:02d}-01"
                frequency = "monthly"

        except Exception:
            # Unexpected responses may leave ambiguous ASP.NET state.
            client.reset()
            raise

        retrieved_at = utc_now_iso()

        data = to_long_form(
            html,
            entity_type=tipo,
            period=period,
            period_date=period_date,
            frequency=frequency,
            source_url=client.url,
            retrieved_at=retrieved_at,
        )

        return FetchResult(
            dataset_id=self.spec.dataset_id,
            data=data,
            raw=raw,
            metadata={
                "tipo": tipo,
                "frequency": frequency,
                "period": period,
                "rows": int(len(data)),
                "url": client.url,
                "transport": "curl_cffi/chrome+webforms",
                "parser_version": self.parser_version,
            },
        )

    def filter_loaded(
        self,
        data: pd.DataFrame,
        *,
        tipos=None,
        desde=None,
        hasta=None,
        **_,
    ) -> pd.DataFrame:
        out = data.copy()

        if tipos is not None:
            if isinstance(tipos, str):
                tipos = [tipos]
            wanted = {
                str(x).upper().strip()
                for x in tipos
            }
            out = out[
                out["entity_type"].isin(wanted)
            ]

        if "period_date" in out.columns:
            dates = pd.to_datetime(
                out["period_date"],
                errors="coerce",
            )

            if desde is not None:
                out = out[
                    dates >= pd.Timestamp(
                        _parse_date(desde)
                    )
                ]
                dates = pd.to_datetime(
                    out["period_date"],
                    errors="coerce",
                )

            if hasta is not None:
                out = out[
                    dates <= pd.Timestamp(
                        _parse_date(hasta)
                    )
                ]

        return out.drop(
            columns=["_period_key"],
            errors="ignore",
        )
