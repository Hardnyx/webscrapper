"""Published SBS stock and flow passive market references."""
from datetime import date, timedelta
import json
import re

import pandas as pd
from bs4 import BeautifulSoup

from fuentes_financieras.exceptions import InvalidQueryError, PeriodUnavailableError, SchemaChangedError
from fuentes_financieras.models import FetchResult, PeriodRequest
from fuentes_financieras.provider import DatasetProvider
from fuentes_financieras.storage import utc_now_iso
from fuentes_financieras.transports.curl_chrome import CurlChromeTransport
from ._webforms import successful_controls, parse_delta_response, merge_hidden_state
from .tasas_pasivas import _parse_date

URL = 'https://www.sbs.gob.pe/app/pp/EstadisticasSAEEPortal/Paginas/TIPasivaMercado.aspx?tip=B'
METHODOLOGY_URL = 'https://www.sbs.gob.pe/app/stats/metodologia/metodologia_ti_promedio.pdf'
SERIES = {
    'TIPMN': ('lblVAL_TIPMN_TASA', 'MN', 'stock', 'outstanding_balances', 'B+F', 'lblFecha'),
    'TIPMEX': ('lblVAL_TIPMEX_TASA', 'ME', 'stock', 'outstanding_balances', 'B+F', 'lblFecha'),
    'FTIPMN': ('lblVAL_FTIPMN', 'MN', 'flow', 'last_30_business_days', 'B', 'lblFecha2'),
    'FTIPMEX': ('lblVAL_FTIPMEX', 'ME', 'flow', 'last_30_business_days', 'B', 'lblFecha2'),
}


def parse_market(html, *, target, retrieved_at):
    soup = BeautifulSoup(html, 'html.parser')
    if re.search(r'No\s+existe\s+informaci[oó]n\s+para\s+la\s+fecha\s+elegida', soup.get_text(' ', strip=True), re.I):
        raise PeriodUnavailableError('SBS no publica información para la fecha elegida.')
    rows = []
    for metric, (value_id, currency, basis, window, scope, date_id) in SERIES.items():
        title = soup.find(id='ctl00_cphContent_' + date_id)
        if title is None:
            raise SchemaChangedError(f'Falta la fecha efectiva de {metric}.')
        effective = re.search(r'\bal\s+(\d{2}/\d{2}/\d{4})', title.get_text(' ', strip=True), re.I)
        if effective is None:
            raise SchemaChangedError(f'Fecha efectiva ilegible para {metric}.')
        if _parse_date(effective.group(1)) != target:
            raise PeriodUnavailableError(f'{metric}: solicitado {target}, publicado {effective.group(1)}.')
        value = soup.find(id='ctl00_cphContent_' + value_id)
        text = value.get_text(strip=True) if value else ''
        if not re.fullmatch(r'\d+(?:[.,]\d+)?', text):
            raise SchemaChangedError(f'Tasa ausente o no numérica para {metric}.')
        rows.append({
            'period': target.isoformat(), 'period_date': target.isoformat(), 'frequency': 'daily',
            'metric': metric, 'currency': currency, 'rate': float(text.replace(',', '.')),
            'unit': 'percent_effective_annual', 'basis': basis, 'observation_window': window,
            'entity_scope': scope, 'reference_kind': 'market_aggregate',
            'source': 'SBS', 'source_url': URL, 'methodology_url': METHODOLOGY_URL,
            'retrieved_at': retrieved_at,
        })
    frame = pd.DataFrame(rows)
    for col in frame.columns:
        frame[col] = frame[col].astype('float64' if col == 'rate' else 'string')
    return frame


class PassiveMarketProvider(DatasetProvider):
    def __init__(self, spec):
        super().__init__(spec)
        self.transport = CurlChromeTransport(state_dir=self.storage.state_root,
            normal_marker='ctl00_cphContent_btnConsultar', timeout=30)
        self.html = None
        self.form_state = {}

    def single_request(self, *, fecha=None, **query):
        if fecha is None or query:
            raise InvalidQueryError('Use fecha=YYYY-MM-DD.')
        target = _parse_date(fecha)
        if target > date.today():
            raise InvalidQueryError('La fecha no puede ser futura.')
        return PeriodRequest(target.isoformat(), f'year={target.year}', {'fecha': target.isoformat()},
            mutable=target >= date.today() - timedelta(days=7))

    def plan_sync(self, *, desde, hasta=None, **query):
        if query:
            raise InvalidQueryError(f'Parámetros desconocidos: {sorted(query)}')
        start, end = _parse_date(desde), _parse_date(hasta or date.today())
        if end < start or end > date.today():
            raise InvalidQueryError('Rango de fechas inválido.')
        while start <= end:
            if start.weekday() < 5:
                yield self.single_request(fecha=start)
            start += timedelta(days=1)

    def _fetch_period(self, request):
        target = _parse_date(request.params['fecha'])
        try:
            if self.html is None:
                response = self.transport.request('GET', URL)
                self.html = response.text
                self.form_state = successful_controls(self.html)
                if '__VIEWSTATE' not in self.form_state:
                    raise SchemaChangedError('La página de mercado no contiene el estado WebForms.')
            try:
                data = parse_market(self.html, target=target, retrieved_at=utc_now_iso())
                raw = self.html
            except PeriodUnavailableError:
                iso, display = target.isoformat(), target.strftime('%d/%m/%Y')
                state_date = iso + '-00-00-00'
                payload = dict(self.form_state)
                for key in list(payload):
                    if key.endswith('$btnExportar'):
                        payload.pop(key)
                payload.update({
                    'ctl00$cphContent$rdpDate': iso,
                    'ctl00$cphContent$rdpDate$dateInput': display,
                    'ctl00_cphContent_rdpDate_dateInput_ClientState': json.dumps({
                        'enabled': True, 'emptyMessage': '', 'validationText': state_date,
                        'valueAsString': state_date, 'minDateStr': '1000-01-01-00-00-00',
                        'maxDateStr': '2099-12-31-00-00-00', 'lastSetTextBoxValue': display}),
                    '__EVENTTARGET': '', '__EVENTARGUMENT': '', '__ASYNCPOST': 'true',
                    'ctl00$MainScriptManager': 'ctl00$cphContent$updConsulta|ctl00$cphContent$btnConsultar',
                    'ctl00$cphContent$btnConsultar': 'Consultar',
                })
                payload.pop('ctl00_MainScriptManager_TSM', None)
                response = self.transport.request('POST', URL, data=payload)
                raw = response.text
                try:
                    self.html, hidden = parse_delta_response(raw)
                except ValueError:
                    self.html, hidden = raw, {}
                self.form_state = merge_hidden_state(payload, self.html, hidden)
                data = parse_market(self.html, target=target, retrieved_at=utc_now_iso())
            return FetchResult(self.spec.dataset_id, data, {'fecha': target.isoformat()}, raw)
        except Exception:
            self.html, self.form_state = None, {}
            self.transport.reset()
            raise

    def filter_loaded(self, data, *, desde=None, hasta=None, metric=None):
        for bound, end in ((desde, False), (hasta, True)):
            if bound is not None:
                value = _parse_date(bound).isoformat()
                data = data[data.period <= value] if end else data[data.period >= value]
        if metric is not None:
            wanted = [metric] if isinstance(metric, str) else list(metric)
            if set(wanted) - SERIES.keys():
                raise InvalidQueryError('Indicadores admitidos: TIPMN, TIPMEX, FTIPMN, FTIPMEX.')
            data = data[data.metric.isin(wanted)]
        return data.drop(columns=['_period_key'], errors='ignore')
