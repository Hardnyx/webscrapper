"""Dated observations of the current SBS deposit-taking universe."""
from datetime import date, datetime
from zoneinfo import ZoneInfo

import pandas as pd

from fuentes_financieras.exceptions import InvalidQueryError, SchemaChangedError
from fuentes_financieras.models import FetchResult, PeriodRequest
from fuentes_financieras.provider import DatasetProvider
from fuentes_financieras.storage import utc_now_iso
from fuentes_financieras.transports.curl_chrome import CurlChromeTransport
from ._universe_parser import parse_entities

URL = 'https://www.sbs.gob.pe/app/pp/empresasweb/Paginas/EmpCaptarDep.aspx'


def today():
    return datetime.now(ZoneInfo('America/Lima')).date()


class DepositUniverseProvider(DatasetProvider):
    def single_request(self, **query):
        if query:
            raise InvalidQueryError('La fuente solo publica el universo vigente; fetch/sync no admiten fechas históricas.')
        observed = today().isoformat()
        return PeriodRequest(observed, observed[:4], {}, mutable=True)

    def plan_sync(self, **query):
        yield self.single_request(**query)

    def _fetch_period(self, request):
        transport = CurlChromeTransport(state_dir=self.storage.state_root, timeout=30)
        response = transport.request('GET', URL)
        # Do not stamp yesterday's date if a request crosses midnight.
        if today().isoformat() != request.period_key:
            raise InvalidQueryError('Cambió la fecha de observación; repita la consulta.')
        return self.from_html(response.text, observed_on=request.period_key)

    def from_html(self, html, *, observed_on):
        """Parse an explicitly dated local capture without making a network request."""
        observed = date.fromisoformat(str(observed_on))
        if observed > today():
            raise InvalidQueryError('La fecha de captura no puede ser futura.')
        try:
            entities = parse_entities(html)
        except ValueError as exc:
            raise SchemaChangedError(str(exc)) from exc
        retrieved = utc_now_iso()
        data = pd.DataFrame([{
            'period': observed.isoformat(), 'period_date': observed.isoformat(),
            'entity_type_code': e.type_code, 'entity_type': e.entity_type,
            'entity_name': e.sbs_name, 'normalized_name': e.normalized_name,
            'source_order': e.source_order, 'source': 'SBS',
            'source_url': URL, 'retrieved_at': retrieved,
        } for e in entities])
        for col in data.columns:
            data[col] = data[col].astype('Int64' if col == 'source_order' else 'string')
        return FetchResult(self.spec.dataset_id, data, {
            'observed_on': observed.isoformat(),
            'counts': {str(k): int(v) for k, v in data.groupby('entity_type_code').size().items()},
            'scope': 'current_public_universe_observation',
        }, raw=html)

    def import_capture(self, path, *, observed_on, force=False, keep_raw=True):
        """Persist a dated capture through the same cache and integrity checks."""
        from pathlib import Path
        fetched = self.from_html(Path(path).read_text(encoding='utf-8-sig'), observed_on=observed_on)
        period = fetched.metadata['observed_on']
        # A dedicated offline provider avoids changing the live provider's behavior.
        class CaptureProvider(DepositUniverseProvider):
            def single_request(self, **query):
                return PeriodRequest(period, period[:4], {}, mutable=False)

            def _fetch_period(self, request):
                return fetched

        offline = CaptureProvider(self.spec)
        offline.storage = self.storage
        return offline.sync(force=force, keep_raw=keep_raw)

    def filter_loaded(self, data, *, desde=None, hasta=None, tipo=None, latest=False):
        for bound, is_end in ((desde, False), (hasta, True)):
            if bound is not None:
                value = date.fromisoformat(str(bound)).isoformat()
                data = data[data.period <= value] if is_end else data[data.period >= value]
        if tipo is not None:
            if tipo not in {'B', 'F', 'C', 'R'}:
                raise InvalidQueryError('tipo debe ser B, F, C o R.')
            data = data[data.entity_type_code == tipo]
        if latest and not data.empty:
            data = data[data.period == data.period.max()]
        return data
