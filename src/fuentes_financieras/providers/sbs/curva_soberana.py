"""Published SBS zero-coupon curves, without interpolation or consumer spreads."""
from datetime import date, timedelta
import json
import math

import pandas as pd

from fuentes_financieras.exceptions import InvalidQueryError, PeriodUnavailableError, SchemaChangedError, SourceUnavailableError
from fuentes_financieras.models import FetchResult, PeriodRequest
from fuentes_financieras.provider import DatasetProvider
from fuentes_financieras.storage import utc_now_iso
from fuentes_financieras.transports.curl_chrome import CurlChromeTransport
from .._series import bound

URL = 'https://www.sbs.gob.pe/app/pp/n_CurvaSoberana/BuscarInformacionHistorica'
REFERER = 'https://www.sbs.gob.pe/app/pp/n_CurvaSoberana/CurvaSoberana/ConsultaHistorica'
CURVES = {'CCPSS', 'CBCRPS', 'CCPVS', 'CCPEDS', 'CBCRS', 'CCSDF', 'CSBCRD', 'CCINFS', 'CCCLD'}


def parse_curve(payload, *, curve, target, retrieved_at):
    if not isinstance(payload, dict) or payload.get('success') is not True or not isinstance(payload.get('data'), list):
        raise SchemaChangedError('Respuesta de curva SBS incompatible o sin éxito confirmado.')
    if not payload['data']:
        raise PeriodUnavailableError('SBS no devolvió puntos para la fecha y curva solicitadas.')
    rows = []
    for row in payload['data']:
        try:
            observed = date.fromisoformat(str(row['fechaProceso'])[:10])
            tenor, rate = row['plazos'], row['tasa']
            if observed.isoformat() != target or row.get('tipoCurva', curve) != curve or isinstance(tenor, bool) or isinstance(rate, bool):
                raise ValueError
            tenor, rate = float(tenor), float(rate)
            if not math.isfinite(rate) or not math.isfinite(tenor) or tenor < 0 or tenor != int(tenor):
                raise ValueError
        except (KeyError, TypeError, ValueError):
            raise SchemaChangedError('Punto de curva inválido o fecha distinta de la solicitada.') from None
        rows.append(dict(curve_code=curve, period=target, period_date=target, frequency='daily',
            tenor_days=int(tenor), rate=rate, unit='percent_annual_as_published',
            source='SBS', source_url=URL, retrieved_at=retrieved_at))
    data = pd.DataFrame(rows)
    if data.tenor_days.duplicated().any():
        raise SchemaChangedError('Plazos duplicados en la curva publicada.')
    for column in data:
        data[column] = data[column].astype('int64' if column == 'tenor_days' else 'float64' if column == 'rate' else 'string')
    return data.sort_values('tenor_days').reset_index(drop=True)


class SovereignCurveProvider(DatasetProvider):
    parser_version = '2026-10-23.1'
    contract_version = '2'

    def __init__(self, spec):
        super().__init__(spec)
        self.transport = CurlChromeTransport(state_dir=self.storage.state_root, timeout=120)

    def single_request(self, *, fecha, curva='CCPSS'):
        if curva not in CURVES:
            raise InvalidQueryError('Código de curva SBS no admitido.')
        target = bound(fecha).isoformat()
        return PeriodRequest(f'{curva}:{target}', f'curva={curva}/year={target[:4]}',
                             {'fecha': target, 'curva': curva}, mutable=True)

    def plan_sync(self, *, desde, hasta=None, curvas=('CCPSS',)):
        start, end = bound(desde), bound(hasta or desde, True)
        selected = [curvas] if isinstance(curvas, str) else list(curvas)
        if not selected or len(set(selected)) != len(selected) or set(selected)-CURVES or end < start or (end-start).days > 3660:
            raise InvalidQueryError('Curvas o rango inválidos; máximo diez años por plan.')
        while start <= end:
            if start.weekday() < 5:
                for curve in selected:
                    yield self.single_request(fecha=start.isoformat(), curva=curve)
            start += timedelta(days=1)

    def _fetch_period(self, request):
        target, curve = request.params['fecha'], request.params['curva']
        response = self.transport.request('POST', URL, ajax=False, json={'tipoCurva': curve, 'fechaInicio': target, 'fechaFin': target},
                                          headers={'Referer': REFERER})
        if len(response.content) > 5_000_000:
            raise SchemaChangedError('Respuesta de curva mayor de 5 MB.')
        try:
            payload = json.loads(response.content.decode('utf-8-sig'))
        except (UnicodeError, ValueError):
            raise SchemaChangedError('La curva no devolvió JSON válido.') from None
        data = parse_curve(payload, curve=curve, target=target, retrieved_at=utc_now_iso())
        return FetchResult(self.spec.dataset_id, data, metadata={'scope': 'published_points_without_interpolation'}, raw=response.content)

    def filter_loaded(self, data, *, desde=None, hasta=None, curvas=None):
        if desde is not None:
            data = data[data.period_date.ge(bound(desde).isoformat())]
        if hasta is not None:
            data = data[data.period_date.le(bound(hasta, True).isoformat())]
        if curvas is not None:
            data = data[data.curve_code.isin([curvas] if isinstance(curvas, str) else curvas)]
        return data.drop(columns='_period_key', errors='ignore')
