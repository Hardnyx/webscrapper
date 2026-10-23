"""Independent BCRP, SGS, SIE and BDE adapters; no report configuration."""
from datetime import date, datetime
import os
import re

from fuentes_financieras.exceptions import InvalidQueryError, SchemaChangedError, SourceUnavailableError
from ._series import SeriesProvider, numeric

MONTHS = dict(zip('ene feb mar abr may jun jul ago set oct nov dic'.split(), range(1, 13)))
MONTHS['sep'] = 9


def localized_date(token, separator='/'):
    try:
        return datetime.strptime(token, '%d'+separator+'%m'+separator+'%Y').date().isoformat()
    except (ValueError, TypeError):
        raise SchemaChangedError('Fecha de banco central ilegible.') from None


class BcrpProvider(SeriesProvider):
    source_name = 'BCRPData'
    endpoint = 'https://estadisticas.bcrp.gob.pe/estadisticas/series/api'
    frequencies = {'daily', 'monthly'}

    def validate_code(self, code):
        if not isinstance(code, str) or not re.fullmatch(r'[A-Z]{2}\d{5}[A-Z]{2}', code):
            raise InvalidQueryError('Indique un código individual BCRPData.')

    def observations(self, code, frequency, year):
        if code[-1] != {'daily': 'D', 'monthly': 'M'}[frequency]:
            raise InvalidQueryError('Frecuencia distinta del sufijo BCRPData.')
        start, end = (f'{year}-1-1', f'{year}-12-31') if frequency == 'daily' else (f'{year}-1', f'{year}-12')
        url = f'{self.endpoint}/{code}/json/{start}/{end}/esp'
        payload = self.json(url)
        try:
            config = payload['config']['series']
            periods = payload['periods']
            if len(config) != 1 or not isinstance(periods, list):
                raise ValueError
            title = config[0]['name']
        except (KeyError, TypeError, ValueError):
            raise SchemaChangedError('Serie o períodos BCRPData incompatibles.') from None
        rows = []
        for row in periods:
            try:
                parts = row['name'].replace(' ', '').split('.')
                if frequency == 'daily':
                    day, mon, yy = parts
                    y = int(yy)
                    # Resolve two-digit years against the requested capture, not today's clock.
                    if y < 100:
                        y += year//100*100
                    observed = date(y, MONTHS[mon[:3].lower()], int(day))
                else:
                    mon, yy = parts
                    observed = date(int(yy), MONTHS[mon[:3].lower()], 1)
                values = row['values']
                if len(values) != 1:
                    raise ValueError
            except (KeyError, ValueError, TypeError):
                raise SchemaChangedError('Observación BCRPData incompleta o frecuencia distinta.') from None
            rows.append(dict(date=observed.isoformat(), period=row['name'], value=numeric(values[0], thousands=True), source_value=str(values[0])))
        return rows, title, url


class BcbProvider(SeriesProvider):
    source_name = 'Banco Central do Brasil - SGS'
    endpoint = 'https://api.bcb.gov.br/dados/serie/bcdata.sgs.{code}/dados'

    def validate_code(self, code):
        if not isinstance(code, str) or not re.fullmatch(r'[1-9]\d{0,8}', code):
            raise InvalidQueryError('Indique un código numérico SGS.')

    def observations(self, code, frequency, year):
        url = self.endpoint.format(code=code)
        payload = self.json(url, params={'formato': 'json', 'dataInicial': f'01/01/{year}', 'dataFinal': f'31/12/{year}'})
        if not isinstance(payload, list):
            raise SchemaChangedError('SGS no devolvió una lista de observaciones.')
        try:
            rows = [dict(date=localized_date(r['data']), value=numeric(r['valor']), source_value=str(r['valor'])) for r in payload]
        except (KeyError, TypeError):
            raise SchemaChangedError('Observación SGS incompleta.') from None
        return rows, f'SGS {code}', url


class BanxicoProvider(SeriesProvider):
    source_name = 'Banco de México - SIE'
    endpoint = 'https://www.banxico.org.mx/SieAPIRest/service/v1/series'

    def validate_code(self, code):
        if not isinstance(code, str) or not re.fullmatch(r'[A-Z]{2}\d{1,10}', code):
            raise InvalidQueryError('Indique un identificador individual SIE.')

    def observations(self, code, frequency, year):
        token = os.environ.get('BANXICO_TOKEN', '')
        if not token:
            raise SourceUnavailableError('Falta BANXICO_TOKEN en el entorno local.')
        url = f'{self.endpoint}/{code}/datos/{year}-01-01/{year}-12-31'
        payload = self.json(url, headers={'Bmx-Token': token})
        try:
            series = payload['bmx']['series']
            if len(series) != 1 or series[0]['idSerie'] != code:
                raise ValueError
            series = series[0]
            rows = [dict(date=localized_date(r['fecha']), value=numeric(r['dato'], thousands=True), source_value=str(r['dato'])) for r in series['datos']]
            title = ' '.join(series['titulo'].split())
        except (KeyError, TypeError, ValueError):
            raise SchemaChangedError('Identidad o estructura SIE incompatible.') from None
        return rows, title, url


class BcchProvider(SeriesProvider):
    source_name = 'Banco Central de Chile - BDE'
    endpoint = 'https://si3.bcentral.cl/SieteRestWS/SieteRestWS.ashx'

    def validate_code(self, code):
        if not isinstance(code, str) or not re.fullmatch(r'[A-Za-z0-9_.]{1,150}', code):
            raise InvalidQueryError('Identificador BDE inválido.')

    def observations(self, code, frequency, year):
        user, password = os.environ.get('BCCH_USER'), os.environ.get('BCCH_PASSWORD')
        if not user or not password:
            raise SourceUnavailableError('Faltan BCCH_USER y BCCH_PASSWORD en el entorno local.')
        response = self.get(self.endpoint, params={'user': user, 'pass': password, 'function': 'GetSeries',
            'timeseries': code, 'firstdate': f'{year}-01-01', 'lastdate': f'{year}-12-31'})
        # This API sometimes declares Latin text. Never store its request URL.
        import json
        payload = None
        for encoding in ('utf-8-sig', 'cp1252'):
            try:
                payload = json.loads(response.content.decode(encoding))
                break
            except (UnicodeError, ValueError):
                continue
        if not isinstance(payload, dict):
            raise SchemaChangedError('BDE devolvió contenido no JSON.')
        if payload.get('Codigo') not in (0, '0'):
            raise SourceUnavailableError('BDE rechazó credenciales o consulta; revisar configuración local.')
        series = payload.get('Series')
        if isinstance(series, list):
            if len(series) != 1:
                raise SchemaChangedError('BDE devolvió varias series.')
            series = series[0]
        try:
            rows = [dict(date=localized_date(r['indexDateString'].replace('/', '-'), '-'),
                value=numeric(r['value'], comma_decimal=True) if r.get('statusCode', 'OK') in ('OK', '') else float('nan'), source_value=str(r['value']), status=r.get('statusCode', ''),
                flags='' if r.get('statusCode', 'OK') in ('OK', '') else 'source_status_not_ok') for r in series['Obs']]
            title = series.get('descripEsp') or series.get('descripIng') or code
        except (TypeError, KeyError):
            raise SchemaChangedError('Observaciones BDE incompletas.') from None
        return rows, title, self.endpoint
