"""Public observations from FRED, Treasury, Yahoo and BCB Focus."""
from datetime import datetime, date, timezone
import re
from zoneinfo import ZoneInfo
from urllib.parse import quote, urlencode

from fuentes_financieras.exceptions import InvalidQueryError, SchemaChangedError, SourceUnavailableError
from ._series import SeriesProvider, csv_rows, iso_period, numeric


class FredProvider(SeriesProvider):
    source_name = 'Federal Reserve Bank of St. Louis - FRED'
    endpoint = 'https://fred.stlouisfed.org/graph/fredgraph.csv'

    def validate_code(self, code):
        if not isinstance(code, str) or not re.fullmatch(r'[A-Za-z0-9_]{1,80}', code):
            raise InvalidQueryError('Código individual FRED inválido.')

    def observations(self, code, frequency, year):
        text = self.get(self.endpoint, params={'id': code, 'cosd': f'{year}-01-01', 'coed': f'{year}-12-31'}).text
        table = csv_rows(text, {code})
        date_column = 'observation_date' if not table or 'observation_date' in table[0] else 'DATE'
        if table and date_column not in table[0]:
            raise SchemaChangedError('Fecha FRED ausente.')
        rows = [dict(date=iso_period(r[date_column]), value=numeric(r[code]), source_value=r[code]) for r in table]
        return rows, code, self.endpoint+'?id='+code


class TreasuryProvider(SeriesProvider):
    source_name = 'U.S. Department of the Treasury'
    endpoint = 'https://home.treasury.gov/resource-center/data-chart-center/interest-rates/daily-treasury-rates.csv/{year}/all'
    frequencies = {'daily'}

    def validate_code(self, code):
        if not isinstance(code, str) or not re.fullmatch(r'(?:1|1\.5|2|3|4|6) Mo|(?:1|2|3|5|7|10|20|30) Yr', code):
            raise InvalidQueryError('Use un plazo Treasury, por ejemplo 3 Mo o 10 Yr.')

    def observations(self, code, frequency, year):
        url = self.endpoint.format(year=year)
        text = self.get(url, params={'type': 'daily_treasury_yield_curve', 'field_tdr_date_value': year, '_format': 'csv'}).text
        table = csv_rows(text, {'Date', code})
        rows = []
        for row in table:
            try:
                observed = datetime.strptime(row['Date'], '%m/%d/%Y').date().isoformat()
            except (TypeError, ValueError):
                raise SchemaChangedError('Fecha Treasury incompatible.') from None
            rows.append(dict(date=observed, value=numeric(row[code]), source_value=row[code], unit='percent_par_yield'))
        return rows, f'Treasury par yield - {code}', url


class YahooProvider(SeriesProvider):
    source_name = 'Yahoo Finance'
    endpoint = 'https://query1.finance.yahoo.com/v8/finance/chart'
    frequencies = {'daily'}

    def validate_code(self, code):
        if not isinstance(code, str) or not re.fullmatch(r'[A-Za-z0-9^=._-]{1,40}', code):
            raise InvalidQueryError('Símbolo individual Yahoo inválido.')

    def observations(self, code, frequency, year):
        url = self.endpoint+'/'+quote(code, safe='')
        start = int(datetime(year, 1, 1, tzinfo=timezone.utc).timestamp())
        end = int(datetime(year+1, 1, 1, tzinfo=timezone.utc).timestamp())
        payload = self.json(url, params={'period1': start, 'period2': end, 'interval': '1d', 'events': 'history'})
        try:
            if payload['chart'].get('error'):
                raise ValueError
            result = payload['chart']['result']
            if len(result) != 1:
                raise ValueError
            result = result[0]
            meta = result['meta']
            if meta['symbol'].upper() != code.upper():
                raise ValueError
            zone = ZoneInfo(meta['exchangeTimezoneName'])
            timestamps = result.get('timestamp') or []
            closes = result['indicators']['quote'][0]['close']
            if len(timestamps) != len(closes):
                raise ValueError
        except (KeyError, TypeError, ValueError, IndexError):
            raise SchemaChangedError('Identidad, horario o vector de precios Yahoo incompatible.') from None
        session = (meta.get('currentTradingPeriod') or {}).get('regular') or {}
        now = datetime.now(timezone.utc).timestamp()
        incomplete = datetime.fromtimestamp(session['start'], zone).date() if session.get('end', 0) > now and session.get('start') else None
        rows = []
        for timestamp, close in zip(timestamps, closes):
            observed = datetime.fromtimestamp(timestamp, zone).date()
            # Today is not a completed close unless the source identifies the session end.
            if observed == incomplete or (observed == datetime.now(zone).date() and not session.get('end')):
                continue
            rows.append(dict(date=observed.isoformat(), value=numeric(close), source_value='' if close is None else str(close),
                unit='index_points' if meta.get('instrumentType') == 'INDEX' else meta.get('currency', 'source_native')))
        return rows, meta.get('longName') or meta.get('shortName') or code, url


class FocusProvider(SeriesProvider):
    source_name = 'Banco Central do Brasil - Focus'
    endpoint = 'https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoInflacao12Meses'
    frequencies = {'daily'}

    def validate_code(self, code):
        if code != 'IPCA12':
            raise InvalidQueryError('Este adaptador Focus admite IPCA12, suavizada S y baseCalculo 0.')

    def observations(self, code, frequency, year):
        # The bound is checked, never silently interpreted as complete pagination.
        # Olinda's filter parser requires percent-encoded spaces instead of plus signs.
        params = {'$format': 'json', '$top': '10000', '$orderby': 'Data asc',
            '$select': 'Data,Mediana', '$filter': f"Indicador eq 'IPCA' and Suavizada eq 'S' and baseCalculo eq 0 and Data ge '{year}-01-01' and Data le '{year}-12-31'"}
        payload = self.json(self.endpoint+'?'+urlencode(params, quote_via=quote))
        rows = payload.get('value')
        if not isinstance(rows, list) or len(rows) >= 10000 or payload.get('@odata.nextLink') or payload.get('odata.nextLink'):
            raise SchemaChangedError('Focus incompleto o con paginación pendiente.')
        try:
            output = [dict(date=iso_period(r['Data']), value=numeric(r['Mediana']), source_value=str(r['Mediana']), unit='percent_expected_inflation_12m') for r in rows]
        except (KeyError, TypeError):
            raise SchemaChangedError('Observación Focus incompleta.') from None
        return output, 'Focus - mediana IPCA a 12 meses, suavizada S, base 0', self.endpoint
