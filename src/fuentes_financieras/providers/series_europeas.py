"""Independent European statistical series, with explicit dimensional checks."""
from datetime import date, datetime
import csv
import io
import re
from urllib.parse import parse_qsl, urlencode, quote

from fuentes_financieras.exceptions import InvalidQueryError, SchemaChangedError
from ._series import SeriesProvider, csv_rows, iso_period, numeric


class BceProvider(SeriesProvider):
    source_name = 'Banco Central Europeo'
    endpoint = 'https://data-api.ecb.europa.eu/service/data'

    def validate_code(self, code):
        if not isinstance(code, str) or not re.fullmatch(r'[A-Z0-9_]+/[A-Za-z0-9_.-]+', code):
            raise InvalidQueryError('Use flujo/clave individual BCE, sin comodines ni selección múltiple.')

    def observations(self, code, frequency, year):
        expected = {'daily': 'D', 'monthly': 'M', 'quarterly': 'Q', 'annual': 'A'}[frequency]
        if code.split('/')[1].split('.')[0] != expected:
            raise InvalidQueryError('Frecuencia distinta de la clave BCE.')
        url = f'{self.endpoint}/{code}'
        text = self.get(url, params={'startPeriod': f'{year}-01-01', 'endPeriod': f'{year}-12-31', 'format': 'csvdata'}).text
        data = csv_rows(text, {'TIME_PERIOD', 'OBS_VALUE', 'KEY'})
        if any(r['KEY'] != code.replace('/', '.', 1) for r in data):
            raise SchemaChangedError('La respuesta BCE corresponde a otra serie o a varias claves.')
        rows = [dict(date=iso_period(r['TIME_PERIOD']), period=r['TIME_PERIOD'], value=numeric(r['OBS_VALUE']),
                     source_value=r['OBS_VALUE'], unit=r.get('UNIT') or 'source_native',
                     source_unit_multiplier=r.get('UNIT_MULT', ''), status=r.get('OBS_STATUS', '')) for r in data]
        return rows, (data[0].get('TITLE') or data[0].get('TITLE_COMPL') or code) if data else code, url


class EurostatProvider(SeriesProvider):
    source_name = 'Eurostat'
    endpoint = 'https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data'
    frequencies = {'monthly', 'quarterly', 'annual'}

    def validate_code(self, code):
        if not isinstance(code, str) or len(code) > 1000:
            raise InvalidQueryError('Código Eurostat inválido.')
        dataset, _, query = code.partition('?')
        pairs = parse_qsl(query, keep_blank_values=True)
        if not re.fullmatch(r'[a-z0-9_]+', dataset) or len(dict(pairs)) != len(pairs):
            raise InvalidQueryError('Dataset o filtros Eurostat inválidos.')
        if any(not re.fullmatch(r'[A-Za-z0-9_]+', key) or not re.fullmatch(r'[A-Za-z0-9_.-]+', val) for key, val in pairs):
            raise InvalidQueryError('Filtros Eurostat deben seleccionar códigos individuales.')
        if set(dict(pairs)) & {'time', 'sinceTimePeriod', 'untilTimePeriod', 'lang'}:
            raise InvalidQueryError('Los períodos Eurostat se controlan mediante desde/hasta.')

    def observations(self, code, frequency, year):
        dataset, _, query = code.partition('?')
        url = f'{self.endpoint}/{dataset}'
        params = dict(parse_qsl(query))
        params.update(sinceTimePeriod=str(year), untilTimePeriod=str(year), lang='EN')
        payload = self.json(url, params=params)
        try:
            dims, sizes = payload['id'], payload['size']
            if len(dims) != len(sizes) or 'time' not in dims or any(n != 1 for d, n in zip(dims, sizes) if d != 'time'):
                raise ValueError
            observed_frequency = payload['dimension']['freq']['category']['index']
            if set(observed_frequency) != {{'monthly': 'M', 'quarterly': 'Q', 'annual': 'A'}[frequency]}:
                raise ValueError
            times = payload['dimension']['time']['category']['index']
            if isinstance(times, list):
                times = dict(zip(times, range(len(times))))
            if sorted(times.values()) != list(range(sizes[dims.index('time')])):
                raise ValueError
            values = payload.get('value', {})
            statuses = payload.get('status', {})
            unit = next(iter(payload['dimension'].get('unit', {}).get('category', {}).get('index', {})), 'source_native')
            def indexed(container, index):
                return container[index] if isinstance(container, list) else container.get(str(index))
            rows = [dict(date=iso_period(p), period=p, value=numeric(indexed(values, i)),
                source_value='' if indexed(values, i) is None else str(indexed(values, i)), unit=unit,
                status=indexed(statuses, i) or '') for p, i in times.items()]
        except (KeyError, TypeError, ValueError, IndexError):
            raise SchemaChangedError('Dimensiones JSON-stat ambiguas o frecuencia incompatible.') from None
        return rows, payload.get('label', dataset), url+'?'+query


class OnsProvider(SeriesProvider):
    source_name = 'Office for National Statistics'
    endpoint = 'https://www.ons.gov.uk/generator'
    frequencies = {'monthly', 'quarterly', 'annual'}

    def validate_code(self, code):
        if not isinstance(code, str) or not re.fullmatch(r'/[a-z0-9/]+/timeseries/[a-z0-9]+/[a-z0-9]+', code):
            raise InvalidQueryError('Use la ruta individual timeseries de ONS.')

    def observations(self, code, frequency, year):
        text = self.get(self.endpoint, params={'format': 'csv', 'uri': code}).text
        table = list(csv.reader(io.StringIO(text.lstrip('\ufeff'))))
        if not table or table[0][0] != 'Title':
            raise SchemaChangedError('CSV ONS sin descripción de serie.')
        patterns = {'monthly': r'(\d{4}) ([A-Z]{3})', 'quarterly': r'(\d{4}) Q([1-4])', 'annual': r'(\d{4})'}
        months = dict(zip('JAN FEB MAR APR MAY JUN JUL AUG SEP OCT NOV DEC'.split(), range(1, 13)))
        rows = []
        for row in table:
            if not row or not re.fullmatch(patterns[frequency], row[0]):
                continue
            if len(row) != 2:
                raise SchemaChangedError('Fila ONS con dimensiones inesperadas.')
            match = re.fullmatch(patterns[frequency], row[0])
            y = int(match[1])
            try:
                month = months[match[2]] if frequency == 'monthly' else (int(match[2])-1)*3+1 if frequency == 'quarterly' else 1
            except KeyError:
                raise SchemaChangedError('Mes ONS no reconocido.') from None
            rows.append(dict(date=date(y, month, 1).isoformat(), period=row[0], value=numeric(row[1]), source_value=row[1]))
        return rows, table[0][1], self.endpoint+'?'+urlencode({'format': 'csv', 'uri': code})


class BundesbankProvider(SeriesProvider):
    source_name = 'Deutsche Bundesbank'
    endpoint = 'https://api.statistiken.bundesbank.de/rest/data/BBSIS'
    frequencies = {'daily'}

    def validate_code(self, code):
        if not isinstance(code, str) or not re.fullmatch(r'D\.[A-Za-z0-9_.]+', code):
            raise InvalidQueryError('Use una clave diaria individual BBSIS.')

    def observations(self, code, frequency, year):
        url = f'{self.endpoint}/{code}'
        text = self.get(url, params={'startPeriod': f'{year}-01-01', 'endPeriod': f'{year}-12-31', 'format': 'csv'}).text
        delimiter = ';' if any(re.match(r'^"?\d{4}-\d{2}-\d{2}"?;', line) for line in text.splitlines()) else ','
        table = list(csv.reader(io.StringIO(text.lstrip('\ufeff')), delimiter=delimiter))
        expected_key = 'BBSIS.'+code
        if not table or not any(expected_key in row for row in table[:12]):
            raise SchemaChangedError('Identidad o formato CSV Bundesbank no reconocido.')
        unit = next((r[1] for r in table[:12] if len(r)>1 and r[0].lower()=='unit'), 'source_native')
        rows = []
        for parts in table:
            if not parts or not re.fullmatch(r'\d{4}-\d{2}-\d{2}', parts[0]):
                continue
            if len(parts) < 2:
                raise SchemaChangedError('Observación Bundesbank incompleta.')
            rows.append(dict(date=iso_period(parts[0]), value=numeric(parts[1], comma_decimal=delimiter==';'),
                             source_value=parts[1], unit=unit, status=parts[2] if len(parts)>2 else ''))
        return rows, code, url


class BoeProvider(SeriesProvider):
    source_name = 'Bank of England'
    endpoint = 'https://www.bankofengland.co.uk/boeapps/database/_iadb-fromshowcolumns.asp'
    frequencies = {'daily'}

    def validate_code(self, code):
        if not isinstance(code, str) or not re.fullmatch(r'[A-Z0-9]{5,15}', code):
            raise InvalidQueryError('Código individual IADB inválido.')

    def observations(self, code, frequency, year):
        text = self.get(self.endpoint, params={'csv.x': 'yes', 'Datefrom': f'01/Jan/{year}', 'Dateto': f'31/Dec/{year}',
            'SeriesCodes': code, 'CSVF': 'TN', 'UsingCodes': 'Y', 'VPD': 'Y', 'VFD': 'N'}).text
        table = csv_rows(text, {'DATE', code})
        rows = []
        for r in table:
            try:
                observed = datetime.strptime(r['DATE'].strip(), '%d %b %Y').date().isoformat()
            except (ValueError, TypeError):
                raise SchemaChangedError('Fecha IADB ilegible.') from None
            rows.append(dict(date=observed, value=numeric(r[code]), source_value=r[code]))
        return rows, code, self.endpoint
