"""Bounded annual captures for independent public time-series providers."""
from calendar import monthrange
from datetime import date
from hashlib import sha256
import csv
import io
import json
import math
import re

import pandas as pd

from fuentes_financieras.exceptions import InvalidQueryError, PeriodUnavailableError, SchemaChangedError, SourceUnavailableError, WAFBlockedError
from fuentes_financieras.models import FetchResult, PeriodRequest
from fuentes_financieras.provider import DatasetProvider
from fuentes_financieras.storage import utc_now_iso
from fuentes_financieras.transports.curl_chrome import CurlChromeTransport

FREQUENCIES = {'daily', 'monthly', 'quarterly', 'annual'}
MISSING = {'', '.', '-', 'n/a', 'na', 'nan', 'nd', 'n.d.', 'n.d', 'n/e', 's.i.'}


def bound(value, end=False):
    try:
        text = str(value)
        if re.fullmatch(r'\d{4}-\d{2}', text):
            year, month = map(int, text.split('-'))
            return date(year, month, monthrange(year, month)[1] if end else 1)
        return date.fromisoformat(text)
    except (ValueError, TypeError):
        raise InvalidQueryError('Use fechas YYYY-MM-DD o meses YYYY-MM.') from None


def numeric(value, *, comma_decimal=False, thousands=False):
    token = '' if value is None else str(value).strip()
    if token.lower() in MISSING:
        return float('nan')
    if isinstance(value, bool):
        raise SchemaChangedError('Valor booleano en una observación numérica.')
    text = token.replace(',', '') if thousands else token.replace(',', '.') if comma_decimal else token
    try:
        number = float(text)
    except (TypeError, ValueError):
        raise SchemaChangedError('Valor numérico de la fuente no reconocido.') from None
    if not math.isfinite(number):
        raise SchemaChangedError('Valor numérico no finito en la fuente.')
    return number


def iso_period(token):
    text = str(token).strip()
    if re.fullmatch(r'\d{4}-Q[1-4]', text):
        return date(int(text[:4]), (int(text[-1])-1)*3+1, 1).isoformat()
    if re.fullmatch(r'\d{4}', text):
        return date(int(text), 1, 1).isoformat()
    return bound(text).isoformat()


def csv_rows(text, required, delimiter=','):
    reader = csv.DictReader(io.StringIO(text.lstrip('\ufeff')), delimiter=delimiter)
    if not reader.fieldnames or not set(required).issubset(reader.fieldnames):
        raise SchemaChangedError('Encabezados CSV ausentes o incompatibles.')
    rows = list(reader)
    if any(None in row for row in rows):
        raise SchemaChangedError('CSV con columnas sobrantes.')
    return rows


class SeriesProvider(DatasetProvider):
    parser_version = '2026-10-23.1'
    contract_version = '1'
    source_name = ''
    endpoint = ''
    frequencies = FREQUENCIES

    def __init__(self, spec):
        super().__init__(spec)
        self.transport = CurlChromeTransport(state_dir=self.storage.state_root, timeout=60)

    def validate_code(self, code):
        if not isinstance(code, str) or not re.fullmatch(r'[A-Za-z0-9_.^=:/?&+% -]{1,250}', code):
            raise InvalidQueryError('Código de serie inválido.')

    def single_request(self, *, codigo, frecuencia, anio):
        self.validate_code(codigo)
        if frecuencia not in self.frequencies or type(anio) is not int or not 1900 <= anio <= 2100:
            raise InvalidQueryError('Frecuencia o año fuera del dominio del proveedor.')
        digest = sha256((codigo+'|'+frecuencia).encode()).hexdigest()[:24]
        return PeriodRequest(f'{digest}:{anio}', f'serie={digest}/year={anio}',
                             {'codigo': codigo, 'frecuencia': frecuencia, 'anio': anio}, mutable=True)

    def plan_sync(self, *, codigo, frecuencia, desde, hasta=None):
        start, end = bound(desde), bound(hasta or date.today().isoformat(), True)
        if end < start or end.year-start.year > 100:
            raise InvalidQueryError('Rango invertido o mayor de 101 años.')
        for year in range(start.year, end.year+1):
            yield self.single_request(codigo=codigo, frecuencia=frecuencia, anio=year)

    def get(self, url, **kwargs):
        # Exception messages never include a credential-bearing URL or headers.
        try:
            response = self.transport.request('GET', url, ajax=False, allow_redirects=False, **kwargs)
        except WAFBlockedError:
            raise WAFBlockedError(f'{self.source_name}: bloqueo HTTP/WAF.') from None
        except Exception as exc:
            status = re.search(r'HTTP (\d{3})', str(exc))
            detail = 'HTTP '+status[1] if status else 'conexión o TLS'
            raise SourceUnavailableError(f'{self.source_name}: fallo {detail}.') from None
        if 300 <= response.status_code < 400:
            raise SourceUnavailableError(f'{self.source_name}: redirección no aceptada.')
        if len(response.content) > 20_000_000:
            raise SchemaChangedError('Respuesta de series mayor de 20 MB.')
        return response

    def json(self, url, **kwargs):
        response = self.get(url, **kwargs)
        try:
            return json.loads(response.content.decode('utf-8-sig'))
        except (ValueError, UnicodeError):
            raise SchemaChangedError(f'{self.source_name}: respuesta no JSON o ilegible.') from None

    def observations(self, code, frequency, year):
        raise NotImplementedError

    def _fetch_period(self, request):
        code, frequency, year = (request.params[k] for k in ('codigo', 'frecuencia', 'anio'))
        records, title, url = self.observations(code, frequency, year)
        rows = []
        for record in records:
            try:
                observed = date.fromisoformat(record['date'])
            except (ValueError, TypeError, KeyError):
                raise SchemaChangedError('Fecha de observación inválida.') from None
            if observed.year != year:
                continue
            value = record['value']
            if isinstance(value, bool) or (not pd.isna(value) and not math.isfinite(float(value))):
                raise SchemaChangedError('Observación numérica inválida.')
            rows.append(dict(series_code=code, title=title, period=record.get('period', observed.isoformat()),
                period_date=observed.isoformat(), frequency=frequency, value=value,
                source_value=record.get('source_value', ''), unit=record.get('unit', 'source_native'),
                source_unit_multiplier=str(record.get('source_unit_multiplier', '')),
                observation_status=str(record.get('status', '')),
                data_quality_flags=';'.join(filter(None, ['source_value_missing' if pd.isna(value) else '', record.get('flags', '')])),
                source=self.source_name, source_url=url, retrieved_at=utc_now_iso()))
        if not rows:
            raise PeriodUnavailableError('La fuente no devolvió observaciones para el año solicitado.')
        data = pd.DataFrame(rows)
        if data.period_date.duplicated().any():
            raise SchemaChangedError('Fechas duplicadas: la respuesta puede contener más de una serie.')
        for column in data:
            data[column] = data[column].astype('float64' if column == 'value' else 'string')
        data = data.sort_values('period_date').reset_index(drop=True)
        return FetchResult(self.spec.dataset_id, data, metadata={
            'frequency_basis': 'query_declaration', 'unit_basis': 'source_native_unless_source_identifies_unit',
            'scope': 'observations_returned_by_source_not_complete_history'},
            raw=data.to_json(orient='records', force_ascii=False))

    def partition_keys_for_load(self, *, codigo=None, frecuencia=None, desde=None, hasta=None):
        if codigo is None or frecuencia is None or desde is None or hasta is None:
            return None
        return [r.partition_key for r in self.plan_sync(codigo=codigo, frecuencia=frecuencia, desde=desde, hasta=hasta)]

    def filter_loaded(self, data, *, codigo=None, frecuencia=None, desde=None, hasta=None):
        if codigo is not None:
            data = data[data.series_code.eq(codigo)]
        if frecuencia is not None:
            data = data[data.frequency.eq(frecuencia)]
        if desde is not None:
            data = data[data.period_date.ge(bound(desde).isoformat())]
        if hasta is not None:
            data = data[data.period_date.le(bound(hasta, True).isoformat())]
        return data.drop(columns='_period_key', errors='ignore')
