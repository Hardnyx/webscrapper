"""Monthly statistical statements, preserving source accounts and coverage."""
from calendar import monthrange
from datetime import date, datetime
from io import BytesIO
import math
import re
from urllib.parse import urljoin, urlparse

from bs4 import BeautifulSoup
import pandas as pd

from fuentes_financieras.exceptions import InvalidQueryError, PeriodUnavailableError, SchemaChangedError
from fuentes_financieras.models import FetchResult, PeriodRequest
from fuentes_financieras.provider import DatasetProvider
from fuentes_financieras.storage import utc_now_iso
from fuentes_financieras.transports.curl_chrome import CurlChromeTransport

INDEX = 'https://www.sbs.gob.pe/app/stats_net/stats/EstadisticaSistemaFinancieroResultados.aspx?c='
CODES = {'B': 'B-2201', 'F': 'B-3101', 'C': 'C-1101', 'R': 'C-2101'}
MONTHS = dict(zip(('en', 'fe', 'ma', 'ab', 'my', 'jn', 'jl', 'ag', 'se', 'oc', 'no', 'di'), range(1, 13)))
ACCOUNTS = {
    ('assets', 'TOTAL ACTIVO'): 'total_assets',
    ('liabilities_equity', 'TOTAL PASIVO'): 'total_liabilities',
    ('liabilities_equity', 'PATRIMONIO'): 'equity',
    ('liabilities_equity', 'TOTAL PASIVO Y PATRIMONIO'): 'total_liabilities_equity',
    ('liabilities_equity', 'RESULTADO NETO DEL EJERCICIO'): 'net_income',
    ('income', 'RESULTADO NETO DEL EJERCICIO'): 'net_income',
    ('income', 'INGRESOS FINANCIEROS'): 'financial_income',
    ('income', 'GASTOS FINANCIEROS'): 'financial_expenses',
}


def clean(value):
    return '' if pd.isna(value) else ' '.join(str(value).split())


def month(value):
    if not isinstance(value, str) or not re.fullmatch(r'\d{4}-\d{2}', value):
        raise InvalidQueryError('Use períodos YYYY-MM.')
    try:
        target = date.fromisoformat(value + '-01')
    except ValueError as exc:
        raise InvalidQueryError('Mes inválido.') from exc
    if target < date(2013, 1, 1) or target > date.today():
        raise InvalidQueryError('Período admitido: enero de 2013 hasta el mes actual.')
    return target


def discover_files(html, entity_type):
    """Use published links, never synthesize download URLs."""
    found = {}
    for anchor in BeautifulSoup(html, 'html.parser').find_all('a', href=True):
        url = urljoin(INDEX + CODES[entity_type], anchor['href'])
        parsed = urlparse(url)
        match = re.search(r'/' + CODES[entity_type] + r'-(\w{2})(\d{4})\.xlsx?$', parsed.path, re.I)
        if not match:
            continue
        if parsed.scheme != 'https' or parsed.hostname != 'intranet2.sbs.gob.pe':
            raise SchemaChangedError('Host de descarga SBS inesperado.')
        token, year = match.groups()
        if token.lower() not in MONTHS:
            raise SchemaChangedError('Mes del enlace SBS desconocido.')
        period = f'{year}-{MONTHS[token.lower()]:02d}'
        if period in found and found[period] != url:
            raise SchemaChangedError(f'Enlaces contradictorios para {period}.')
        found[period] = url
    if not found:
        raise SchemaChangedError('El índice no contiene enlaces de estados financieros.')
    return found


def parse_workbook(content, *, entity_type, period, source_url, retrieved_at):
    target = month(period)
    expected_date = date(target.year, target.month, monthrange(target.year, target.month)[1])
    if content.startswith(b'PK\x03\x04'):
        engine = 'openpyxl'
    elif content.startswith(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'):
        engine = 'xlrd'
    else:
        raise SchemaChangedError('La descarga no es un libro Excel XLS/XLSX.')
    try:
        sheets = pd.read_excel(BytesIO(content), sheet_name=None, header=None, engine=engine)
    except Exception as exc:
        raise SchemaChangedError('Libro Excel ilegible.') from exc
    rows, notes, seen = [], [], set()
    for sheet, frame in sheets.items():
        if frame.empty:
            continue
        titles = [(i, clean(v)) for i, v in frame.iloc[:, 0].items()
                  if clean(v).startswith(('Balance General', 'Estado de Ganancias'))]
        if not titles:
            continue
        for block, (title_row, title) in enumerate(titles):
            stop = titles[block + 1][0] if block + 1 < len(titles) else len(frame)
            part = frame.iloc[title_row:stop]
            dates = [v.date() for v in part.iloc[:5, 0] if isinstance(v, (pd.Timestamp, datetime))]
            if dates != [expected_date]:
                raise SchemaChangedError('La fecha del libro no coincide con el período solicitado.')
            if not any(re.fullmatch(r'\(En\s+miles\s+de\s+soles\)', clean(v), re.I) for v in part.iloc[:5, 0]):
                raise SchemaChangedError('Unidad del libro no reconocida: se esperan miles de soles.')
            headers = [i for i in range(title_row + 1, min(stop, title_row + 10))
                       if 'MN' in [clean(v) for v in frame.iloc[i]]]
            if len(headers) != 1:
                raise SchemaChangedError('Encabezado de monedas ausente o ambiguo.')
            header = headers[0]
            section = 'income' if title.startswith('Estado') else ('assets' if clean(frame.iloc[header - 1, 0]) == 'Activo' else 'liabilities_equity')
            if section == 'liabilities_equity' and clean(frame.iloc[header - 1, 0]) != 'Pasivo':
                raise SchemaChangedError('Sección de balance no reconocida.')
            seen.add(section)
            groups = []
            for col in range(1, frame.shape[1]):
                if clean(frame.iloc[header, col]) != 'MN':
                    continue
                if col + 2 >= frame.shape[1] or [clean(v) for v in frame.iloc[header, col:col+3]] != ['MN', 'ME', 'TOTAL']:
                    raise SchemaChangedError('Grupo de monedas incompleto.')
                raw_name = clean(frame.iloc[header - 1, col])
                if not raw_name:
                    raise SchemaChangedError('Nombre de entidad ausente.')
                name = re.sub(r'\*+$', '', raw_name).strip()
                # The August 2026 B workbook uses both labels at column 5.
                # Keep the original header in source_entity_name.
                if entity_type == 'B' and period == '2026-08' and name == 'BANCOM':
                    name = 'Banco de Comercio'
                scope = 'system_aggregate' if name.upper().startswith('TOTAL ') else ('entity_with_foreign_branches' if 'SUCURSALES EN EL EXTERIOR' in name.upper() else 'entity')
                groups.append((col, name, raw_name, scope))
            if not groups:
                raise SchemaChangedError('No hay entidades en el encabezado.')
            account_end = next((i for i in range(header + 1, stop)
                if clean(frame.iloc[i, 0]).startswith('Tipo de Cambio Contable')), stop)
            for row in range(header + 1, stop):
                label_raw = frame.iloc[row, 0]
                label = clean(label_raw)
                if not label or row >= account_end:
                    if label and not label.startswith(('Activo', 'Pasivo')):
                        notes.append({'sheet': sheet, 'row': row + 1, 'text': label})
                    continue
                for col, name, raw_name, scope in groups:
                    for offset, currency in enumerate(('MN', 'ME', 'TOTAL')):
                        value = frame.iloc[row, col + offset]
                        if pd.isna(value):
                            amount = float('nan')
                        elif isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value):
                            amount = float(value)
                        else:
                            raise SchemaChangedError(f'Importe no numérico en {sheet}:{row+1}, {name}.')
                        rows.append(dict(period=period, period_date=expected_date.isoformat(), frequency='monthly',
                            entity_type=entity_type, entity_name=name, source_entity_name=raw_name, entity_scope=scope,
                            statement='income' if section == 'income' else 'balance', section=section,
                            account=label, source_account=str(label_raw), account_code=ACCOUNTS.get((section, label.upper()), ''),
                            source_sheet=str(sheet), source_row=row+1, currency=currency, amount=amount,
                            unit='thousands_PEN', unit_multiplier=1000,
                            measurement_basis='year_to_date' if section == 'income' else 'closing_balance',
                            source='SBS', source_url=source_url, retrieved_at=retrieved_at))
    if seen != {'assets', 'liabilities_equity', 'income'}:
        raise SchemaChangedError('Libro incompleto: faltan activo, pasivo/patrimonio o resultados.')
    data = pd.DataFrame(rows)
    for col in data:
        data[col] = data[col].astype('float64' if col == 'amount' else ('int64' if col in ('source_row', 'unit_multiplier') else 'string'))
    validate_totals(data)
    return data, notes


def validate_totals(data):
    """Reject truncated or inconsistent statements; allow source rounding."""
    total = data[(data.currency == 'TOTAL') & data.account_code.isin(('total_assets', 'total_liabilities', 'equity', 'total_liabilities_equity', 'net_income'))]
    for name, group in total.groupby('entity_name'):
        if group.duplicated(['section', 'account_code']).any():
            raise SchemaChangedError(f'Totales duplicados para {name}.')
        amounts = {(r.section, r.account_code): r.amount for r in group.itertuples()}
        required = [('assets', 'total_assets')] + [('liabilities_equity', c) for c in ('total_liabilities', 'equity', 'total_liabilities_equity', 'net_income')] + [('income', 'net_income')]
        if any(k not in amounts or not math.isfinite(amounts[k]) for k in required):
            raise SchemaChangedError(f'Totales incompletos para {name}.')
        a, l, e, le, ni, income = [amounts[k] for k in required]
        if any(abs(x-y) > 0.02 for x, y in ((a, le), (l+e, le), (ni, income))):
            raise SchemaChangedError(f'Identidades contables inconsistentes para {name}.')
    if total.empty or set(total.entity_name) != set(data.entity_name):
        raise SchemaChangedError('Faltan totales de alguna entidad.')


class FinancialStatementsProvider(DatasetProvider):
    parser_version = '2026-10-07.1'

    def __init__(self, spec):
        super().__init__(spec)
        self.transport = CurlChromeTransport(state_dir=self.storage.state_root, timeout=30)
        self.indices = {}

    def single_request(self, *, periodo=None, tipo='B', **query):
        if query or tipo not in CODES:
            raise InvalidQueryError('Use periodo=YYYY-MM y tipo=B/F/C/R.')
        target = month(periodo)
        return PeriodRequest(f'{tipo}:{periodo}', f'tipo={tipo}/year={target.year}', {'periodo': periodo, 'tipo': tipo}, mutable=True)

    def plan_sync(self, *, desde, hasta=None, tipos=('B', 'F', 'C', 'R'), **query):
        start, end = month(desde), month(hasta or desde)
        selected = [tipos] if isinstance(tipos, str) else list(tipos)
        if query or not selected or len(set(selected)) != len(selected) or set(selected) - CODES.keys() or end < start:
            raise InvalidQueryError('Rango o tipos inválidos.')
        while start <= end:
            for tipo in selected:
                yield self.single_request(periodo=start.strftime('%Y-%m'), tipo=tipo)
            start = date(start.year + (start.month == 12), start.month % 12 + 1, 1)

    def _fetch_period(self, request):
        tipo, period = request.params['tipo'], request.params['periodo']
        if tipo not in self.indices:
            self.indices[tipo] = discover_files(self.transport.request('GET', INDEX + CODES[tipo]).text, tipo)
        if period not in self.indices[tipo]:
            raise PeriodUnavailableError(f'SBS no enlaza estados de {tipo} para {period}.')
        url = self.indices[tipo][period]
        content = self.transport.request('GET', url).content
        data, notes = parse_workbook(content, entity_type=tipo, period=period, source_url=url, retrieved_at=utc_now_iso())
        return FetchResult(self.spec.dataset_id, data, {'period': period, 'entity_type': tipo, 'source_url': url, 'notes': notes}, content)

    def sync(self, **query):
        self.indices = {}
        return super().sync(**query)

    def filter_loaded(self, data, *, desde=None, hasta=None, tipos=None):
        for value, upper in ((desde, False), (hasta, True)):
            if value is not None:
                month(value)
                data = data[data.period <= value] if upper else data[data.period >= value]
        if tipos is not None:
            selected = [tipos] if isinstance(tipos, str) else list(tipos)
            if set(selected) - CODES.keys():
                raise InvalidQueryError('Tipos admitidos: B/F/C/R.')
            data = data[data.entity_type.isin(selected)]
        return data.drop(columns=['_period_key'], errors='ignore')
