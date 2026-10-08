"""Published deposit funding, system size bands and financial obligations."""
from datetime import datetime
import math

import pandas as pd

from fuentes_financieras.exceptions import SchemaChangedError
from ._excel_mensual import MonthlyExcelProvider, clean, read_workbook
from .calidad_cartera import closing, number, scope, source_notes, guard_percentage_formats

PERSON_CODES = dict(zip('BFCR', ('B-2372', 'B-3231', 'C-1245', 'C-2250')))
SCALE_CODES = dict(zip('BFCR', ('B-2321', 'B-3256', 'C-1211', 'C-2211')))
TERM_CODES = {'B': 'B-220513', 'F': 'B-3251'}
DEBT_CODES = dict(zip('BFCR', ('B-2310', 'B-3239', 'C-1219', 'C-2219')))
PERSON_LABELS = ('Personas Naturales', 'Personas Jurídicas sin fines de lucro', 'Otras Personas Jurídicas')
PERSON_IDS = ('natural', 'nonprofit_legal', 'other_legal')
DEPOSIT_LABELS = ('Depósitos a la Vista', 'Depósitos de Ahorros', 'Depósitos a Plazo', 'Depósitos CTS', 'Depósitos Totales')
DEPOSIT_IDS = ('sight', 'savings', 'term', 'cts', 'total')
SCALE_LABELS = ('Depósitos Vista', 'Depósitos de Ahorro', 'Depósitos a Plazo', 'Depósitos CTS', 'Depósitos Totales')
TERM_LABELS = ('Cuenta Corriente', 'Ahorros', 'Hasta 30 días', 'De 31 a 90 días', 'De 91 a 180 días', 'De 181 a 360 días', 'Más de 360 días', 'C.T.S.', 'Total')
TERM_BUCKETS = ('', '', 'le_30', '31_90', '91_180', '181_360', 'gt_360', '', '')


def compact(value):
    return clean(value).replace(' ', '').lower()


def table(content, prefix):
    matches = [(s, f) for s, f in read_workbook(content).items()
               if not f.empty and any(clean(v).startswith(prefix) for v in f.iloc[:4, 0])]
    if len(matches) != 1:
        raise SchemaChangedError('Hoja de fondeo ausente o ambigua.')
    return matches[0]


def header(frame, label='Empresas'):
    matches = [i for i, v in frame.iloc[:10, 0].items() if clean(v) == label]
    if len(matches) != 1:
        raise SchemaChangedError('Encabezado de fondeo ausente o ambiguo.')
    return matches[0]


def period_date(frame, period, head, *, col=0, allow_month_date=False):
    dates = [v.date() for v in frame.iloc[:head, col] if isinstance(v, (pd.Timestamp, datetime))]
    expected = closing(period)
    if len(dates) != 1 or (dates[0] != expected and not (allow_month_date and dates[0].strftime('%Y-%m') == period)):
        raise SchemaChangedError('Fecha del cuadro distinta del período solicitado.')
    return dates[0].isoformat()


def footer(frame, start):
    stop = next((i for i in range(start, len(frame))
                 if clean(frame.iloc[i, 0]).lower().startswith(('fuente:', 'nota:'))), None)
    if stop is None:
        raise SchemaChangedError('Cierre de fuente ausente.')
    return stop, source_notes(frame, stop)


def entities(frame, start, stop, columns):
    names = set()
    for row in range(start, stop):
        raw = frame.iloc[row, 0]; name = clean(raw)
        if not name:
            if any(not pd.isna(frame.iloc[row, col]) for col in columns):
                raise SchemaChangedError('Importes sin entidad.')
            continue
        placeholder = not isinstance(raw, str)
        if placeholder and raw != 0:
            raise SchemaChangedError('Identificador de entidad no reconocido.')
        if not placeholder:
            if name in names:
                raise SchemaChangedError('Entidad duplicada.')
            names.add(name)
        yield row, '' if placeholder else name, name, placeholder


def record(*, entity_type, period, sheet, row, col, name, metric, label, raw,
           source_url, retrieved_at, unit='thousands_PEN', currency='TOTAL', flags=(), **fields):
    value = number(raw)
    notices = list(flags) + (['source_value_missing'] if math.isnan(value) else [])
    return dict(period=period, period_date=closing(period).isoformat(), frequency='monthly',
        entity_type=entity_type, entity_name=name, entity_scope=scope(name), metric=metric,
        source_metric=label, value=value, source_value_token=clean(raw) if isinstance(raw, str) else '',
        unit=unit, unit_multiplier=1000.0 if unit.startswith('thousands_') else float('nan'),
        currency=currency, unit_evidence='table_header', data_quality_flags=';'.join(notices),
        source_sheet=str(sheet), source_row=row+1, source_column=col+1,
        source='SBS', source_url=source_url, retrieved_at=retrieved_at, **fields)


def finish(rows, notes):
    if not rows or not any(r['entity_scope'] == 'system_aggregate' for r in rows):
        raise SchemaChangedError('Cuadro sin observaciones o agregado publicado.')
    data = pd.DataFrame(rows)
    numeric = {'value', 'unit_multiplier', 'band_lower_PEN', 'band_upper_PEN'}
    for col in data:
        data[col] = data[col].astype('float64' if col in numeric else
                                    'int64' if col in ('source_row', 'source_column') else 'string')
    return data, notes


def mark_sum(rows, total, components, flag):
    # Never replace source markers with zero to validate a total.
    values = [rows[i]['value'] for i in [total, *components]]
    if all(math.isfinite(v) for v in values) and abs(values[0]-sum(values[1:])) > .02:
        for i in [total, *components]:
            rows[i]['data_quality_flags'] = ';'.join(filter(None, (rows[i]['data_quality_flags'], flag)))


def parse_person(content, *, entity_type, period, source_url, retrieved_at):
    sheet, frame = table(content, 'Depósitos por Tipo')
    head = header(frame); period_date(frame, period, head)
    if not any(compact(v) == '(enmilesdesoles)' for v in frame.iloc[:head, 0]):
        raise SchemaChangedError('Unidad de depósitos por persona cambiada.')
    labels = DEPOSIT_LABELS if entity_type in 'BF' else DEPOSIT_LABELS[1:]
    groups = [(j, clean(v)) for j, v in enumerate(frame.iloc[head]) if j and clean(v)]
    if [v for _, v in groups] != list(labels):
        raise SchemaChangedError('Tipos de depósitos cambiados.')
    columns = [j+k for j, _ in groups for k in range(3)]
    if frame.shape[1] != max(columns)+1:
        raise SchemaChangedError('Dimensiones de depósitos por persona cambiadas.')
    for col, _ in groups:
        if [clean(v) for v in frame.iloc[head+1, col:col+3]] != list(PERSON_LABELS):
            raise SchemaChangedError('Tipos de persona cambiados.')
    stop, notes = footer(frame, head+2); rows = []
    for row, name, original, placeholder in entities(frame, head+2, stop, columns):
        unused = set(range(1, frame.shape[1]))-set(columns)
        if any(not pd.isna(frame.iloc[row, j]) for j in unused):
            raise SchemaChangedError('Valores fuera de columnas de depósitos.')
        start = len(rows)
        for col, label in groups:
            for k, person in enumerate(PERSON_IDS):
                r = record(entity_type=entity_type, period=period, sheet=sheet, row=row, col=col+k,
                    name=name, metric='deposit_amount', label=f'{label} / {PERSON_LABELS[k]}',
                    raw=frame.iloc[row, col+k], source_url=source_url, retrieved_at=retrieved_at,
                    deposit_type=DEPOSIT_IDS[DEPOSIT_LABELS.index(label)], person_type=person,
                    source_entity_name=original, measurement_basis='published_deposit_balance',
                    flags=['source_entity_placeholder'] if placeholder else [])
                if placeholder:r['entity_scope'] = 'unidentified_source_row'
                rows.append(r)
        for k in range(3):
            mark_sum(rows, start+(len(groups)-1)*3+k,
                     [start+g*3+k for g in range(len(groups)-1)], 'published_components_mismatch')
    return finish(rows, notes)


def parse_debt(content, *, entity_type, period, source_url, retrieved_at):
    sheet, frame = table(content, 'Estructura de los Adeudos')
    head = header(frame); period_date(frame, period, head)
    if frame.shape[1] != 6 or not any(compact(v) == '(enporcentaje)' for v in frame.iloc[:head, 0]):
        raise SchemaChangedError('Dimensiones o unidad de adeudos cambiadas.')
    if clean(frame.iloc[head, 1]) != 'Instituciones del País' or clean(frame.iloc[head, 3]) != 'Instituciones del Exterior y Organismos Internacionales':
        raise SchemaChangedError('Origen de adeudos cambiado.')
    if [clean(v).lower() for v in frame.iloc[head+1, 1:5]] != ['corto plazo', 'largo plazo']*2:
        raise SchemaChangedError('Plazos de adeudos cambiados.')
    if not clean(frame.iloc[head, 5]).startswith('Total Adeudos y Obligaciones Financieras') or 'milesdesoles' not in compact(frame.iloc[head, 5])+compact(frame.iloc[head+1, 5]):
        raise SchemaChangedError('Unidad del total de adeudos ausente.')
    stop, notes = footer(frame, head+2); rows = []
    for row, name, original, placeholder in entities(frame, head+2, stop, range(1, 6)):
        for col in range(1, 6):
            flags = ['source_entity_placeholder'] if placeholder else []
            value = number(frame.iloc[row, col])
            if col < 5 and math.isfinite(value) and not 0 <= value <= 100:
                flags.append('published_percentage_out_of_range')
            r = record(entity_type=entity_type, period=period, sheet=sheet, row=row, col=col,
                name=name, metric='financial_obligations_total' if col == 5 else 'financial_obligations_share',
                label=clean(frame.iloc[head, 5]) if col == 5 else f'{clean(frame.iloc[head, 1 if col < 3 else 3])} / {clean(frame.iloc[head+1, col])}',
                raw=frame.iloc[row, col], source_url=source_url, retrieved_at=retrieved_at,
                unit='thousands_PEN' if col == 5 else 'percent',
                funding_origin='' if col == 5 else 'domestic' if col < 3 else 'foreign_and_international',
                term_bucket='' if col == 5 else 'short_term' if col in (1, 3) else 'long_term',
                source_entity_name=original, measurement_basis='published_financial_obligations', flags=flags)
            if placeholder:r['entity_scope'] = 'unidentified_source_row'
            rows.append(r)
        shares = [r['value'] for r in rows[-5:-1]]
        if all(math.isfinite(v) for v in shares) and math.isfinite(value) and value > 0 and abs(sum(shares)-100) > .02:
            for r in rows[-5:]:r['data_quality_flags'] += ';published_shares_sum_mismatch' if r['data_quality_flags'] else 'published_shares_sum_mismatch'
    guard_percentage_formats(content, rows)
    return finish(rows, notes)


def parse_terms(content, *, entity_type, period, source_url, retrieved_at):
    rows, notes, currencies = [], [], set()
    for sheet, frame in read_workbook(content).items():
        titles = [clean(v) for v in frame.iloc[:4, 0] if clean(v).startswith('Depósitos del público en Moneda')]
        if not titles:continue
        if len(titles) != 1 or frame.shape[1] != 10:
            raise SchemaChangedError('Cuadro de plazos ambiguo o dimensiones cambiadas.')
        currency = 'MN' if 'Moneda Nacional' in titles[0] else 'ME' if 'Moneda Extranjera' in titles[0] else None
        if currency is None or currency in currencies:raise SchemaChangedError('Moneda de depósitos ambigua.')
        currencies.add(currency)
        heads = [i for i in range(min(10, len(frame))) if clean(frame.iloc[i, 3]) == 'Cuentas a Plazo']
        if len(heads) != 1:raise SchemaChangedError('Encabezado de plazos ausente.')
        head = heads[0]; date_value = period_date(frame, period, head, allow_month_date=True)
        if [clean(v) for v in frame.iloc[head+1, 1:]] != list(TERM_LABELS):
            raise SchemaChangedError('Tramos de plazo cambiados.')
        unit = 'thousands_PEN' if currency == 'MN' else 'thousands_USD'
        expected = '(enmilesdesoles)' if currency == 'MN' else '(enmilesdedólares)'
        if not any(compact(v) == expected for v in frame.iloc[:head, 0]):
            raise SchemaChangedError('Unidad monetaria de plazos cambiada.')
        stop, footer_notes = footer(frame, head+2)
        notes.append({'sheet': str(sheet), 'source_date': date_value,
            'date_contract': 'Published date belongs to the requested month; no month-end date inferred.',
            'balance_definition': 'Report 6-B published balances; no daily average or closing-balance assumption.'})
        notes.extend(footer_notes)
        for row, name, original, placeholder in entities(frame, head+2, stop, range(1, 10)):
            start = len(rows)
            for col, label in enumerate(TERM_LABELS, start=1):
                flags = ['source_entity_placeholder'] if placeholder else []
                if date_value != closing(period).isoformat():flags.append('source_date_not_month_end')
                r = record(entity_type=entity_type, period=period, sheet=sheet, row=row, col=col,
                    name=name, metric='deposit_amount', label=label, raw=frame.iloc[row, col],
                    source_url=source_url, retrieved_at=retrieved_at, unit=unit, currency=currency,
                    deposit_type='sight' if col == 1 else 'savings' if col == 2 else 'cts' if col == 8 else 'total' if col == 9 else 'term',
                    term_bucket=TERM_BUCKETS[col-1], source_entity_name=original,
                    measurement_basis='report_6B_published_balance', flags=flags)
                r['period_date'] = date_value
                if placeholder:r['entity_scope'] = 'unidentified_source_row'
                rows.append(r)
            mark_sum(rows, start+8, list(range(start, start+8)), 'published_components_mismatch')
    if currencies != {'MN', 'ME'}:raise SchemaChangedError('Falta una moneda del cuadro de plazos.')
    return finish(rows, notes)


def parse_scale(content, *, entity_type, period, source_url, retrieved_at):
    sheet, frame = table(content, 'Depósitos de')
    head = header(frame, 'Escala'); period_date(frame, period, head, col=1 if entity_type in 'BF' else 0)
    if frame.shape[1] != 20 or clean(frame.iloc[head, 5]) != 'Personas Naturales' or clean(frame.iloc[head, 9]) != 'Personas Jurídicas' or clean(frame.iloc[head, 17]) != 'TOTAL':
        raise SchemaChangedError('Encabezado de escalas cambiado.')
    if clean(frame.iloc[head+1, 9]) != 'Privadas sin fines de lucro' or clean(frame.iloc[head+1, 13]).lower() != 'otras personas jurídicas':
        raise SchemaChangedError('Personas de escalas cambiadas.')
    if compact(frame.iloc[head+2, 0]) != '(ensoles)':raise SchemaChangedError('Unidad de límites de escala cambiada.')
    for col in (5, 9, 13, 17):
        if clean(frame.iloc[head+2, col]) != 'Número' or clean(frame.iloc[head+2, col+2]) != 'Monto':
            raise SchemaChangedError('Número o monto de escalas cambiado.')
        allowed = {'(milesdesoles)'} if entity_type in 'BF' else {'(enmiles)'}
        if compact(frame.iloc[head+3, col+2]) not in allowed:raise SchemaChangedError('Unidad de montos de escala cambiada.')
    stop, notes = footer(frame, head+4)
    sections = [(i, clean(frame.iloc[i, 0])) for i in range(head+4, stop) if clean(frame.iloc[i, 0]).startswith('Depósitos')]
    expected = SCALE_LABELS if entity_type in 'BF' else SCALE_LABELS[1:]
    if [label for _, label in sections] != list(expected):raise SchemaChangedError('Productos de escala cambiados.')
    name = next(clean(v) for v in frame.iloc[:4, 0] if clean(v).startswith('Depósitos de'))
    rows = []
    for pos, (start, label) in enumerate(sections):
        end = sections[pos+1][0] if pos+1 < len(sections) else stop
        bands = []; previous = None; open_ended = False
        for row in range(start+1, end):
            if not frame.iloc[row].notna().any():continue
            if open_ended:raise SchemaChangedError('Fila después del tramo abierto.')
            first = clean(frame.iloc[row, 1]) == 'Hasta'
            if first:
                if bands:raise SchemaChangedError('Primer tramo duplicado.')
                lower = float('nan')
            else:
                if clean(frame.iloc[row, 0]) != 'de' or clean(frame.iloc[row, 2]) != 'a' or not bands:
                    raise SchemaChangedError('Fila de escala desconocida.')
                lower = number(frame.iloc[row, 1])
                if not math.isfinite(lower) or lower != previous:raise SchemaChangedError('Escalas discontinuas.')
            open_ended = clean(frame.iloc[row, 3]) == 'más'
            upper = float('nan') if open_ended else number(frame.iloc[row, 3])
            if not open_ended and (not math.isfinite(upper) or upper <= (0 if first else lower)):
                raise SchemaChangedError('Límites de escala inválidos.')
            bands.append((row, lower, upper, 'upper_open' if open_ended else 'upper_bounded'))
            previous = upper
        if not bands or not open_ended:raise SchemaChangedError('Escala sin cierre abierto.')
        product_start = len(rows)
        for row, lower, upper, band_kind in [(start, float('nan'), float('nan'), 'product_total'), *bands]:
            if any(not pd.isna(frame.iloc[row, j]) for j in (4, 6, 8, 10, 12, 14, 16, 18)):
                raise SchemaChangedError('Valores fuera de columnas de escala.')
            for col, person in zip((5, 9, 13, 17), (*PERSON_IDS, 'total')):
                for offset, metric, unit in ((0, 'published_number', 'count'), (2, 'deposit_amount', 'thousands_PEN')):
                    value = number(frame.iloc[row, col+offset])
                    if metric == 'published_number' and math.isfinite(value) and (value < 0 or not value.is_integer()):
                        raise SchemaChangedError('Conteo de escala inválido.')
                    r = record(entity_type=entity_type, period=period, sheet=sheet, row=row, col=col+offset,
                        name=name, metric=metric, label=f'{label} / {person} / {clean(frame.iloc[head+2, col+offset])}',
                        raw=frame.iloc[row, col+offset], source_url=source_url, retrieved_at=retrieved_at, unit=unit,
                        deposit_type=DEPOSIT_IDS[SCALE_LABELS.index(label)], person_type=person,
                        band_kind=band_kind, band_lower_PEN=lower, band_upper_PEN=upper,
                        source_band_label=' '.join(clean(v) for v in frame.iloc[row, :4] if clean(v)),
                        boundary_convention='unspecified_by_source', measurement_basis='published_system_scale')
                    r['entity_scope'] = 'system_aggregate'; rows.append(r)
        for k in range(8):
            mark_sum(rows, product_start+k, [product_start+b*8+k for b in range(1, len(bands)+1)], 'published_bands_sum_mismatch')
    notes.append({'scope': 'System aggregate only; no individual entity concentration or top depositors.',
                  'number_semantics': 'Published Número; do not sum across products or infer unique system-wide depositors.',
                  'boundaries': 'Raw soles; interval inclusivity and first lower bound not specified by source.'})
    return finish(rows, notes)


class DepositsByPersonProvider(MonthlyExcelProvider):
    codes = PERSON_CODES
    parser_version = '2026-10-08.1'
    parse_workbook = staticmethod(parse_person)


class DepositSizeBandsProvider(MonthlyExcelProvider):
    codes = SCALE_CODES
    parser_version = '2026-10-08.1'
    parse_workbook = staticmethod(parse_scale)


class DepositsByTermProvider(MonthlyExcelProvider):
    codes = TERM_CODES
    parser_version = '2026-10-08.1'
    parse_workbook = staticmethod(parse_terms)

    def plan_sync(self, *, tipos=('B', 'F'), **query):
        return super().plan_sync(tipos=tipos, **query)


class FinancialObligationsProvider(MonthlyExcelProvider):
    codes = DEBT_CODES
    parser_version = '2026-10-08.1'
    parse_workbook = staticmethod(parse_debt)
