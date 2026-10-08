"""Published credit quality ratios, debtor categories and arrears thresholds."""
from calendar import monthrange
from datetime import date, datetime
import math
import re

import pandas as pd

from fuentes_financieras.exceptions import SchemaChangedError
from ._excel_mensual import MonthlyExcelProvider, clean, month, read_workbook, cell_formats

QUALITY_CODES = dict(zip('BFCR', ('B-2401', 'B-3301', 'C-1301', 'C-2301')))
CATEGORY_CODES = dict(zip('BFCR', ('B-2309', 'B-3205', 'C-120201', 'C-220201')))
ARREARS_CODES = dict(zip('BFCR', ('B-220512', 'B-3230', 'C-1230', 'C-2230')))
GLOSSARY_URL = 'https://intranet2.sbs.gob.pe/estadistica/financiera/2025/Enero/SF-0002-en2025.PDF'
QUALITY_METRICS = {
    'Créditos Atrasados (criterio SBS) / Créditos Directos': ('past_due_ratio_sbs', 'TOTAL'),
    'Créditos Atrasados con más de 90 días de atraso / Créditos Directos': ('past_due_over_90_days_ratio', 'TOTAL'),
    'Créditos Refinanciados y Reestructurados / Créditos Directos': ('refinanced_restructured_ratio', 'TOTAL'),
    'Créditos Atrasados MN (criterio SBS) / Créditos Directos MN': ('past_due_ratio_sbs', 'MN'),
    'Créditos Atrasados ME (criterio SBS) / Créditos Directos ME': ('past_due_ratio_sbs', 'ME'),
    'Provisiones / Créditos Atrasados': ('provisions_past_due_coverage', 'TOTAL'),
    'Cartera de Alto Riesgo / Créditos Directos': ('high_risk_credit_ratio', 'TOTAL'),
    'Cartera Atrasada Ajustada': ('adjusted_past_due_ratio', 'TOTAL'),
    'Cartera de Alto Riesgo Ajustada': ('adjusted_high_risk_ratio', 'TOTAL'),
}
EXPECTED_QUALITY = {
    'B': set(QUALITY_METRICS) - {'Cartera de Alto Riesgo / Créditos Directos'},
    'F': {label for label in QUALITY_METRICS if ' MN ' not in label and ' ME ' not in label} - {'Cartera de Alto Riesgo / Créditos Directos'},
    'C': set(QUALITY_METRICS) - {'Créditos Refinanciados y Reestructurados / Créditos Directos'},
    'R': set(QUALITY_METRICS) - {'Créditos Refinanciados y Reestructurados / Créditos Directos'},
}
CATEGORIES = ('Normal (0)', 'Con Problemas Potenciales (1)', 'Deficiente (2)', 'Dudoso (3)', 'Pérdida (4)')
CATEGORY_IDS = ('normal', 'potential_problems', 'substandard', 'doubtful', 'loss')


def normalized(value):
    return clean(re.sub(r'\*+|\(\s*%\s*\)', '', clean(value)))


def number(value):
    if pd.isna(value) or (isinstance(value, str) and value.strip() == '-'):
        return float('nan')
    if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
        raise SchemaChangedError('Valor de cartera no numérico.')
    return float(value)


def closing(period):
    target = month(period)
    return date(target.year, target.month, monthrange(target.year, target.month)[1])


def frame_for(content, prefix):
    candidates = [(s, d) for s, d in read_workbook(content).items() if not d.empty and
                  any(clean(v).lower().startswith(prefix.lower()) for v in d.iloc[:4, 0])]
    if len(candidates) != 1:
        raise SchemaChangedError('Hoja de cartera ausente o ambigua.')
    return candidates[0]


def validate_date(frame, period, before, col=0):
    dates = [v.date() for v in frame.iloc[:before, col] if isinstance(v, (pd.Timestamp, datetime))]
    if dates != [closing(period)]:
        raise SchemaChangedError('Fecha del cuadro de cartera distinta del mes solicitado.')


def source_notes(frame, start):
    return [{'row': i+1, 'column': j+1, 'text': clean(value)}
            for i in range(start, len(frame)) for j, value in enumerate(frame.iloc[i])
            if isinstance(value, str) and clean(value)]


def scope(name):
    if name.upper().startswith('TOTAL '):
        return 'system_aggregate'
    return 'entity_with_foreign_branches' if 'SUCURSALES EN EL EXTERIOR' in name.upper() else 'entity'


def record(*, entity_type, period, name, sheet, row, col, metric, label, value,
           source_url, retrieved_at, unit='percent', currency='TOTAL', credit_scope='direct_credit',
           category='', threshold=None, flags=(), unit_evidence='table_header', source_value_token=''):
    notices = list(flags) + (['source_value_missing'] if math.isnan(value) else [])
    return dict(period=period, period_date=closing(period).isoformat(), frequency='monthly',
        entity_type=entity_type, entity_name=name, entity_scope=scope(name), metric=metric,
        source_metric=label, value=value, source_value_token=source_value_token, unit=unit,
        unit_multiplier=1000.0 if unit == 'thousands_PEN' else float('nan'),
        currency=currency, credit_scope=credit_scope, risk_category=category,
        arrears_threshold_days=float(threshold) if threshold is not None else float('nan'),
        unit_evidence=unit_evidence, data_quality_flags=';'.join(notices),
        source_sheet=str(sheet), source_row=row+1, source_column=col+1,
        source='SBS', source_url=source_url, retrieved_at=retrieved_at)


def guard_percentage_formats(content, rows):
    positions = {(r['source_sheet'], r['source_row']-1, r['source_column']-1)
                 for r in rows if r['unit'] == 'percent'}
    if any('%' in style for style in cell_formats(content, positions).values()):
        raise SchemaChangedError('Formato porcentual Excel inesperado: requiere revisar la escala.')


def finish(rows, notes):
    if not rows or not any(r['entity_scope'] == 'system_aggregate' for r in rows):
        raise SchemaChangedError('Cuadro de cartera sin observaciones o agregado.')
    data = pd.DataFrame(rows)
    for col in data:
        data[col] = data[col].astype('float64' if col in ('value', 'unit_multiplier', 'arrears_threshold_days')
            else ('int64' if col in ('source_row', 'source_column') else 'string'))
    return data, notes


def parse_quality(content, *, entity_type, period, source_url, retrieved_at):
    sheet, frame = frame_for(content, 'Indicadores Financieros por')
    titles = [(i, j) for i in range(min(4, len(frame))) for j, v in enumerate(frame.iloc[i])
              if clean(v).startswith('Indicadores Financieros por')]
    label_cols = sorted({j for _, j in titles})
    sections = [i for i, v in frame.iloc[:, 0].items() if normalized(v) == 'CALIDAD DE ACTIVOS']
    if len(sections) != 1:
        raise SchemaChangedError('Sección de calidad de activos ausente o ambigua.')
    start = sections[0]
    stop = next((i for i in range(start+1, len(frame)) if clean(frame.iloc[i, 0]) == 'EFICIENCIA Y GESTIÓN'), None)
    if stop is None:
        raise SchemaChangedError('Cierre de calidad de activos ausente.')
    validate_date(frame, period, start)
    headers = [i for i in range(start) if not clean(frame.iloc[i, 0]) and
               isinstance(frame.iloc[i, 1], str) and clean(frame.iloc[i, 1])]
    if len(headers) != 1:
        raise SchemaChangedError('Encabezado de entidades de calidad ausente o ambiguo.')
    head = headers[0]
    notes_start = next((i for i in range(stop, len(frame)) if any(clean(v).startswith('Nota:') for v in frame.iloc[i])), None)
    if notes_start is None:
        raise SchemaChangedError('Notas de indicadores financieros ausentes.')
    notes = source_notes(frame, notes_start)
    declared_units = [clean(v) for v in frame.iloc[:head, 0]
                      if re.match(r'^\(\s*en\b', clean(v), re.I)]
    if any(v.replace(' ', '').lower() != '(enporcentaje)' for v in declared_units):
        raise SchemaChangedError('Unidad declarada de calidad no reconocida.')
    rows, names = [], set()
    for pos, label_col in enumerate(label_cols):
        if normalized(frame.iloc[start, label_col]) != 'CALIDAD DE ACTIVOS':
            raise SchemaChangedError('Bloques repetidos de indicadores desalineados.')
        validate_date(frame, period, head, label_col)
        metric_rows = {normalized(frame.iloc[i, label_col]): i for i in range(start+1, stop) if clean(frame.iloc[i, label_col])}
        if len(metric_rows) != sum(bool(clean(frame.iloc[i, label_col])) for i in range(start+1, stop)) or set(metric_rows) != EXPECTED_QUALITY[entity_type]:
            raise SchemaChangedError('Indicadores de calidad ausentes, duplicados o desconocidos.')
        end_col = label_cols[pos+1] if pos+1 < len(label_cols) else frame.shape[1]
        for col in range(label_col+1, end_col):
            name = clean(frame.iloc[head, col])
            if not name:
                if frame.iloc[list(metric_rows.values()), col].notna().any():
                    raise SchemaChangedError('Valores de calidad sin encabezado de entidad.')
                continue
            if name in names:
                raise SchemaChangedError('Entidad de calidad duplicada.')
            names.add(name)
            for label, row in metric_rows.items():
                metric, currency = QUALITY_METRICS[label]
                raw_label = clean(frame.iloc[row, label_col])
                flags = []
                if metric.startswith('adjusted_'):
                    marker = re.search(r'\*+$', raw_label)
                    if marker and not any(re.match(re.escape(marker[0])+r'(?!\*)', n['text']) for n in notes):
                        flags.append('source_definition_missing')
                evidence = 'table_header' if entity_type in 'BF' else 'reviewed_indicator_definition'
                rows.append(record(entity_type=entity_type, period=period, name=name, sheet=sheet,
                    row=row, col=col, metric=metric, label=raw_label, value=number(frame.iloc[row, col]),
                    source_url=source_url, retrieved_at=retrieved_at, currency=currency,
                    flags=flags, unit_evidence=evidence,
                    source_value_token=clean(frame.iloc[row, col]) if isinstance(frame.iloc[row, col], str) else ''))
    # This explicit dictionary supplies the percentage unit for C/R labels
    # that omit it; the official glossary defines the unadjusted indicators.
    notes.append({'reviewed_unit_contract': 'quality indicators expressed as percentage points',
                  'glossary_url': GLOSSARY_URL})
    if entity_type in 'BF' and not any('porcentaje' in clean(v).lower() for v in frame.iloc[:head, 0]):
        raise SchemaChangedError('Unidad porcentual del cuadro B/F ausente.')
    guard_percentage_formats(content, rows)
    return finish(rows, notes)


def parse_vertical(content, *, entity_type, period, source_url, retrieved_at, kind):
    prefix = 'Estructura de Créditos Directos' if kind == 'categories' else 'Ratios de Morosidad según días'
    sheet, frame = frame_for(content, prefix)
    if frame.shape[1] != (7 if kind == 'categories' else 6):
        raise SchemaChangedError('Dimensiones del cuadro de cartera cambiadas.')
    heads = [i for i, v in frame.iloc[:10, 0].items() if clean(v) == 'Empresas']
    if len(heads) != 1:
        raise SchemaChangedError('Encabezado de empresas ausente o ambiguo.')
    head = heads[0]; validate_date(frame, period, head)
    if kind == 'categories':
        if [clean(v) for v in frame.iloc[head, 1:6]] != list(CATEGORIES):
            raise SchemaChangedError('Categorías de riesgo cambiadas.')
        if not any(clean(v) == '(En porcentaje)' for v in frame.iloc[:head, 0]):
            raise SchemaChangedError('Unidad porcentual de categorías ausente.')
        credit_scope = 'direct_and_credit_equivalent_indirect' if entity_type in 'BF' else 'direct_and_credit_equivalent_contingent'
        total_label = clean(frame.iloc[head, 6])
        expected_total = 'Total Créditos Directos e Indirectos' if entity_type in 'BF' else 'Total Créditos Directos y Contingentes'
        if not total_label.startswith(expected_total) or 'miles de soles' not in total_label:
            raise SchemaChangedError('Cobertura o unidad del total de créditos cambiada.')
    else:
        if clean(frame.iloc[head, 1]) != 'Porcentaje de créditos con' or not clean(frame.iloc[head, 5]).startswith('Morosidad según criterio contable SBS'):
            raise SchemaChangedError('Unidad o criterio de morosidad cambiado.')
        for col, days in enumerate((30, 60, 90, 120), start=1):
            label = re.sub(r'\*+|\d+/$', '', clean(frame.iloc[head+1, col])).strip()
            if label != f'Más de {days} días de incumplimiento':
                raise SchemaChangedError('Umbral de morosidad cambiado.')
        credit_scope = 'direct_credit'
    stop = next((i for i in range(head+1, len(frame)) if clean(frame.iloc[i, 0]).lower().startswith(('fuente:', 'nota:'))), None)
    if stop is None:
        raise SchemaChangedError('Falta el cierre de fuente del cuadro de cartera.')
    notes = source_notes(frame, stop); rows, names = [], set()
    for row in range(head+1 if kind == 'categories' else head+2, stop):
        name = clean(frame.iloc[row, 0])
        if not name:
            if frame.iloc[row, 1:].notna().any():
                raise SchemaChangedError('Valores de cartera sin entidad.')
            continue
        if name in names:
            raise SchemaChangedError('Entidad de cartera duplicada.')
        names.add(name); values = [number(v) for v in frame.iloc[row, 1:]]; flags = []
        shares = values[:5] if kind == 'categories' else values[:4]
        if any(math.isfinite(v) and not 0 <= v <= 100 for v in shares):
            flags.append('published_percentage_out_of_range')
        if all(math.isfinite(v) for v in shares):
            if kind == 'categories' and abs(sum(shares)-100) > .02:
                flags.append('published_categories_sum_mismatch')
            if kind == 'arrears' and any(a+.02 < b for a,b in zip(shares, shares[1:])):
                flags.append('published_arrears_order_mismatch')
        for col, value in enumerate(values, start=1):
            category = CATEGORY_IDS[col-1] if kind == 'categories' and col <= 5 else ''
            threshold = (30, 60, 90, 120)[col-1] if kind == 'arrears' and col <= 4 else None
            metric = ('risk_category_share' if col <= 5 else 'total_classified_credit') if kind == 'categories' else ('arrears_over_days_ratio' if col <= 4 else 'past_due_ratio_sbs')
            label = clean(frame.iloc[head if kind == 'categories' or col == 5 else head+1, col])
            rows.append(record(entity_type=entity_type, period=period, name=name, sheet=sheet,
                row=row, col=col, metric=metric, label=label, value=value, source_url=source_url,
                retrieved_at=retrieved_at, unit='thousands_PEN' if kind == 'categories' and col == 6 else 'percent',
                credit_scope=credit_scope, category=category, threshold=threshold, flags=flags,
                source_value_token=clean(frame.iloc[row, col]) if isinstance(frame.iloc[row, col], str) else ''))
    guard_percentage_formats(content, rows)
    return finish(rows, notes)


def parse_categories(content, **kwargs):
    return parse_vertical(content, kind='categories', **kwargs)


def parse_arrears(content, **kwargs):
    return parse_vertical(content, kind='arrears', **kwargs)


class CreditQualityProvider(MonthlyExcelProvider):
    codes = QUALITY_CODES
    parser_version = '2026-10-08.1'
    parse_workbook = staticmethod(parse_quality)


class CreditRiskCategoriesProvider(MonthlyExcelProvider):
    codes = CATEGORY_CODES
    parser_version = '2026-10-08.1'
    parse_workbook = staticmethod(parse_categories)


class CreditArrearsProvider(MonthlyExcelProvider):
    codes = ARREARS_CODES
    parser_version = '2026-10-08.1'
    parse_workbook = staticmethod(parse_arrears)
