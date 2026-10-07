"""Published capital requirements, risk-weighted assets and capital structure."""
from calendar import monthrange
from datetime import date, datetime
import math
import re

import pandas as pd

from fuentes_financieras.exceptions import SchemaChangedError
from ._excel_mensual import MonthlyExcelProvider, clean, month, read_workbook

RATIO_CODES = {'B': 'B-2402', 'F': 'B-3302', 'C': 'C-1252', 'R': 'C-2257'}
CAPITAL_CODES = {'B': 'B-2370', 'F': 'B-3252', 'C': 'C-1257', 'R': 'C-2262'}
RATIO_METRICS = {
    1: ('required_capital_credit', 'POR RIESGO DE CRÉDITO'),
    2: ('required_capital_market', 'POR RIESGO DE MERCADO'),
    3: ('required_capital_operational', 'POR RIESGO OPERACIONAL'),
    5: ('risk_weighted_assets_credit', 'POR RIESGO DE CRÉDITO'),
    6: ('risk_weighted_assets_market', 'POR RIESGO DE MERCADO'),
    7: ('risk_weighted_assets_operational', 'POR RIESGO OPERACIONAL'),
    8: ('risk_weighted_assets_total', 'ACTIVOS Y CONTINGENTES PONDERADOS POR RIESGO TOTALES (APR)'),
    10: ('common_equity_tier1_ratio', 'CAPITAL ORDINARIO DE NIVEL 1 / APR'),
    11: ('tier1_capital_ratio', 'PATRIMONIO EFECTIVO DE NIVEL 1 / APR'),
    12: ('global_capital_ratio', 'RATIO DE CAPITAL GLOBAL'),
}
CAPITAL_METRICS = {
    1: ('common_equity_tier1', 'Capital Ordinario de Nivel 1'),
    2: ('additional_tier1_capital', 'Capital Adicional de Nivel 1'),
    3: ('tier2_capital', 'Patrimonio Efectivo de Nivel 2'),
    4: ('effective_capital_total', 'Patrimonio Efectivo Total'),
}


def header(value):
    return re.sub(r'\d+/$', '', clean(value).replace('\\n', ' ')).strip()


def number(value):
    if pd.isna(value):
        return float('nan')
    if not isinstance(value, (int, float)) or isinstance(value, bool) or not math.isfinite(value):
        raise SchemaChangedError('Celda de solvencia no numérica.')
    return float(value)


def parse_table(content, *, entity_type, period, source_url, retrieved_at, kind):
    target = month(period)
    expected = date(target.year, target.month, monthrange(target.year, target.month)[1])
    book = read_workbook(content)
    candidates = [(name, frame) for name, frame in book.items() if not frame.empty and
        any('Patrimonio Efectivo' in clean(v) for v in frame.iloc[:4, 0])]
    if len(candidates) != 1:
        raise SchemaChangedError('Hoja de solvencia ausente o ambigua.')
    sheet, frame = candidates[0]
    if len(frame) < 12 or frame.shape[1] != (13 if kind == 'ratios' else 5):
        raise SchemaChangedError('Dimensiones del cuadro de solvencia no reconocidas.')
    entity_headers = [i for i, v in frame.iloc[:12, 0].items()
                      if clean(v) == ('EMPRESAS' if kind == 'ratios' else 'ENTIDAD')]
    if len(entity_headers) != 1:
        raise SchemaChangedError('Encabezado de entidades ausente o ambiguo.')
    entity_header = entity_headers[0]
    date_rows = [i for i, v in frame.iloc[:entity_header, 0].items()
                 if isinstance(v, (pd.Timestamp, datetime))]
    if not date_rows:
        raise SchemaChangedError('Fecha de solvencia ausente.')
    observed = frame.iloc[date_rows[0], 0]
    if not isinstance(observed, (pd.Timestamp, datetime)) or observed.date() != expected:
        raise SchemaChangedError('Fecha del cuadro distinta del período solicitado.')
    title = next((clean(v) for v in frame.iloc[:date_rows[0], 0] if 'Patrimonio Efectivo' in clean(v)), '')
    if kind == 'ratios' and title != 'Requerimiento de Patrimonio Efectivo y Ratio de Capital Global':
        raise SchemaChangedError('Título de solvencia no reconocido.')
    declared = clean(frame.iloc[date_rows[0] + 1, 0])
    if declared not in ('(En miles de soles)', '(En porcentaje)'):
        raise SchemaChangedError('Unidad del cuadro de solvencia no reconocida.')
    if kind == 'ratios' and declared != '(En miles de soles)':
        raise SchemaChangedError('Los importes de solvencia deben declarar miles de soles.')
    metrics = RATIO_METRICS if kind == 'ratios' else CAPITAL_METRICS
    label_rows = {}
    for col, (_, label) in metrics.items():
        matches = [i for i in range(entity_header + 2) if header(frame.iloc[i, col]) == label]
        if len(matches) != 1:
            raise SchemaChangedError(f'Encabezado ausente o ambiguo en columna {col + 1}.')
        label_rows[col] = matches[0]
    if kind == 'ratios':
        if 'REQUERIMIENTO DE PATRIMONIO EFECTIVO' != clean(frame.iloc[entity_header - 2, 1]) or 'PONDERADOS' not in clean(frame.iloc[entity_header - 2, 5]):
            raise SchemaChangedError('Grupos de requerimientos y APR no reconocidos.')
        if any(clean(frame.iloc[entity_header + 1, col]) != '(En porcentaje)' for col in (10, 11, 12)):
            raise SchemaChangedError('Ratios sin unidad porcentual explícita.')
    auxiliary = frame.iloc[date_rows[1], 0] if len(date_rows) > 1 else None
    auxiliary = auxiliary.date().isoformat() if isinstance(auxiliary, (pd.Timestamp, datetime)) else ''
    stop = next((i for i in range(entity_header + 3, len(frame)) if clean(frame.iloc[i, 0]).startswith('Fuente:')), None)
    if stop is None:
        raise SchemaChangedError('Falta el cierre del cuadro y su fuente.')
    notes = [{'sheet': str(sheet), 'row': i + 1, 'text': clean(frame.iloc[i, 0])}
             for i in range(stop, len(frame)) if clean(frame.iloc[i, 0])]
    rows, entities, warnings = [], set(), []
    for row in range(entity_header + 3, stop):
        name = clean(frame.iloc[row, 0])
        if not name:
            if any(pd.notna(frame.iloc[row, col]) for col in metrics):
                raise SchemaChangedError('Datos numéricos sin nombre de entidad.')
            continue
        if name in entities:
            raise SchemaChangedError('Entidad duplicada en el cuadro.')
        entities.add(name)
        values = {col: number(frame.iloc[row, col]) for col in metrics}
        if not math.isfinite(values[8 if kind == 'ratios' else 4]):
            raise SchemaChangedError(f'Total esencial ausente para {name}.')
        flags = []
        if kind == 'ratios':
            if any(not math.isfinite(values[c]) for c in (5, 6, 7, 10, 11, 12)):
                raise SchemaChangedError(f'APR o ratios incompletos para {name}.')
            if abs(sum(values[c] for c in (5, 6, 7)) - values[8]) > 0.02:
                raise SchemaChangedError(f'APR total no coincide con sus componentes para {name}.')
        elif all(math.isfinite(values[c]) for c in (1, 2, 3)):
            target_sum = 100 if declared == '(En porcentaje)' else values[4]
            if abs(sum(values[c] for c in (1, 2, 3)) - target_sum) > 0.02:
                # Preserve contradictory published data instead of correcting it.
                flags.append('published_components_sum_mismatch')
                warnings.append({'entity_name': name, 'warning': flags[-1]})
        for col, (metric, _) in metrics.items():
            unit = 'percent' if (kind == 'ratios' and col >= 10) or (kind == 'capital' and declared == '(En porcentaje)' and col < 4) else 'thousands_PEN'
            metric_flags = list(flags)
            if kind == 'capital' and declared == '(En porcentaje)' and col == 4:
                # The C/R header declares percentages, but its total column
                # contains amounts with no separate unit. Do not infer a scale.
                unit = 'unspecified_by_source'
                metric_flags.append('source_unit_unspecified')
            label_row = label_rows[col]
            rows.append(dict(period=period, period_date=expected.isoformat(), frequency='monthly',
                entity_type=entity_type, entity_name=name, entity_scope='system_aggregate' if name.upper().startswith('TOTAL ') else ('entity_with_foreign_branches' if 'SUCURSALES EN EL EXTERIOR' in name.upper() else 'entity'),
                table_kind=kind, metric=metric, source_metric=clean(frame.iloc[label_row, col]),
                value=values[col], unit=unit, unit_multiplier=1000.0 if unit == 'thousands_PEN' else float('nan'),
                data_quality_flags=';'.join(metric_flags), source_sheet=str(sheet), source_row=row+1,
                source_auxiliary_date=auxiliary, source='SBS', source_url=source_url, retrieved_at=retrieved_at))
    if not entities or not any(n.upper().startswith('TOTAL ') for n in entities):
        raise SchemaChangedError('Cuadro sin entidades o total del sistema.')
    data = pd.DataFrame(rows)
    for col in data:
        data[col] = data[col].astype('float64' if col in ('value', 'unit_multiplier') else ('int64' if col == 'source_row' else 'string'))
    if warnings:
        notes.append({'validation_warnings': warnings})
    return data, notes


def parse_ratios(content, **kwargs):
    return parse_table(content, kind='ratios', **kwargs)


def parse_capital(content, **kwargs):
    return parse_table(content, kind='capital', **kwargs)


class SolvencyProvider(MonthlyExcelProvider):
    codes = RATIO_CODES
    parser_version = '2026-10-07.1'
    parse_workbook = staticmethod(parse_ratios)


class EffectiveCapitalProvider(MonthlyExcelProvider):
    codes = CAPITAL_CODES
    parser_version = '2026-10-07.1'
    parse_workbook = staticmethod(parse_capital)
