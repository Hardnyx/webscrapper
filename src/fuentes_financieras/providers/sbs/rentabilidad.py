"""Published profitability and efficiency indicators with explicit measurement bases."""
from datetime import date
import math

import pandas as pd

from fuentes_financieras.exceptions import SchemaChangedError
from ._excel_mensual import MonthlyExcelProvider, clean
from .calidad_cartera import (
    QUALITY_CODES, closing, frame_for, validate_date, source_notes, normalized,
    number, scope, guard_percentage_formats,
)

ANNUAL_NOTE = ('Los valores anualizados se obtienen de la siguiente manera: valor del mes + valor a diciembre del año anterior '
               '- valor del mismo mes del año anterior. El promedio corresponde a los últimos doce meses.')
# Explicit contracts distinguish the denominators and periods in the source labels.
PROFIT_METRICS = {
    'B': {
        'Utilidad Neta Anualizada / Patrimonio Promedio': ('return_on_equity', 'net_profit', 'average_equity'),
        'Utilidad Neta Anualizada / Activo Promedio': ('return_on_assets', 'net_profit', 'average_assets'),
    },
    'F': {
        'Utilidad Anualizada / Patrimonio Promedio': ('return_on_equity', 'profit_as_labeled', 'average_equity'),
        'Utilidad Anualizada / Activo Promedio': ('return_on_assets', 'profit_as_labeled', 'average_assets'),
    },
    'C': {
        'Utilidad Neta Anualizada sobre Patrimonio Promedio': ('return_on_equity', 'net_profit', 'average_equity'),
        'Utilidad Neta Anualizada sobre Activo Promedio': ('return_on_assets', 'net_profit', 'average_assets'),
    },
}
PROFIT_METRICS['R'] = PROFIT_METRICS['C']

BF_EFFICIENCY = {
    'Gastos de Administración Anualizados / Activo Productivo Promedio':
        ('administrative_expenses_productive_assets_ratio', 'percent', 'rolling_12_months', 'annualized_administrative_expenses', 'average_productive_assets'),
    'Gastos de Operación / Margen Financiero Total':
        ('operating_expenses_financial_margin_ratio', 'percent', 'unspecified_by_source', 'operating_expenses', 'total_financial_margin'),
    'Ingresos Financieros / Ingresos Totales':
        ('financial_income_total_income_ratio', 'percent', 'unspecified_by_source', 'financial_income', 'total_income'),
    'Ingresos Financieros Anualizados / Activo Productivo Promedio':
        ('financial_income_productive_assets_ratio', 'percent', 'rolling_12_months', 'annualized_financial_income', 'average_productive_assets'),
    'Créditos Directos / Personal ( S/ Miles )':
        ('direct_credit_per_person', 'thousands_PEN_per_person', 'point_in_time', 'direct_credit', 'personnel'),
    'Depósitos / Número de Oficinas ( S/ Miles )':
        ('deposits_per_office', 'thousands_PEN_per_office', 'point_in_time', 'deposits', 'offices'),
}
CR_EFFICIENCY = {
    'Gastos de Administración Anualizados/ Créditos Directos e Indirectos Promedio':
        ('administrative_expenses_average_credit_ratio', 'percent', 'rolling_12_months', 'annualized_administrative_expenses', 'average_direct_and_indirect_credit'),
    'Gastos de Operación Anualizados / Margen Financiero Total Anualizado':
        ('annualized_operating_expenses_financial_margin_ratio', 'percent', 'rolling_12_months', 'annualized_operating_expenses', 'annualized_total_financial_margin'),
    'Ingresos Financieros Anualizados / Activo Productivo Promedio':
        ('financial_income_productive_assets_ratio', 'percent', 'rolling_12_months', 'annualized_financial_income', 'average_productive_assets'),
    'Créditos Directos / Empleados (Miles S/)':
        ('direct_credit_per_employee', 'thousands_PEN_per_employee', 'point_in_time', 'direct_credit', 'employees'),
    'Créditos Directos / Número de Oficinas (Miles S/)':
        ('direct_credit_per_office', 'thousands_PEN_per_office', 'point_in_time', 'direct_credit', 'offices'),
    'Depósitos/ Créditos Directos':
        ('deposits_direct_credit_ratio', 'percent', 'point_in_time', 'deposits', 'direct_credit'),
}
EFFICIENCY_METRICS = {'B': BF_EFFICIENCY, 'F': {
    label.replace('Administración Anualizados', 'Administración Anualizado').replace('( S/ Miles )', '(S/ Miles)'): value
    for label, value in BF_EFFICIENCY.items()}, 'C': CR_EFFICIENCY, 'R': CR_EFFICIENCY}
TITLES = dict(zip('BFCR', ('Indicadores Financieros por Empresa Bancaria',
    'Indicadores Financieros por Empresa Financiera', 'Indicadores Financieros por Caja Municipal',
    'Indicadores Financieros por Caja Rural de Ahorro y Crédito')))


def observation_window(period, basis):
    end = closing(period)
    if basis == 'rolling_12_months':
        start = date(end.year-1+(end.month == 12), end.month % 12+1, 1)
        return start.isoformat(), end.isoformat()
    if basis == 'point_in_time':return end.isoformat(), end.isoformat()
    return '', ''


def parse_indicators(content, *, entity_type, period, source_url, retrieved_at, kind):
    sheet, frame = frame_for(content, 'Indicadores Financieros por')
    titles = [(i, j, clean(v)) for i in range(min(4, len(frame))) for j, v in enumerate(frame.iloc[i])
              if clean(v).startswith('Indicadores Financieros por')]
    if any(title != TITLES[entity_type] for _, _, title in titles):
        raise SchemaChangedError('Tipo de entidad del libro incompatible con la consulta.')
    label_cols = sorted({j for _, j, _ in titles})
    if not label_cols or label_cols[0] != 0:
        raise SchemaChangedError('Bloques de indicadores ausentes.')
    section, next_section = ('RENTABILIDAD', 'LIQUIDEZ') if kind == 'profitability' else ('EFICIENCIA Y GESTIÓN', 'RENTABILIDAD')
    heads = [i for i in range(min(10, len(frame))) if not clean(frame.iloc[i, 0])
             and isinstance(frame.iloc[i, 1], str) and clean(frame.iloc[i, 1])]
    if not heads:raise SchemaChangedError('Encabezado de entidades ausente.')
    # The first entity row is the header; subsequent header rows may contain auxiliary names.
    head = heads[0]
    notes_start = next((i for i in range(head+1, len(frame)) if any(clean(v).startswith(('Nota:', 'Los valores anualizados')) for v in frame.iloc[i])), None)
    if notes_start is None:raise SchemaChangedError('Notas de indicadores ausentes.')
    notes = source_notes(frame, notes_start)
    annual_texts = {n['text'] for n in notes if n['text'].startswith('Los valores anualizados')}
    if annual_texts != {ANNUAL_NOTE}:
        raise SchemaChangedError('Definición de anualización o promedio ausente o cambiada.')
    definitions = PROFIT_METRICS[entity_type] if kind == 'profitability' else EFFICIENCY_METRICS[entity_type]
    rows, names = [], set()
    for pos, label_col in enumerate(label_cols):
        validate_date(frame, period, head, label_col)
        if entity_type in 'BF' and not any(clean(v).replace(' ', '').lower() == '(enporcentaje)' for v in frame.iloc[:head, label_col]):
            raise SchemaChangedError('Unidad porcentual de indicadores ausente.')
        starts = [i for i in range(head+1, notes_start) if normalized(frame.iloc[i, label_col]) == section]
        stops = [i for i in range(head+1, notes_start) if normalized(frame.iloc[i, label_col]) == next_section]
        if len(starts) != 1 or len(stops) != 1 or stops[0] <= starts[0]:
            raise SchemaChangedError('Sección de indicadores ausente, duplicada o desordenada.')
        start, stop = starts[0], stops[0]
        pairs = [(normalized(frame.iloc[i, label_col]), i) for i in range(start+1, stop) if clean(frame.iloc[i, label_col])]
        if len({label for label, _ in pairs}) != len(pairs) or {label for label, _ in pairs} != set(definitions):
            raise SchemaChangedError('Indicadores ausentes, duplicados o no revisados.')
        metric_rows = dict(pairs)
        blank_rows = [i for i in range(start+1, stop) if not clean(frame.iloc[i, label_col])]
        end_col = label_cols[pos+1] if pos+1 < len(label_cols) else frame.shape[1]
        for col in range(label_col+1, end_col):
            raw_name = frame.iloc[head, col]; name = clean(raw_name)
            if frame.iloc[blank_rows, col].notna().any():
                raise SchemaChangedError('Valores en fila sin indicador.')
            if not name:
                if frame.iloc[list(metric_rows.values()), col].notna().any():
                    raise SchemaChangedError('Indicadores sin entidad.')
                continue
            if not isinstance(raw_name, str) or name in names:
                raise SchemaChangedError('Nombre de entidad no reconocido o duplicado.')
            names.add(name)
            # An auxiliary name is source evidence, not a dated legal alias.
            body_start = next((i for i in range(head+1, start) if normalized(frame.iloc[i, label_col]) == 'SOLVENCIA'), None)
            if body_start is None:raise SchemaChangedError('Inicio del cuerpo de indicadores ausente.')
            auxiliary = '; '.join(clean(v) for v in frame.iloc[head+1:body_start, col] if isinstance(v, str) and clean(v))
            for label, row in pairs:
                raw = frame.iloc[row, col]; value = number(raw)
                if kind == 'profitability':
                    metric, numerator, denominator = definitions[label]
                    unit, basis = 'percent', 'rolling_12_months'
                    numerator = 'annualized_'+numerator
                else:metric, unit, basis, numerator, denominator = definitions[label]
                observed_start, observed_end = observation_window(period, basis)
                flags = []
                if math.isnan(value):flags.append('source_value_missing')
                if auxiliary:flags.append('source_auxiliary_entity_name')
                rows.append(dict(period=period, period_date=closing(period).isoformat(), frequency='monthly',
                    entity_type=entity_type, entity_name=name, entity_scope=scope(name), section=section,
                    metric=metric, source_metric=clean(frame.iloc[row, label_col]), value=value,
                    source_value_token=clean(raw) if isinstance(raw, str) else '', unit=unit,
                    unit_multiplier=1000.0 if unit.startswith('thousands_PEN') else float('nan'), currency='TOTAL',
                    numerator_basis=numerator, denominator_basis=denominator, measurement_basis=basis,
                    observation_start=observed_start, observation_end=observed_end,
                    annualization_definition=ANNUAL_NOTE if basis == 'rolling_12_months' else '',
                    unit_evidence='table_header' if entity_type in 'BF' else 'metric_label',
                    source_auxiliary_entity_name=auxiliary, data_quality_flags=';'.join(flags),
                    source_sheet=str(sheet), source_row=row+1, source_column=col+1,
                    source='SBS', source_url=source_url, retrieved_at=retrieved_at))
    if not rows or not any(r['entity_scope'] == 'system_aggregate' for r in rows):
        raise SchemaChangedError('Indicadores sin observaciones o agregado publicado.')
    guard_percentage_formats(content, rows)
    notes.append({'period_contract': 'Annualized values and averages use the published twelve-month definition; do not multiply current-month profit by twelve.',
                  'comparability': 'Preserve different denominators and explicit annualized labels; unqualified flow ratios have no inferred observation window.',
                  'auxiliary_names': 'Additional source names are not legal equivalences or dated succession evidence.'})
    data = pd.DataFrame(rows)
    for col in data:
        data[col] = data[col].astype('float64' if col in ('value', 'unit_multiplier') else
                                    'int64' if col in ('source_row', 'source_column') else 'string')
    return data, notes


def parse_profitability(content, **kwargs):
    return parse_indicators(content, kind='profitability', **kwargs)


def parse_efficiency(content, **kwargs):
    return parse_indicators(content, kind='efficiency', **kwargs)


class ProfitabilityProvider(MonthlyExcelProvider):
    codes = QUALITY_CODES
    parser_version = '2026-10-08.1'
    parse_workbook = staticmethod(parse_profitability)


class EfficiencyProvider(MonthlyExcelProvider):
    codes = QUALITY_CODES
    parser_version = '2026-10-08.1'
    parse_workbook = staticmethod(parse_efficiency)
