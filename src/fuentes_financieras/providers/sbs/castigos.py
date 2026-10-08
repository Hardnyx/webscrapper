"""Explicit monthly credit write-off flows, preserving published classifications."""
from datetime import date
import re

from fuentes_financieras.exceptions import SchemaChangedError
from ._excel_mensual import MonthlyExcelProvider, clean
from .calidad_cartera import closing
from .fondeo import table, header, period_date, footer, entities, record, finish, compact, mark_sum

CODES = dict(zip('BFCR', ('B-2369', 'B-3234', 'C-1253', 'C-2258')))
CREDIT_LABELS = ('corporativos', 'grandesempresas', 'medianasempresas', 'pequeñasempresas', 'microempresas', 'consumo', 'hipotecarios')
CREDIT_IDS = ('corporate', 'large_enterprise', 'medium_enterprise', 'small_enterprise', 'microenterprise', 'consumer', 'mortgage', 'total')
MONTH_NAMES = ('enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre')


def parse_workbook(content, *, entity_type, period, source_url, retrieved_at):
    sheet, frame = table(content, 'Flujo de Créditos Castigados')
    head = header(frame); first = 2 if entity_type in 'BF' else 1
    if frame.shape[1] != first+8 or clean(frame.iloc[head, first+7]).lower() != 'total':
        raise SchemaChangedError('Dimensiones o total de castigos cambiados.')
    if [compact(v) for v in frame.iloc[head+1, first:first+7]] != list(CREDIT_LABELS):
        raise SchemaChangedError('Tipos de crédito de castigos cambiados.')
    if not any(compact(v) == '(enmilesdesoles)' for v in frame.iloc[:head, 0]):
        raise SchemaChangedError('Unidad de castigos cambiada.')
    if entity_type in 'BF':
        dates = [re.fullmatch(r'en el mes de (\w+) de (\d{4})', clean(v), re.I) for v in frame.iloc[:head, 0]]
        dates = [m for m in dates if m]
        if len(dates) != 1 or dates[0][1].lower() not in MONTH_NAMES:
            raise SchemaChangedError('Castigos sin período mensual explícito.')
        observed = f'{dates[0][2]}-{MONTH_NAMES.index(dates[0][1].lower())+1:02d}'
        if observed != period:raise SchemaChangedError('Mes de castigos distinto del solicitado.')
        allowed = {'flujodecastigos', 'flujomensualdecastigos'}
    else:
        period_date(frame, period, head)
        allowed = {'flujomensualdecastigos'}
    if compact(frame.iloc[head, first]) not in allowed:
        raise SchemaChangedError('Base de castigos no mensual o desconocida.')
    stop, notes = footer(frame, head+2); rows = []
    for row, name, original, placeholder in entities(frame, head+2, stop, range(first, first+8)):
        if first == 2 and clean(frame.iloc[row, 1]):raise SchemaChangedError('Valor en columna auxiliar de castigos.')
        start = len(rows)
        for k, credit_type in enumerate(CREDIT_IDS):
            col = first+k
            r = record(entity_type=entity_type, period=period, sheet=sheet, row=row, col=col,
                name=name, metric='writeoff_flow', label=clean(frame.iloc[head+1, col]) if k < 7 else clean(frame.iloc[head, col]),
                raw=frame.iloc[row, col], source_url=source_url, retrieved_at=retrieved_at,
                credit_type=credit_type, source_entity_name=original, measurement_basis='monthly_flow',
                observation_start=date.fromisoformat(period+'-01').isoformat(), observation_end=closing(period).isoformat(),
                flags=['source_entity_placeholder'] if placeholder else [])
            if placeholder:r['entity_scope'] = 'unidentified_source_row'
            rows.append(r)
        mark_sum(rows, start+7, list(range(start, start+7)), 'published_components_mismatch')
    notes.append({'measurement_basis': 'Monthly flow; not YTD or rolling twelve months.',
                  'comparability': 'Preserve source notes and classifications; no historical reclassification.'})
    return finish(rows, notes)


class CreditWriteoffsProvider(MonthlyExcelProvider):
    codes = CODES
    parser_version = '2026-10-08.1'
    parse_workbook = staticmethod(parse_workbook)
