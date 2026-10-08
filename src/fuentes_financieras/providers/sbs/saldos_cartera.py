"""Credit balances and provisions from the published financial statement block."""
import math

import pandas as pd

from fuentes_financieras.exceptions import SchemaChangedError
from ._excel_mensual import MonthlyExcelProvider
from .estados_financieros import CODES, parse_workbook as parse_statements

NET_CREDIT = 'CRÉDITOS NETOS DE PROVISIONES Y DE INGRESOS NO DEVENGADOS'
BALANCE_METRICS = {
    NET_CREDIT: 'net_credit_after_provisions_and_unearned_income',
    'Vigentes': 'performing_credit',
    'Refinanciados y Reestructurados': 'refinanced_restructured_credit',
    'Atrasados': 'past_due_credit',
    'Vencidos': 'overdue_credit',
    'En Cobranza Judicial': 'credit_in_judicial_collection',
    'Provisiones': 'credit_provisions_contra_asset',
    'Intereses y Comisiones no Devengados': 'unearned_interest_and_fees',
}


def parse_balances(content, **kwargs):
    # Reuse the statement parser and accounting guards, without calling another
    # provider or deriving any ratio. Provisions outside this block are excluded.
    data, notes = parse_statements(content, **kwargs)
    assets = data[data.section == 'assets'].copy()
    assets['label'] = assets.account.str.replace(r'\*+$', '', regex=True).str.strip()
    blocks = []
    for sheet, group in assets.groupby('source_sheet', sort=False):
        starts = group[group.label == NET_CREDIT].source_row.unique()
        if len(starts) != 1:
            raise SchemaChangedError('Bloque de créditos netos ausente o ambiguo.')
        start = starts[0]
        following = group[(group.source_row > start) & (group.label == 'CUENTAS POR COBRAR NETAS DE PROVISIONES')].source_row.unique()
        if len(following) != 1:
            raise SchemaChangedError('Cierre del bloque de créditos ausente o ambiguo.')
        block = group[(group.source_row >= start) & (group.source_row < following[0]) & group.label.isin(BALANCE_METRICS)].copy()
        for (_, currency), entity in block.groupby(['entity_name', 'currency'], sort=False):
            if set(entity.label) != set(BALANCE_METRICS) or entity.label.duplicated().any():
                raise SchemaChangedError('Saldos esenciales de créditos ausentes o duplicados.')
        if set(block.entity_name) != set(group.entity_name):
            raise SchemaChangedError('Bloque de créditos incompleto para alguna entidad.')
        block['metric'] = block.label.map(BALANCE_METRICS)
        blocks.append(block)
    if not blocks:
        raise SchemaChangedError('Sin bloques de créditos.')
    selected = pd.concat(blocks, ignore_index=True).drop(columns=['label', 'account_code'])
    selected = selected.rename(columns={'amount': 'value', 'account': 'source_metric'})
    selected['data_quality_flags'] = selected.value.isna().map({True: 'source_value_missing', False: ''}).astype('string')
    selected['metric'] = selected.metric.astype('string')
    for (_, _), group in selected.groupby(['entity_name', 'currency'], sort=False):
        values = dict(zip(group.metric, group.value))
        if all(math.isfinite(v) for v in values.values()):
            net = sum(values[k] for k in ('performing_credit', 'refinanced_restructured_credit',
                'past_due_credit', 'credit_provisions_contra_asset', 'unearned_interest_and_fees'))
            past_due = values['overdue_credit'] + values['credit_in_judicial_collection']
            if abs(net-values['net_credit_after_provisions_and_unearned_income']) > .02 or abs(past_due-values['past_due_credit']) > .02:
                selected.loc[group.index, 'data_quality_flags'] = 'published_credit_components_mismatch'

    notes.append({'provisions_sign': 'Preserve negative contra-asset amounts; do not take absolute values.',
                  'currency_scale': 'All MN/ME/TOTAL columns remain thousands of PEN.'})
    return selected, notes


class CreditBalancesProvider(MonthlyExcelProvider):
    codes = CODES
    parser_version = '2026-10-08.1'
    parse_workbook = staticmethod(parse_balances)
