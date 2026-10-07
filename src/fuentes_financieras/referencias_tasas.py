"""Product references from published SBS averages, without new downloads."""
import pandas as pd
from .exceptions import SchemaChangedError
from .providers.sbs.tasas_pasivas_mercado import METHODOLOGY_URL
from .providers.sbs._universe_parser import normalized_entity_name

PRODUCT_KEYS = ['entity_type', 'frequency', 'period', 'period_date', 'currency',
                'table_kind', 'person_type', 'metric']
PRODUCT_COLUMNS = PRODUCT_KEYS + ['rate', 'unit', 'basis', 'observation_window',
    'reference_kind', 'source', 'source_url', 'methodology_url', 'retrieved_at']


def product_benchmarks(rates):
    if rates.empty:
        return pd.DataFrame(columns=PRODUCT_COLUMNS)
    required = set(PRODUCT_KEYS + ['entity_name', 'rate', 'source', 'source_url', 'retrieved_at'])
    if required - set(rates.columns):
        raise SchemaChangedError(f'Faltan columnas de tasas: {sorted(required - set(rates.columns))}')
    frame = rates[rates.entity_name.map(normalized_entity_name) == 'PROMEDIO'].copy()
    if frame.duplicated(PRODUCT_KEYS).any():
        raise SchemaChangedError('Promedios duplicados para la misma clave de comparación.')
    if (~frame.entity_type.isin(['B', 'F', 'C', 'R'])).any():
        raise SchemaChangedError('Tipo desconocido en los promedios.')
    expected = frame.entity_type.map({'B': 'daily', 'F': 'daily', 'C': 'monthly', 'R': 'monthly'})
    if (expected != frame.frequency).any():
        raise SchemaChangedError('Frecuencia incompatible con el tipo de entidad.')
    frame['unit'] = 'percent_effective_annual'
    frame['basis'] = 'flow'
    frame['observation_window'] = frame.frequency.map({'daily': 'last_30_business_days', 'monthly': 'calendar_month'})
    frame['reference_kind'] = 'published_product_average'
    frame['methodology_url'] = METHODOLOGY_URL
    return frame[PRODUCT_COLUMNS].reset_index(drop=True)
