"""Conservative, dated name resolution independent of source providers."""
from datetime import date
from pathlib import Path
import json
import os
import uuid

import pandas as pd

from .providers.sbs._universe_parser import normalized_entity_name

NAMESPACE = uuid.UUID('3152f1ba-22de-4a75-9c09-c43420df58f3')
MATCH_COLUMNS = ['dataset', 'period', 'entity_type_code', 'entity_name',
                 'entity_id', 'match_status', 'evidence']


class EntityCatalog:
    """Internal identifiers, never presented as official SBS identifiers."""
    def __init__(self, path):
        self.path = Path(path)
        self.records = json.loads(self.path.read_text(encoding='utf-8')) if self.path.exists() else []
        if not isinstance(self.records, list):
            raise ValueError('El catálogo de identidades debe ser una lista.')
        ids = [r['entity_id'] for r in self.records]
        keys = [(r['entity_type_code'], r['normalized_name']) for r in self.records]
        if len(ids) != len(set(ids)) or len(keys) != len(set(keys)):
            raise ValueError('Identidades duplicadas en el catálogo.')

    def observe(self, universe, aliases=()):
        index = {(r['entity_type_code'], r['normalized_name']): r for r in self.records}
        for row in universe.sort_values('period').to_dict('records'):
            key = (row['entity_type_code'], normalized_entity_name(row['entity_name']))
            period = date.fromisoformat(row['period']).isoformat()
            identity, status, _ = self.resolve(dataset='pe.sbs.universo_depositos',
                type_code=key[0], name=row['entity_name'], period=period, aliases=aliases)
            if status == 'ambiguous':
                raise ValueError(f'Identidad ambigua en el universo: {row["entity_name"]}')
            if identity is not None and key not in index:
                record = next(r for r in self.records if r['entity_id'] == identity)
                record['first_observed'] = min(record['first_observed'], period)
                record['last_observed'] = max(record['last_observed'], period)
                continue
            if key not in index:
                record = {
                    'entity_id': 'internal-' + str(uuid.uuid5(NAMESPACE, ':'.join(key))),
                    'entity_type_code': key[0], 'normalized_name': key[1],
                    'entity_name': row['entity_name'],
                    'first_observed': period, 'last_observed': period,
                    'source_url': row['source_url'],
                }
                self.records.append(record)
                index[key] = record
            else:
                record = index[key]
                record['first_observed'] = min(record['first_observed'], period)
                record['last_observed'] = max(record['last_observed'], period)

    def save(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary = self.path.with_suffix('.json.tmp')
        temporary.write_text(json.dumps(self.records, ensure_ascii=False, indent=2), encoding='utf-8')
        os.replace(temporary, self.path)

    def resolve(self, *, dataset, type_code, name, period, aliases=()):
        norm = normalized_entity_name(name)
        when = date.fromisoformat(period).isoformat()
        ids = {r['entity_id'] for r in self.records}
        candidates = []
        for alias in aliases:
            if alias['entity_id'] not in ids or not alias.get('evidence', '').strip():
                raise ValueError('Cada alias requiere una identidad existente y evidencia.')
            start = date.fromisoformat(alias['valid_from']).isoformat()
            end = date.fromisoformat(alias['valid_to']).isoformat() if alias.get('valid_to') else '9999-12-31'
            if start > end:
                raise ValueError('Intervalo de alias inválido.')
            if (alias['dataset'] == dataset and alias['entity_type_code'] == type_code
                    and normalized_entity_name(alias['alias']) == norm and start <= when <= end):
                candidates.append((alias['entity_id'], 'alias', alias['evidence']))
        # An explicit alias can describe a dated rename or type conversion.
        # Conflicts with an exact identity remain unresolved, never silently overridden.
        candidates.extend((r['entity_id'], 'exact', r['source_url']) for r in self.records
                          if r['entity_type_code'] == type_code and r['normalized_name'] == norm)
        unique = {c[0] for c in candidates}
        if len(unique) > 1:
            return None, 'ambiguous', ''
        if not candidates:
            return None, 'unmatched', ''
        return candidates[0]


def correspondence_report(catalog, *, rates=None, ratings=None, aliases=()):
    rows = []
    for dataset, frame, type_column in (
        ('pe.sbs.tasas_pasivas', rates, 'entity_type'),
        ('pe.sbs.clasificaciones_riesgo', ratings, 'entity_type_code'),
    ):
        if frame is None or frame.empty:
            continue
        columns = [type_column, 'entity_name', 'period_date']
        for row in frame[columns].drop_duplicates().to_dict('records'):
            period = str(row['period_date'])[:10]
            identity, status, evidence = catalog.resolve(
                dataset=dataset, type_code=row[type_column], name=row['entity_name'],
                period=period, aliases=aliases,
            )
            rows.append(dict(zip(MATCH_COLUMNS, [dataset, period, row[type_column],
                row['entity_name'], identity, status, evidence])))
    return pd.DataFrame(rows, columns=MATCH_COLUMNS)


def observation_changes(universe, catalog, aliases=()):
    """Differences between captured lists, not legal authorization events."""
    columns = ['period', 'previous_period', 'entity_id', 'entity_name', 'change']
    rows, previous, previous_date = [], None, None
    for period, frame in universe.groupby('period', sort=True):
        current = {}
        for row in frame.to_dict('records'):
            identity, _, _ = catalog.resolve(dataset='pe.sbs.universo_depositos',
                type_code=row['entity_type_code'], name=row['entity_name'], period=period, aliases=aliases)
            current[identity] = row['entity_name']
        if previous is not None:
            for identity in current.keys() - previous.keys():
                rows.append([period, previous_date, identity, current[identity], 'appeared'])
            for identity in previous.keys() - current.keys():
                rows.append([period, previous_date, identity, previous[identity], 'disappeared'])
        previous, previous_date = current, period
    return pd.DataFrame(rows, columns=columns)
