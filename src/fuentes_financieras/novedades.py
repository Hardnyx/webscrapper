"""Transactional capture revisions, independent of provider implementations."""
from contextlib import contextmanager
import json
from pathlib import Path
import sqlite3

from .storage import utc_now_iso

COLUMNS = ['Evento', 'Ejecución', 'Detectado UTC', 'Dataset', 'Período', 'Tipo',
           'Huella anterior', 'Huella actual']


def observations(planned, coverage):
    """Only accept captures already verified by the batch executor."""
    valid = {(row['Trabajo'], row['Dataset'], row['Período solicitado'])
             for row in coverage if row['Estado de captura'] == 'Captura validada'}
    found = {}
    for job in planned:
        dataset = job.provider.spec.dataset_id
        for request in job.requests:
            if (job.name, dataset, request.period_key) not in valid:
                continue
            entry = job.provider.storage.manifest.get(request.period_key)
            fingerprint = {'contenido': entry['content_hash'],
                'pdf': entry.get('metadata', {}).get('pdf_sha256'),
                'parser': entry['parser_version'], 'contrato': entry.get('contract_version', '1')}
            key = (dataset, request.period_key)
            if key in found and found[key] != fingerprint:
                raise ValueError('Capturas contradictorias para el mismo dataset y período.')
            found[key] = fingerprint
    return found


@contextmanager
def record_changes(path, run_id, captures):
    """Commit observations only after the consumer has exported its report."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path, timeout=30)
    try:
        connection.execute('BEGIN IMMEDIATE')
        version = connection.execute('PRAGMA user_version').fetchone()[0]
        if version not in (0, 1):
            raise ValueError('Versión del registro de novedades incompatible.')
        connection.execute('CREATE TABLE IF NOT EXISTS captures '
            '(dataset TEXT, period TEXT, fingerprint TEXT NOT NULL, PRIMARY KEY(dataset, period))')
        connection.execute('CREATE TABLE IF NOT EXISTS events '
            '(id INTEGER PRIMARY KEY AUTOINCREMENT, run_id TEXT NOT NULL, detected TEXT NOT NULL, '
            'dataset TEXT NOT NULL, period TEXT NOT NULL, kind TEXT NOT NULL, previous TEXT, current TEXT NOT NULL)')
        connection.execute('PRAGMA user_version=1')
        changes = []
        detected = utc_now_iso()
        for (dataset, period), fingerprint in sorted(captures.items()):
            current = json.dumps(fingerprint, ensure_ascii=False, sort_keys=True)
            row = connection.execute('SELECT fingerprint FROM captures WHERE dataset=? AND period=?',
                                     (dataset, period)).fetchone()
            previous = row[0] if row else None
            if previous == current:
                continue
            old = json.loads(previous) if previous else None
            if old is None:
                kind = 'Primera captura observada'
            elif old['pdf'] != fingerprint['pdf']:
                kind = 'Documento PDF cambiado'
            elif (old['parser'], old['contrato']) != (fingerprint['parser'], fingerprint['contrato']):
                kind = 'Reprocesamiento o cambio de contrato'
            else:
                kind = 'Contenido canónico cambiado'
            values = (run_id, detected, dataset, period, kind, previous, current)
            cursor = connection.execute('INSERT INTO events '
                '(run_id,detected,dataset,period,kind,previous,current) VALUES (?,?,?,?,?,?,?)', values)
            connection.execute('INSERT INTO captures VALUES (?,?,?) '
                'ON CONFLICT(dataset,period) DO UPDATE SET fingerprint=excluded.fingerprint',
                (dataset, period, current))
            changes.append(dict(zip(COLUMNS, (cursor.lastrowid, *values))))
        yield changes
        connection.commit()
    except BaseException:
        connection.rollback()
        raise
    finally:
        connection.close()
