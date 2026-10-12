"""Capture revisions persist across runs without inferring withdrawals."""
import json
import sqlite3

import pytest

from fuentes_financieras.novedades import record_changes, observations
from fuentes_financieras import ejecucion
from test_ejecucion import setup_providers, job


def fingerprint(content='a', pdf=None, parser='1'):
    return {'contenido': content, 'pdf': pdf, 'parser': parser, 'contrato': '1'}


def observe(path, run, captures):
    with record_changes(path, run, captures) as changes:
        return changes


def test_repeat_reversal_and_unselected_periods(tmp_path):
    path = tmp_path/'changes.sqlite3'
    key = ('dataset', '2026-01')
    first = observe(path, '1', {key: fingerprint()})
    assert first[0]['Tipo'] == 'Primera captura observada'
    assert observe(path, '2', {key: fingerprint()}) == []
    assert observe(path, '3', {}) == []
    changed = observe(path, '4', {key: fingerprint('b')})
    assert changed[0]['Tipo'] == 'Contenido canónico cambiado'
    assert json.loads(changed[0]['Huella anterior'])['contenido'] == 'a'
    assert observe(path, '5', {key: fingerprint()})[0]['Evento'] == 3
    with sqlite3.connect(path) as conn:
        assert conn.execute('SELECT COUNT(*) FROM events').fetchone()[0] == 3


def test_pdf_and_parser_changes_are_distinguished(tmp_path):
    path = tmp_path/'changes.sqlite3'; key = ('dataset','p')
    observe(path,'1',{key:fingerprint(pdf='p1')})
    assert observe(path,'2',{key:fingerprint(pdf='p2')})[0]['Tipo'] == 'Documento PDF cambiado'
    assert observe(path,'3',{key:fingerprint('b',pdf='p2',parser='2')})[0]['Tipo'].startswith('Reprocesamiento')
    assert observe(path,'4',{key:fingerprint('b',pdf='p2',parser='3')})[0]['Tipo'].startswith('Reprocesamiento')


def test_export_failure_rolls_back_existing_state_and_events(tmp_path):
    path = tmp_path/'changes.sqlite3'; key = ('d','p')
    observe(path,'1',{key:fingerprint()})
    with pytest.raises(OSError):
        with record_changes(path,'2',{key:fingerprint('b')}) as changes:
            assert len(changes) == 1
            raise OSError('Export failed')
    changes = observe(path,'3',{key:fingerprint('b')})
    assert len(changes) == 1 and changes[0]['Evento'] == 2
    assert json.loads(changes[0]['Huella anterior'])['contenido'] == 'a'


def test_future_ledger_version_is_not_overwritten(tmp_path):
    path = tmp_path/'changes.sqlite3'
    with sqlite3.connect(path) as conn:
        conn.execute('PRAGMA user_version=2')
    with pytest.raises(ValueError):
        observe(path,'1',{})
    with sqlite3.connect(path) as conn:
        assert conn.execute('PRAGMA user_version').fetchone()[0] == 2


def test_verified_capture_deduplicates_jobs_and_excludes_invalid(tmp_path,monkeypatch):
    p = setup_providers(tmp_path,monkeypatch)['pe.sbs.rentabilidad']
    plan = ejecucion.plan_jobs([job(),job('alias')])
    _, coverage, _ = ejecucion.execute_jobs(plan)
    assert len(observations(plan,coverage)) == 1
    p.storage.partition_path('year=2026').write_bytes(b'broken')
    _, coverage, _ = ejecucion.execute_jobs(plan,cache_only=True)
    assert observations(plan,coverage) == {}


def test_cli_rollback_plan_and_failed_refresh(tmp_path,monkeypatch):
    from fuentes_financieras.cli import ejecucion as cli
    p = setup_providers(tmp_path,monkeypatch)['pe.sbs.rentabilidad']
    profile = tmp_path/'profile.toml'
    profile.write_text('version=1\n[[trabajos]]\nnombre="x"\ndataset="pe.sbs.rentabilidad"')
    args = ['--perfil',str(profile),'--data-root',str(tmp_path/'cache'),'--output-dir',str(tmp_path/'reports')]
    ledger = tmp_path/'cache/state/novedades.sqlite3'
    assert cli.run(args+['--solo-plan']) == 0 and not ledger.exists()
    writer = cli.write_report
    def fail(*args):
        raise OSError('Export unavailable')
    monkeypatch.setattr(cli,'write_report',fail)
    with pytest.raises(OSError):
        cli.run(args)
    monkeypatch.setattr(cli,'write_report',writer)
    assert cli.run(args+['--solo-cache']) == 0
    assert cli.run(args) == 0
    payloads = [json.loads(f.read_text()) for f in (tmp_path/'reports').glob('*/ejecucion.json')]
    assert sum(len(x['novedades']) for x in payloads) == 1
    assert all(x['version'] == 2 for x in payloads)
    p.fail = True
    profile.write_text(profile.read_text()+'\n[trabajos.opciones]\nforce=true\n')
    assert cli.run(args) == 1
    with sqlite3.connect(ledger) as conn:
        assert conn.execute('SELECT COUNT(*) FROM events').fetchone()[0] == 1
