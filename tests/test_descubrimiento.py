"""Stable publication identities across pages and transactional ledger upgrades."""
import json
import sqlite3
from types import SimpleNamespace

import pandas as pd
import pytest

from fuentes_financieras.descubrimiento import reference_observations
from fuentes_financieras.ejecucion import PlannedJob
from fuentes_financieras.models import PeriodRequest
from fuentes_financieras.novedades import record_changes
from fuentes_financieras.storage import DatasetStorage


def plan(tmp_path, rows, dataset='pe.sbs.indice_noticias', periods=('1',)):
    storage = DatasetStorage(tmp_path, dataset)
    for period, frame in rows.items():
        storage.upsert_period('indice', period, pd.DataFrame(frame).astype('string'))
    provider = SimpleNamespace(storage=storage, spec=SimpleNamespace(dataset_id=dataset))
    job = PlannedJob('index',provider,{}, {},[PeriodRequest(p,'indice',{}) for p in periods])
    coverage = [{'Trabajo':'index','Dataset':dataset,'Período solicitado':p,'Estado de captura':'Captura validada'} for p in periods]
    return [job], coverage


def news(identifier='a',title='Noticia',date='2026-10-01',page='1'):
    return dict(announcement_id=identifier,article_url='https://www.sbs.gob.pe/noticia/'+identifier,
                source_title=title,listed_date=date,index_page=page,retrieved_at='irrelevant')


def observe(path,run,refs):
    with record_changes(path,run,{},refs) as events:
        return events


def test_moving_news_deduplicates_and_unselected_pages_are_ignored(tmp_path):
    planned,coverage=plan(tmp_path,{'1':[news()], '2':[news(page='2')], '3':[news('unselected')]},periods=('1','2'))
    refs=reference_observations(planned,coverage)
    assert len(refs)==1
    path=tmp_path/'ledger.sqlite3'
    events=observe(path,'1',refs)
    assert events[0]['Tipo']=='Referencia inicial observada'
    planned,coverage=plan(tmp_path,{'1':[news('new')], '2':[news(page='2')]},periods=('1','2'))
    events=observe(path,'2',reference_observations(planned,coverage))
    assert len(events)==1 and events[0]['Referencia']=='new'
    assert events[0]['Tipo']=='Referencia nueva observada'
    assert events[0]['Proveedor destino']=='pe.sbs.anuncios_regulatorios'
    assert observe(path,'3',refs)==[]


def test_metadata_update_and_absence_do_not_infer_withdrawal(tmp_path):
    planned,coverage=plan(tmp_path,{'1':[news()]})
    path=tmp_path/'ledger.sqlite3'
    observe(path,'1',reference_observations(planned,coverage))
    assert observe(path,'2',{})==[]
    planned,coverage=plan(tmp_path,{'1':[news(title='Título corregido')]})
    events=observe(path,'3',reference_observations(planned,coverage))
    assert events[0]['Tipo']=='Referencia actualizada'
    assert events[0]['Título']=='Título corregido'
    assert json.loads(events[0]['Huella anterior'])['titulo']=='Noticia'


def test_conflicting_pages_reject_before_observation(tmp_path):
    planned,coverage=plan(tmp_path,{'1':[news()], '2':[news(title='Contradictorio')]},periods=('1','2'))
    with pytest.raises(ValueError,match='contradictoria'):
        reference_observations(planned,coverage)
    coverage[1]['Estado de captura']='Actualización pendiente'
    assert len(reference_observations(planned,coverage))==1


@pytest.mark.parametrize('dataset,key,url,title,date,target',[
    ('pe.moodys.indice_comunicados','action_id','article_url','source_title','listed_date','pe.moodys.referencias_comunicados'),
    ('pe.moodys.referencias_comunicados','document_id','pdf_url','source_title','listed_date','pe.clasificadoras.retiros'),
    ('pe.sbs.informes_riesgo','report_id','report_url',None,'report_date','pe.sbs.documentos_riesgo'),
])
def test_adapters_preserve_date_without_using_summary_period(tmp_path,dataset,key,url,title,date,target):
    row={key:'identifier',url:'https://example.test/document',date:'' if title is None else '2026-10-01'}
    if title: row[title]='Publicación'
    planned,coverage=plan(tmp_path,{'1':[row]},dataset=dataset)
    refs=reference_observations(planned,coverage)
    payload=next(iter(refs.values()))
    assert payload['proveedor_destino']==target
    assert payload['fecha_publicada']==row[date]


def test_missing_inventory_links_are_skipped_and_missing_schema_rejected(tmp_path):
    row={'report_id':'','report_url':'','report_date':''}
    planned,coverage=plan(tmp_path,{'1':[row]},dataset='pe.sbs.informes_riesgo')
    assert reference_observations(planned,coverage)=={}
    planned,coverage=plan(tmp_path,{'1':[{'announcement_id':'a'}]})
    with pytest.raises(ValueError,match='incompatible'):
        reference_observations(planned,coverage)


def test_migrate_v1_and_rollback_publication_events(tmp_path):
    path=tmp_path/'ledger.sqlite3'
    with sqlite3.connect(path) as conn:
        conn.execute('CREATE TABLE captures(dataset TEXT,period TEXT,fingerprint TEXT NOT NULL,PRIMARY KEY(dataset,period))')
        conn.execute('CREATE TABLE events(id INTEGER PRIMARY KEY AUTOINCREMENT,run_id TEXT NOT NULL,detected TEXT NOT NULL,dataset TEXT NOT NULL,period TEXT NOT NULL,kind TEXT NOT NULL,previous TEXT,current TEXT NOT NULL)')
        conn.execute("INSERT INTO events VALUES (1,'old','date','d','p','old',NULL,'{}')")
        conn.execute('PRAGMA user_version=1')
    planned,coverage=plan(tmp_path,{'1':[news()]})
    refs=reference_observations(planned,coverage)
    with pytest.raises(OSError):
        with record_changes(path,'failed',{},refs):
            raise OSError('Export failed')
    with sqlite3.connect(path) as conn:
        assert conn.execute('PRAGMA user_version').fetchone()[0]==1
    events=observe(path,'success',refs)
    assert len(events)==1 and events[0]['Evento']==2
    with sqlite3.connect(path) as conn:
        assert conn.execute('PRAGMA user_version').fetchone()[0]==2
        assert conn.execute('SELECT COUNT(*) FROM events').fetchone()[0]==2
