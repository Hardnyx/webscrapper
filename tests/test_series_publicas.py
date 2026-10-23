"""Independent public-series contracts, cache integrity and credential redaction."""
import json
from types import SimpleNamespace
from datetime import datetime, timezone

import pandas as pd
import pytest

from fuentes_financieras import source
from fuentes_financieras.registry import get_provider
from fuentes_financieras.exceptions import InvalidQueryError, SchemaChangedError, SourceUnavailableError
from fuentes_financieras.providers._series import bound, numeric
from fuentes_financieras.providers.sbs.curva_soberana import parse_curve


@pytest.fixture(autouse=True)
def isolate(tmp_path,monkeypatch):
    monkeypatch.setenv('FINANCIAL_SOURCES_DATA_ROOT',str(tmp_path/'cache'))
    get_provider.cache_clear()
    yield
    get_provider.cache_clear()


def mock_body(monkeypatch,provider,payload):
    raw=payload if isinstance(payload,str) else json.dumps(payload)
    calls=[]
    def request(*args,**kwargs):
        calls.append((args,kwargs))
        return SimpleNamespace(status_code=200,content=raw.encode(),text=raw)
    monkeypatch.setattr(provider.transport,'request',request)
    return calls


def test_annual_capture_reused_for_narrower_range_and_missing_zero_preserved(monkeypatch):
    p=source('br.bcb.series')
    calls=mock_body(monkeypatch,p,[{'data':'01/08/2026','valor':'0'},{'data':'01/09/2026','valor':'-'},{'data':'02/09/2026','valor':'-1.25'}])
    query=dict(codigo='432',frecuencia='daily',desde='2026-08',hasta='2026-09')
    first=p.sync(**query)
    assert first.downloaded==1
    assert p.sync(**{**query,'desde':'2026-09'}).skipped_existing==1 and len(calls)==1
    data=p.load(**query)
    assert data.value.iloc[0]==0 and pd.isna(data.value.iloc[1]) and data.value.iloc[2]==-1.25
    assert data.data_quality_flags.iloc[1]=='source_value_missing'
    assert len(p.load(**{**query,'desde':'2026-09'}))==2
    assert calls[0][1]['params']['dataInicial']=='01/01/2026'


def test_failed_or_duplicate_refresh_keeps_previous_canonical(monkeypatch):
    p=source('br.bcb.series');query=dict(codigo='432',frecuencia='daily',desde='2026-09',hasta='2026-09')
    mock_body(monkeypatch,p,[{'data':'01/09/2026','valor':'1'}]);assert p.sync(**query).downloaded==1
    mock_body(monkeypatch,p,[{'data':'01/09/2026','valor':'1'},{'data':'01/09/2026','valor':'2'}])
    assert p.sync(force=True,**query).failed==1
    assert p.load(**query).value.tolist()==[1]
    mock_body(monkeypatch,p,[{'data':'broken','valor':'1'}])
    assert p.sync(force=True,**query).failed==1
    assert p.load(**query).value.tolist()==[1]


def test_bcrp_native_dates_and_query_frequency(monkeypatch):
    p=source('pe.bcrp.series')
    mock_body(monkeypatch,p,{'config':{'series':[{'name':'IPC'}]},'periods':[{'name':'Set.2026','values':['1,234.5']}]})
    d=p.fetch(codigo='PN01206PM',frecuencia='monthly',anio=2026).data
    assert d.period_date.tolist()==['2026-09-01'] and d.value.tolist()==[1234.5]
    with pytest.raises(InvalidQueryError):p.fetch(codigo='PN01206PM',frecuencia='daily',anio=2026)
    mock_body(monkeypatch,p,{'config':{'series':[{'name':'FX'}]},'periods':[{'name':'03.Ene.97','values':['3.1']}]})
    assert p.fetch(codigo='PD04638PD',frecuencia='daily',anio=1997).data.period_date.tolist()==['1997-01-03']


def test_credential_absence_and_transport_redaction(monkeypatch):
    p=source('cl.bcch.series')
    monkeypatch.delenv('BCCH_USER',raising=False);monkeypatch.delenv('BCCH_PASSWORD',raising=False)
    with pytest.raises(SourceUnavailableError,match='BCCH_USER'):p.fetch(codigo='F073.TPMTAS.D',frecuencia='daily',anio=2026)
    monkeypatch.setenv('BCCH_USER','private-user');monkeypatch.setenv('BCCH_PASSWORD','private-secret')
    def fail(*args,**kwargs):raise RuntimeError('https://host/?user=private-user&pass=private-secret')
    monkeypatch.setattr(p.transport,'request',fail)
    result=p.sync(codigo='F073.TPMTAS.D',frecuencia='daily',desde='2026-09',hasta='2026-09')
    assert result.failed==1 and 'private-' not in json.dumps(result.details)


def test_banxico_identity_and_header(monkeypatch):
    p=source('mx.banxico.series');monkeypatch.setenv('BANXICO_TOKEN','not-a-real-token')
    calls=mock_body(monkeypatch,p,{'bmx':{'series':[{'idSerie':'SF61745','titulo':'Tasa','datos':[{'fecha':'01/09/2026','dato':'N/E'}]}]}})
    data=p.fetch(codigo='SF61745',frecuencia='daily',anio=2026).data
    assert data.value.isna().all() and calls[0][1]['headers']['Bmx-Token']=='not-a-real-token'
    assert 'not-a-real-token' not in data.to_json()


def test_bcch_latin_title_and_list_form(monkeypatch):
    p=source('cl.bcch.series');monkeypatch.setenv('BCCH_USER','u');monkeypatch.setenv('BCCH_PASSWORD','p')
    mock_body(monkeypatch,p,{'Codigo':0,'Series':[{'descripEsp':'Inflación','Obs':[{'indexDateString':'01-09-2026','value':'1,25','statusCode':'OK'}]}]})
    d=p.fetch(codigo='ABC.M',frecuencia='monthly',anio=2026).data
    assert d.value.tolist()==[1.25] and '?' not in d.source_url.iloc[0]


def test_bce_checks_series_and_preserves_scale(monkeypatch):
    p=source('eu.bce.series')
    mock_body(monkeypatch,p,'KEY,TIME_PERIOD,OBS_VALUE,UNIT,UNIT_MULT,OBS_STATUS\nEXR.D.USD.EUR.SP00.A,2026-09-01,1.1,USD,0,A\n')
    d=p.fetch(codigo='EXR/D.USD.EUR.SP00.A',frecuencia='daily',anio=2026).data
    assert d.value.tolist()==[1.1] and d.unit.tolist()==['USD'] and d.source_unit_multiplier.tolist()==['0']
    with pytest.raises(SchemaChangedError):p.fetch(codigo='EXR/D.GBP.EUR.SP00.A',frecuencia='daily',anio=2026)


def test_eurostat_singleton_dimensions_and_sparse_missing(monkeypatch):
    p=source('eu.eurostat.series')
    doc={'id':['freq','geo','time'],'size':[1,1,2],'dimension':{'freq':{'category':{'index':{'M':0}}},'geo':{'category':{'index':{'FR':0}}},'time':{'category':{'index':{'2026-08':0,'2026-09':1}}}},'value':{'0':2.1},'label':'Series'}
    mock_body(monkeypatch,p,doc)
    d=p.fetch(codigo='irt_lt_mcby_m?geo=FR',frecuencia='monthly',anio=2026).data
    assert len(d)==2 and pd.isna(d.value.iloc[1])
    doc['size'][1]=2;mock_body(monkeypatch,p,doc)
    with pytest.raises(SchemaChangedError):p.fetch(codigo='irt_lt_mcby_m?geo=FR',frecuencia='monthly',anio=2026)


@pytest.mark.parametrize('dataset,code,freq,body,expected',[
 ('us.fred.series','DGS10','daily','observation_date,DGS10\n2026-09-01,.\n2026-09-02,4.5\n',[None,4.5]),
 ('us.tesoro.series','10 Yr','daily','Date,10 Yr\n09/01/2026,4.5\n',[4.5]),
 ('uk.boe.series','IUDBEDR','daily','DATE,IUDBEDR\n01 Sep 2026,4.5\n',[4.5]),
 ('de.bundesbank.series','D.I.ZST','daily','BBSIS.D.I.ZST\n2026-09-01;2,34;\n',[2.34]),
 ('uk.ons.series','/economy/inflationandpriceindices/timeseries/d7g7/mm23','monthly','"Title","CPI"\n"2026","5"\n"2026 Q3","6"\n"2026 SEP","2.4"\n',[2.4]),
])
def test_csv_adapters(monkeypatch,dataset,code,freq,body,expected):
    p=source(dataset);mock_body(monkeypatch,p,body)
    d=p.fetch(codigo=code,frecuencia=freq,anio=2026).data
    assert len(d)==len(expected)
    for actual,target in zip(d.value,expected):assert pd.isna(actual) if target is None else actual==target


def test_yahoo_preserves_precision_uses_exchange_zone_and_excludes_live_session(monkeypatch):
    p=source('global.yahoo.precios');now=datetime.now(timezone.utc).timestamp()
    old=datetime(2026,1,2,1,tzinfo=timezone.utc).timestamp()
    doc={'chart':{'error':None,'result':[{'meta':{'symbol':'^TEST','exchangeTimezoneName':'America/Lima','instrumentType':'INDEX','currentTradingPeriod':{'regular':{'start':now-10,'end':now+100}}},'timestamp':[old,now], 'indicators':{'quote':[{'close':[123.123456,999]}]}}]}}
    mock_body(monkeypatch,p,doc)
    d=p.fetch(codigo='^TEST',frecuencia='daily',anio=2026).data
    assert d.period_date.tolist()==['2026-01-01'] and d.value.tolist()==[123.123456]


def test_focus_rejects_truncation_and_keeps_survey_scope(monkeypatch):
    p=source('br.bcb.focus')
    mock_body(monkeypatch,p,{'value':[{'Data':'2026-09-01','Mediana':4.21}]})
    assert p.fetch(codigo='IPCA12',frecuencia='daily',anio=2026).data.value.tolist()==[4.21]
    mock_body(monkeypatch,p,{'value':[],'@odata.nextLink':'https://host/next'})
    with pytest.raises(SchemaChangedError):p.fetch(codigo='IPCA12',frecuencia='daily',anio=2026)


def test_invalid_values_dates_and_oversized_query():
    assert bound('2026-02',True).isoformat()=='2026-02-28'
    assert numeric('0')==0 and pd.isna(numeric('N/E'))
    for value in ['bad','inf',True]:
        with pytest.raises(SchemaChangedError):numeric(value)
    p=source('br.bcb.series')
    with pytest.raises(InvalidQueryError):list(p.plan_sync(codigo='432',frecuencia='daily',desde='1900-01',hasta='2100-01'))
    with pytest.raises(InvalidQueryError):p.single_request(codigo='432&evil=x',frecuencia='daily',anio=2026)


def test_curve_full_precision_duplicate_and_wrong_date():
    payload={'success':True,'data':[{'fechaProceso':'2026-10-22T00:00:00','plazos':90,'tasa':4.123456789},{'fechaProceso':'2026-10-22','plazos':0,'tasa':0}]}
    data=parse_curve(payload,curve='CCPSS',target='2026-10-22',retrieved_at='now')
    assert data.tenor_days.tolist()==[0,90] and data.rate.tolist()==[0,4.123456789]
    with pytest.raises(SchemaChangedError):parse_curve(payload,curve='CCPSS',target='2026-10-21',retrieved_at='now')
    payload['data'].append(payload['data'][0])
    with pytest.raises(SchemaChangedError):parse_curve(payload,curve='CCPSS',target='2026-10-22',retrieved_at='now')


def test_bundesbank_current_comma_csv_has_identified_key(monkeypatch):
    p=source('de.bundesbank.series')
    calls=mock_body(monkeypatch,p,'"",BBSIS.D.I.ZST,BBSIS.D.I.ZST_FLAGS\nunit,Prozent,\n2026-09-01,2.96,\n2026-09-02,.,No value available\n')
    d=p.fetch(codigo='D.I.ZST',frecuencia='daily',anio=2026).data
    assert d.value.iloc[0]==2.96 and pd.isna(d.value.iloc[1]) and d.unit.iloc[0]=='Prozent'
    with pytest.raises(SchemaChangedError):p.fetch(codigo='D.I.OTHER',frecuencia='daily',anio=2026)


def test_focus_percent_encoded_filter_not_form_plus(monkeypatch):
    p=source('br.bcb.focus');calls=mock_body(monkeypatch,p,{'value':[{'Data':'2026-09-01','Mediana':4.21}]})
    p.fetch(codigo='IPCA12',frecuencia='daily',anio=2026)
    url=calls[0][0][1]
    assert '%20eq%20' in url and 'Suavizada%20eq%20%27S%27' in url and '+' not in url


def test_spf_selects_vintage_column_without_merging_horizons(monkeypatch):
    import io
    from openpyxl import Workbook
    book=Workbook();sheet=book.active
    sheet.append(['YEAR','QUARTER','CPIA','CPIB'])
    sheet.append([2026,1,2.5,3.1]);sheet.append([2026,2,None,3.2])
    buffer=io.BytesIO();book.save(buffer)
    p=source('us.fed_filadelfia.spf')
    monkeypatch.setattr(p.transport,'request',lambda *a,**k:SimpleNamespace(status_code=200,content=buffer.getvalue(),text=''))
    d=p.fetch(codigo='CPIA',frecuencia='quarterly',anio=2026).data
    assert len(d)==2 and d.value.iloc[0]==2.5 and pd.isna(d.value.iloc[1])
    assert set(d.observation_status)=={'survey_vintage_not_target_period'}


def test_spanish_security_rejection_is_a_block_even_with_http_200():
    from fuentes_financieras.transports.curl_chrome import looks_like_waf
    assert looks_like_waf('<html><title>Solicitud Rechazada</title></html>',200)


def test_curve_checks_echoed_curve_type():
    with pytest.raises(SchemaChangedError):
        parse_curve({'success':True,'data':[{'fechaProceso':'2026-10-22','tipoCurva':'CCPVS','plazos':90,'tasa':4.1}]},curve='CCPSS',target='2026-10-22',retrieved_at='now')


def test_explicit_load_reads_only_requested_series_and_year(monkeypatch):
    p=source('br.bcb.series');mock_body(monkeypatch,p,[{'data':'01/09/2026','valor':'1'}])
    query=dict(codigo='432',frecuencia='daily',desde='2026-09',hasta='2026-09');p.sync(**query)
    def deny():raise AssertionError('Full history scan forbidden')
    monkeypatch.setattr(p.storage,'read_all',deny)
    assert p.load(**query).value.tolist()==[1]


def test_transport_ajax_header_is_explicit_without_changing_existing_default(monkeypatch):
    p=source('br.bcb.series');calls=[]
    def response(*a,**kwargs):
        calls.append(kwargs)
        return SimpleNamespace(status_code=200,text='[]',content=b'[]')
    monkeypatch.setattr(p.transport.session,'request',response)
    p.transport.request('GET','https://example.test',ajax=False)
    p.transport.request('GET','https://example.test')
    assert 'X-Requested-With' not in calls[0]['headers']
    assert calls[1]['headers']['X-Requested-With']=='XMLHttpRequest'
    assert all(call['verify']==str(p.transport.ca_bundle) for call in calls)
