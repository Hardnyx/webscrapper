"""Moody's action index and independent, explicit PDF reference resolution."""
import re
from datetime import datetime
from hashlib import sha256
from urllib.parse import parse_qs, urlsplit

import pandas as pd
from bs4 import BeautifulSoup
from fuentes_financieras.exceptions import InvalidQueryError, SchemaChangedError, SourceUnavailableError
from fuentes_financieras.models import FetchResult, PeriodRequest
from fuentes_financieras.provider import DatasetProvider
from fuentes_financieras.storage import utc_now_iso
from fuentes_financieras.transports.curl_chrome import CurlChromeTransport
from .retiros_clasificaciones import document_reference

BASE = 'https://moodyslocal.com.pe'
INDEX = BASE + '/reportes/acciones-de-calificacion/'
MONTHS = dict(zip('Ene Feb Mar Abr May Jun Jul Ago Sep Oct Nov Dic'.split(), range(1,13)))
MONTHS.update(Jan=1, Apr=4, Aug=8, Dec=12)


def action_reference(url):
    if not isinstance(url,str):raise InvalidQueryError('Indique una URL de acción Moody’s Local Perú.')
    p=urlsplit(url);q=parse_qs(p.query,keep_blank_values=True)
    friendly=bool(re.fullmatch(r'/reporte/rating-action/[a-z0-9-]+/?',p.path)) and not p.query
    numeric=(p.path=='/' and set(q)=={'post_type','p'} and q['post_type']==['rating-action']
             and len(q['p'])==1 and re.fullmatch(r'[1-9]\d{0,9}',q['p'][0]))
    if p.scheme!='https' or p.netloc!='moodyslocal.com.pe' or p.fragment or not (friendly or numeric):
        raise InvalidQueryError('Enlace de acción oficial ausente o ambiguo.')
    canonical=BASE+p.path.rstrip('/')+'/' if friendly else BASE+'/?post_type=rating-action&p='+q['p'][0]
    return sha256(canonical.encode()).hexdigest(),canonical


def soup_page(html):
    if len(html.encode())>2_000_000 or not re.search(r'</html\s*>',html,re.I):
        raise SchemaChangedError('HTML incompleto o demasiado grande.')
    return BeautifulSoup(html,'lxml')


def parse_index(html,*,retrieved_at):
    soup=soup_page(html);tables=soup.select('table#table_1[data-wpdatatable_id="35"]')
    if len(tables)!=1 or [t.get_text(' ',strip=True) for t in tables[0].select('thead th')]!=['Fecha','Título']:
        raise SchemaChangedError('Tabla de acciones ausente o ambigua.')
    rows=[];seen=set()
    for tr in tables[0].select('tbody tr'):
        cells=tr.find_all('td',recursive=False)
        if len(cells)!=2 or len(cells[1].select('a[href]'))!=1:raise SchemaChangedError('Fila incompleta.')
        a=cells[1].select_one('a[href]');key,url=action_reference(a['href']);title=a.get_text(' ',strip=True)
        try:dated=datetime.strptime(cells[0].get_text(strip=True),'%d/%m/%Y').date().isoformat()
        except ValueError:raise SchemaChangedError('Fecha de índice inválida.') from None
        if key in seen or not title:raise SchemaChangedError('Referencia duplicada o sin título.')
        seen.add(key);rows.append(dict(action_id=key,article_url=url,source_title=title,listed_date=dated,
            source='Moody’s Local Perú',source_url=INDEX,retrieved_at=retrieved_at))
    if not rows:raise SchemaChangedError('Índice vacío; no acredita ausencia de acciones.')
    return pd.DataFrame(rows).astype('string'),dict(html_sha256=sha256(html.encode()).hexdigest(),record_count=len(rows),
        coverage_scope='rows_in_delivered_html_not_complete_archive')


def parse_reference(html,*,url,retrieved_at):
    key,requested=action_reference(url);soup=soup_page(html)
    canonical=soup.select('link[rel="canonical"]');headings=soup.select('div.et_pb_title_container h1.entry-title')
    if len(canonical)!=1 or len(headings)!=1:raise SchemaChangedError('Identidad o título de acción ambiguos.')
    _,final=action_reference(canonical[0].get('href'))
    q=parse_qs(urlsplit(requested).query)
    if q:
        short=soup.select('link[rel="shortlink"]')
        if len(short)!=1 or short[0].get('href')!=BASE+'/?p='+q['p'][0]:raise SchemaChangedError('Identificador de acción distinto.')
    elif final!=requested:raise SchemaChangedError('Acción canónica distinta de la solicitada.')
    dates=headings[0].parent.select('span.published')
    downloads=[a for a in soup.select('a[href]') if a.get_text(' ',strip=True)=='Download']
    if len(dates)!=1 or len(downloads)!=1:raise SchemaChangedError('Fecha o descarga ausente o ambigua.')
    m=re.fullmatch(r'([A-Za-z]{3}) (\d{1,2}), (\d{4})',dates[0].get_text(strip=True))
    try:
        from datetime import date
        dated=date(int(m[3]),MONTHS[m[1]],int(m[2])).isoformat()
    except (TypeError,ValueError,KeyError):raise SchemaChangedError('Fecha de acción inválida.') from None
    ref=document_reference(downloads[0]['href'])
    if ref['rating_agency']!='Moody’s Local Perú':raise SchemaChangedError('PDF de otra clasificadora.')
    row=dict(action_id=key,article_url=requested,canonical_article_url=final,source_title=headings[0].get_text(' ',strip=True),
        listed_date=dated,pdf_url=ref['source_url'],document_id=ref['document_id'],source='Moody’s Local Perú',
        source_url=requested,retrieved_at=retrieved_at)
    return pd.DataFrame([row]).astype('string'),dict(html_sha256=sha256(html.encode()).hexdigest(),extraction_scope='dated_action_download_link')


class ActionIndexProvider(DatasetProvider):
    parser_version='2026-10-09.1';contract_version='1'
    def __init__(self,spec):
        super().__init__(spec);self.transport=CurlChromeTransport(state_dir=self.storage.state_root,timeout=180)
    def single_request(self,**_):return PeriodRequest(period_key='index',partition_key='indice',params={},mutable=True)
    def plan_sync(self,**_):yield self.single_request()
    def _fetch_period(self,request):
        r=self.transport.request('GET',INDEX)
        if r.url!=INDEX:raise SourceUnavailableError('Redirección del índice; captura rechazada.')
        d,m=parse_index(r.text,retrieved_at=utc_now_iso());return FetchResult(self.spec.dataset_id,d,metadata=m,raw=r.text)
    def filter_loaded(self,data,**_):return data.drop(columns=['_period_key'],errors='ignore')


class ActionReferencesProvider(ActionIndexProvider):
    def single_request(self,*,url=None,**_):
        key,canonical=action_reference(url)
        return PeriodRequest(period_key=key,partition_key='referencias',params={'url':canonical},mutable=True)
    def plan_sync(self,*,urls=None,**_):
        if urls is None:raise InvalidQueryError('Indique acciones explícitas.')
        seen=set()
        for url in [urls] if isinstance(urls,str) else urls:
            r=self.single_request(url=url)
            if r.period_key not in seen:seen.add(r.period_key);yield r
    def _fetch_period(self,request):
        r=self.transport.request('GET',request.params['url']);action_reference(r.url)
        d,m=parse_reference(r.text,url=request.params['url'],retrieved_at=utc_now_iso())
        if action_reference(r.url)[1] not in (request.params['url'],d.canonical_article_url.iloc[0]):
            raise SourceUnavailableError('Redirección a otra acción.')
        return FetchResult(self.spec.dataset_id,d,metadata=m,raw=r.text)
    def filter_loaded(self,data,*,urls=None,**_):
        if urls is not None:data=data[data.action_id.isin({action_reference(u)[0] for u in ([urls] if isinstance(urls,str) else urls)})]
        return data.drop(columns=['_period_key'],errors='ignore')
