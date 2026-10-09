"""Explicit SBS news announcements, with conservative disposition extraction."""
import re
from datetime import date
from hashlib import sha256
from urllib.parse import parse_qs, urlsplit

import pandas as pd
from bs4 import BeautifulSoup

from fuentes_financieras.exceptions import InvalidQueryError, SchemaChangedError, SourceUnavailableError
from fuentes_financieras.models import FetchResult, PeriodRequest
from fuentes_financieras.provider import DatasetProvider
from fuentes_financieras.storage import utc_now_iso
from fuentes_financieras.transports.curl_chrome import CurlChromeTransport

MONTHS = dict(zip('enero febrero marzo abril mayo junio julio agosto septiembre octubre noviembre diciembre'.split(), range(1, 13)))
MONTHS['setiembre'] = 9
ACTOR = r'la Superintendencia de Banca, Seguros y AFP \(SBS\)'
ENTITY = r'(?P<entity>(?:Caja Municipal de Ahorro y Crédito|Caja Rural de Ahorro y Crédito|Financiera|Banco) [^();\n]{1,180}? S\.A(?:\.A)?\.)'
INTERVENTION = (
    r'(?:En resguardo de los intereses de los ahorristas|En cumplimiento de su mandato constitucional de cautelar los intereses del público ahorrista y la estabilidad del sistema financiero), '
    + ACTOR + r' (?:ha intervenido a la|ha sometido a Régimen de Intervención a) '
    + ENTITY + r'(?: \([^)]{1,60}\))? por (?:haber incurrido|el significativo deterioro)'
)
LIQUIDATION = (
    r'Con la finalidad de cautelar el valor de los activos de la ' + ENTITY
    + r'(?: \([^)]{1,60}\))? en Intervención, ' + ACTOR
    + r', mediante Resolución N\.°\s*(?P<resolution>\d{1,6}-\d{4}), ha dispuesto su disolución y el inicio del proceso de liquidación\.'
)


def announcement_reference(url):
    if not isinstance(url, str):
        raise InvalidQueryError('Indique una URL de noticia oficial SBS.')
    parsed = urlsplit(url)
    match = re.fullmatch(r'/noticia/detallenoticia/idnoticia/(\d{1,10})/?', parsed.path, re.I)
    query = parse_qs(parsed.query, keep_blank_values=True)
    if (parsed.scheme != 'https' or parsed.netloc != 'www.sbs.gob.pe' or not match
            or parsed.fragment or any(k.lower() != 'title' or len(v) != 1 for k, v in query.items())
            or (parsed.query and not query) or len(query) > 1):
        raise InvalidQueryError('Solo se admiten enlaces HTTPS oficiales de DetalleNoticia, con título opcional.')
    identifier = str(int(match[1]))
    if identifier == '0':
        raise InvalidQueryError('Identificador de noticia inválido.')
    return identifier, 'https://www.sbs.gob.pe/noticia/detallenoticia/idnoticia/' + identifier


def parse_announcement(html, *, url, retrieved_at):
    identifier, canonical = announcement_reference(url)
    if not re.search(r'</html\s*>', html, re.I):
        raise SchemaChangedError('HTML incompleto; captura rechazada.')
    if len(html.encode('utf-8')) > 2_000_000:
        raise SchemaChangedError('Página de noticia mayor de 2 MB.')
    soup = BeautifulSoup(html, 'lxml')
    titles = soup.select('h2.boletin-sala-prensa__title')
    bodies = soup.select('div.boletin-sala-prensa__text')
    if len(titles) != 1 or len(bodies) != 1:
        raise SchemaChangedError('Estructura de noticia ausente o ambigua.')
    title = re.sub(r'\s+', ' ', titles[0].get_text(' ', strip=True)).strip()
    paragraphs = [re.sub(r'\s+', ' ', p.get_text(' ', strip=True)).strip() for p in bodies[0].find_all('p')]
    paragraphs = [p for p in paragraphs if p]
    if not title or not paragraphs:
        raise SchemaChangedError('Noticia sin título o párrafo inicial.')
    lead = paragraphs[0]
    row = dict(announcement_id=identifier, announcement_date='', effective_date='',
               event_type='', entity_name='', resolution_number='', extraction_status='needs_review',
               source_title=title, evidence_text=lead, evidence_locator='div.boletin-sala-prensa__text > p:first-nonempty',
               html_sha256=sha256(html.encode('utf-8')).hexdigest(),
               source='SBS', source_url=canonical, retrieved_at=retrieved_at)
    dateline = re.match(r'Lima, (\d{1,2}) de ([a-z]+) de (\d{4})\s*\.\s*[-–]\s*', lead, re.I)
    if dateline:
        try:
            row['announcement_date'] = date(int(dateline[3]), MONTHS[dateline[2].lower()], int(dateline[1])).isoformat()
        except (ValueError, KeyError):
            raise SchemaChangedError('Fecha de anuncio inválida.') from None
        statement = lead[dateline.end():]
        for kind, pattern, title_marker in [
            ('intervention_announced', INTERVENTION, r'intervien|intervención'),
            ('dissolution_liquidation_announced', LIQUIDATION, r'disolución.*liquidación'),
        ]:
            match = re.match(pattern, statement, re.I)
            if match and re.search(title_marker, title, re.I):
                entity = match['entity']
                if len(re.findall(r'S\.A(?:\.A)?\.', entity)) != 1:
                    continue
                row.update(event_type=kind, entity_name=entity,
                           resolution_number=match.groupdict().get('resolution') or '',
                           extraction_status='extracted')
                break
    data = pd.DataFrame([row]).astype('string')
    return data, dict(announcement_id=identifier, html_sha256=row['html_sha256'],
                     extraction_scope='dated_first_paragraph', extraction_status=row['extraction_status'])


class RegulatoryAnnouncementsProvider(DatasetProvider):
    parser_version = '2026-10-09.1'
    contract_version = '1'

    def __init__(self, spec):
        super().__init__(spec)
        self.transport = CurlChromeTransport(state_dir=self.storage.state_root, timeout=180)

    def single_request(self, *, url=None, **_):
        identifier, canonical = announcement_reference(url)
        return PeriodRequest(period_key=identifier, partition_key='noticias',
                             params={'url': canonical}, mutable=True)

    def plan_sync(self, *, urls=None, **_):
        if urls is None:
            raise InvalidQueryError('Indique urls; no se rastrea automáticamente toda la sala de prensa.')
        seen = set()
        for url in [urls] if isinstance(urls, str) else urls:
            request = self.single_request(url=url)
            if request.period_key not in seen:
                seen.add(request.period_key)
                yield request

    def _fetch_period(self, request):
        response = self.transport.request('GET', request.params['url'])
        _, final_url = announcement_reference(response.url)
        if final_url != request.params['url']:
            raise SourceUnavailableError('Redirección a otra noticia; captura rechazada.')
        data, metadata = parse_announcement(response.text, url=final_url, retrieved_at=utc_now_iso())
        return FetchResult(self.spec.dataset_id, data, raw=response.text, metadata=metadata)

    def filter_loaded(self, data, *, urls=None, **_):
        if urls is not None:
            identifiers = {announcement_reference(url)[0] for url in ([urls] if isinstance(urls, str) else urls)}
            data = data[data.announcement_id.isin(identifiers)]
        return data.drop(columns=['_period_key'], errors='ignore')
