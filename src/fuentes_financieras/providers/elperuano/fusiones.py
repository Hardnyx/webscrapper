"""Explicit SBS merger dispositions in individually identified gazette HTML."""
import re
from datetime import date
from hashlib import sha256
from urllib.parse import urlsplit

import pandas as pd
from bs4 import BeautifulSoup

from fuentes_financieras.exceptions import InvalidQueryError, SchemaChangedError, SourceUnavailableError
from fuentes_financieras.models import FetchResult, PeriodRequest
from fuentes_financieras.provider import DatasetProvider
from fuentes_financieras.storage import utc_now_iso
from fuentes_financieras.transports.curl_chrome import CurlChromeTransport

BASE = 'https://busquedas.elperuano.pe'
MONTHS = dict(zip('enero febrero marzo abril mayo junio julio agosto septiembre octubre noviembre diciembre'.split(), range(1, 13)))
MONTHS['setiembre'] = 9
PARTY = r'[^,;()\n]{1,180}?(?:S\.A(?:\.A|\.C)?\.|EDPYME)'
ARTICLE = r'Artículo (?:Primero|Único|1[º°]?)\.\s*-\s*'
AUTHORIZATION = re.compile(ARTICLE + r'Autorizar (?:(?P<direct>la fusión por absorción de)|a) (?P<primary>' + PARTY
    + r')(?P<indirect>, la fusión por absorción)? con (?:la empresa )?(?P<other>' + PARTY + r')(?=[,. ])', re.I)
CLARIFICATION = re.compile(ARTICLE + r'Autorizar a (?P<primary>' + PARTY
    + r'), la aprobación de la aclaratoria de la minuta de fusión por absorción con (?P<other>' + PARTY
    + r') relativa a la fecha de entrada en vigencia de dicha fusión, estipulada el (?P<date>\d{1,2} de [a-z]+ de \d{4}), así como, autorizar la modificación del Estatuto social', re.I)


def resolution_reference(url):
    if not isinstance(url, str):
        raise InvalidQueryError('Indique un enlace oficial de dispositivo o visor HTML de El Peruano.')
    parsed = urlsplit(url)
    match = re.fullmatch(r'/(?:dispositivo/NL|api/visor_html)/(\d{7,10}-\d{1,3})/?', parsed.path)
    if (parsed.scheme != 'https' or parsed.netloc != 'busquedas.elperuano.pe'
            or not match or parsed.query or parsed.fragment):
        raise InvalidQueryError('Solo enlaces HTTPS oficiales de dispositivo NL o visor HTML, sin parámetros.')
    identifier = match[1]
    if any(int(part) == 0 for part in identifier.split('-')):
        raise InvalidQueryError('Identificador de dispositivo inválido.')
    return identifier, BASE + '/api/visor_html/' + identifier


def explicit_date(value):
    match = re.fullmatch(r'(\d{1,2}) de ([a-z]+) de (\d{4})', value, re.I)
    try:
        if not match:
            raise ValueError('date')
        return date(int(match[3]), MONTHS[match[2].lower()], int(match[1])).isoformat()
    except (ValueError, KeyError):
        raise SchemaChangedError('Fecha explícita de resolución inválida.') from None


def parse_resolution(html, *, url, retrieved_at):
    identifier, canonical = resolution_reference(url)
    if len(html.encode('utf-8')) > 2_000_000 or not re.search(r'</html\s*>', html, re.I):
        raise SchemaChangedError('HTML incompleto o mayor de 2 MB.')
    soup = BeautifulSoup(html, 'lxml')
    containers = soup.find_all('div', id='x' + identifier)
    stories = soup.select('div.story')
    if len(containers) != 1 or len(stories) != 1 or stories[0] not in containers[0].descendants:
        raise SchemaChangedError('Dispositivo ausente, distinto del solicitado o ambiguo.')
    story = stories[0]
    titles = story.select('h1.sumilla')
    headers = story.select('h2.resoluci-n')
    paragraphs = [re.sub(r'\s+', ' ', p.get_text(' ', strip=True)).strip()
                  for p in story.find_all('p', recursive=False)]
    paragraphs = [p for p in paragraphs if p]
    if len(titles) != 1 or len(headers) != 1 or not paragraphs or paragraphs[-1] != identifier:
        raise SchemaChangedError('Título, encabezado o cierre de dispositivo ausente o ambiguo.')
    title = re.sub(r'\s+', ' ', titles[0].get_text(' ', strip=True)).strip()
    header = re.sub(r'\s+', ' ', headers[0].get_text(' ', strip=True)).strip()
    resolution = re.fullmatch(r'RESOLUCIÓN SBS N[º°] (\d{1,6}-\d{4})', header, re.I)
    dateline = re.fullmatch(r'Lima, (\d{1,2} de [a-z]+ de \d{4})', paragraphs[0], re.I)
    if not title or not resolution or not dateline:
        raise SchemaChangedError('Se requiere una resolución SBS con fecha explícita en el encabezado.')
    resolved = [i for i, p in enumerate(paragraphs) if p == 'RESUELVE:']
    endings = [i for i, p in enumerate(paragraphs) if re.fullmatch(r'Regístrese, comuníquese y publíquese[.,]', p, re.I)]
    if len(resolved) != 1 or len(endings) != 1 or endings[0] <= resolved[0] + 1:
        raise SchemaChangedError('Bloque resolutivo incompleto o ambiguo.')
    # Only the first operative paragraph is eligible; recitals and later quotations cannot create an event.
    index = resolved[0] + 1
    evidence = paragraphs[index]
    row = dict(norm_id=identifier, resolution_number=resolution[1], resolution_date=explicit_date(dateline[1]),
               event_type='', entity_name='', counterparty_name='', absorbing_entity_name='', absorbed_entity_name='',
               roles_status='unspecified', effective_date='', effective_basis='not_extracted',
               effective_condition='', extraction_status='needs_review', source_title=title,
               evidence_text=evidence, evidence_locator=f'div#x{identifier} div.story > p:nonempty({index + 1})',
               html_sha256=sha256(html.encode('utf-8')).hexdigest(), source='El Peruano / SBS',
               source_url=canonical, retrieved_at=retrieved_at)
    if re.search(r'^Autorizan .*fusión por absorción', title, re.I):
        clarification = CLARIFICATION.match(evidence)
        authorization = AUTHORIZATION.match(evidence)
        if authorization and bool(authorization['direct']) == bool(authorization['indirect']):
            authorization = None
        if clarification and not re.search(r'aclaratoria', title, re.I):
            clarification = None
        match = clarification or authorization
        if (match and match['primary'].casefold() != match['other'].casefold()
                and all(match[name].casefold() in title.casefold() for name in ('primary', 'other'))
                and all(len(re.findall(r'S\.A(?:\.A|\.C)?\.|EDPYME', match[name], re.I)) == 1 for name in ('primary', 'other'))):
            row.update(entity_name=match['primary'], counterparty_name=match['other'], extraction_status='extracted',
                       event_type='merger_date_clarification' if clarification else 'merger_authorized')
            if clarification:
                row.update(effective_date=explicit_date(match['date']), effective_basis='date_in_clarification_article')
            else:
                row['effective_condition'] = evidence[match.end():].strip(' ,')
                if re.match(r', extinguiéndose esta última sin (?:disolverse ni )?liquidarse\.', evidence[match.end():], re.I):
                    row.update(absorbing_entity_name=match['primary'], absorbed_entity_name=match['other'], roles_status='operative_article')
    data = pd.DataFrame([row]).astype('string')
    return data, dict(norm_id=identifier, html_sha256=row['html_sha256'],
                     extraction_scope='first_operative_paragraph', extraction_status=row['extraction_status'])


class MergerResolutionsProvider(DatasetProvider):
    parser_version = '2026-10-09.1'
    contract_version = '1'

    def __init__(self, spec):
        super().__init__(spec)
        self.transport = CurlChromeTransport(state_dir=self.storage.state_root, timeout=180)

    def single_request(self, *, url=None, **_):
        identifier, canonical = resolution_reference(url)
        return PeriodRequest(period_key=identifier, partition_key='resoluciones', params={'url': canonical}, mutable=True)

    def plan_sync(self, *, urls=None, **_):
        if urls is None:
            raise InvalidQueryError('Indique urls; no se recorre automáticamente el archivo legal.')
        seen = set()
        for url in [urls] if isinstance(urls, str) else urls:
            request = self.single_request(url=url)
            if request.period_key not in seen:
                seen.add(request.period_key)
                yield request

    def _fetch_period(self, request):
        response = self.transport.request('GET', request.params['url'])
        _, final = resolution_reference(response.url)
        if final != request.params['url']:
            raise SourceUnavailableError('Redirección a otro dispositivo; captura rechazada.')
        data, metadata = parse_resolution(response.text, url=final, retrieved_at=utc_now_iso())
        return FetchResult(self.spec.dataset_id, data, metadata=metadata, raw=response.text)

    def filter_loaded(self, data, *, urls=None, **_):
        if urls is not None:
            identifiers = {resolution_reference(url)[0] for url in ([urls] if isinstance(urls, str) else urls)}
            data = data[data.norm_id.isin(identifiers)]
        return data.drop(columns=['_period_key'], errors='ignore')
