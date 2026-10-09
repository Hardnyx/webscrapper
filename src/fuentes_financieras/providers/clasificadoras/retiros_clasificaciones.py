"""Verified agency PDFs and narrowly scoped withdrawal announcements."""
from datetime import date
from hashlib import sha256
from io import BytesIO
import os
import re
from urllib.parse import quote, unquote, urlsplit

import pandas as pd

from fuentes_financieras.exceptions import InvalidQueryError, SchemaChangedError, SourceUnavailableError
from fuentes_financieras.models import FetchResult, PeriodRequest
from fuentes_financieras.provider import DatasetProvider
from fuentes_financieras.storage import utc_now_iso
from fuentes_financieras.transports.curl_chrome import CurlChromeTransport

AGENCIES = {'moodyslocal.com.pe': 'Moody’s Local Perú', 'www.aai.com.pe': 'Apoyo & Asociados'}
MONTHS = dict(zip('enero febrero marzo abril mayo junio julio agosto septiembre octubre noviembre diciembre'.split(), range(1, 13)))
MONTHS['setiembre'] = 9
DATE = r'\d{1,2} de [a-z]+ de \d{4}'
ENTITY = r'[^();\n]{1,180}? S\.A(?:\.A|\.C)?\.'
AUTHORITY = r'Moody[’\x27]s Local PE Clasificadora de Riesgo S\.A\. \(en adelante, Moody[’\x27]s Local Perú\)'
MOODY_ENTITY = re.compile(r'afirma y retira la clasificación (?P<entity_rating>[A-C][+-]?) otorgada como Entidad a (?P<entity>'
    + ENTITY + r') \(en adelante, [^()]{1,120}\) y la clasificación (?P<issuer_rating>[A-C]{1,3}[+-]?\.pe) como Emisor\.', re.I)
MOODY_PROGRAM = re.compile(r'afirma la categoría [A-C][+-]? como Entidad a (?P<entity>' + ENTITY
    + r') \(en adelante, [^()]{1,120}\)\. Asimismo, afirma [^\n]{1,1600}?\. De otro lado, retira la clasificación del '
    + r'(?P<scope>(?:Primer|Segundo|Tercer|Cuarto|Quinto|Sexto|Séptimo|Octavo) Programa de Certificados de Depósito Negociables), '
    + r'debido a (?P<reason>su vencimiento)\. La Perspectiva es (?:Estable|Positiva|Negativa)\.', re.I)
AAI = re.compile(r'se pone en conocimiento público que se retiran las clasificaci[oó]nes de (?P<entity>[^,;():]{1,120}) '
    + r'de (?P<scope>institución de [A-C][+-]? y de sus instrumentos: depósitos de largo plazo de [A-C]{1,3}[+-]?\(pe\), '
    + r'depósitos de corto plazo de CP[123]\s*[+-]?\(pe\), certificados de depósitos negociables de CP[123]\s*[+-]?\(pe\), '
    + r'y bonos subordinados de [A-C]{1,3}[+-]?\(pe\)) por (?P<reason>resolución de contrato)\.', re.I)


def document_reference(url):
    if not isinstance(url, str):
        raise InvalidQueryError('Indique una URL oficial de comunicado PDF.')
    parsed = urlsplit(url)
    path = unquote(parsed.path)
    if (parsed.scheme != 'https' or parsed.netloc not in AGENCIES or parsed.query or parsed.fragment
            or not re.fullmatch(r'/wp-content/uploads/\d{4}/(?:0[1-9]|1[0-2])/[^/\\]{1,220}\.pdf', path)
            or any(c in path for c in '\x00\r\n') or '%' in path):
        raise InvalidQueryError('Solo PDF HTTPS de uploads oficiales de Moody’s Local Perú o Apoyo & Asociados, sin parámetros.')
    canonical = 'https://' + parsed.netloc + quote(path, safe='/.-_')
    return dict(document_id=sha256(canonical.encode('utf-8')).hexdigest(), document_filename=path.rsplit('/', 1)[1],
                rating_agency=AGENCIES[parsed.netloc], source_url=canonical)


def explicit_date(value):
    match = re.fullmatch(r'(\d{1,2}) de ([a-z]+) de (\d{4})', value, re.I)
    try:
        if not match:
            raise ValueError('date')
        return date(int(match[3]), MONTHS[match[2].lower()], int(match[1])).isoformat()
    except (ValueError, KeyError):
        raise SchemaChangedError('Fecha de comunicado inválida.') from None


def extract_withdrawal(text, *, agency):
    compact = re.sub(r'\s+', ' ', text).strip()
    row = dict(announcement_date='', event_type='', entity_name='', withdrawal_scope='', withdrawal_reason='',
               extraction_status='needs_review', evidence_text=compact[:2500], evidence_locator='PDF, página 1; formato no reconocido')
    if agency == 'Moody’s Local Perú':
        # The dated action block ends before the table, foundations and historical narrative.
        anchors = list(re.finditer(r'ACCIÓN DE CLASIFICACIÓN LIMA, PERÚ (?P<date>' + DATE + r') ' + AUTHORITY + r' ', compact, re.I))
        ends = list(re.finditer(r'La acción de clasificación se resume en el siguiente detalle:', compact, re.I))
        if len(anchors) != 1 or len(ends) != 1 or ends[0].start() <= anchors[0].end():
            return row
        anchor = anchors[0]
        lead = compact[anchor.end():ends[0].start()].strip()
        row.update(announcement_date=explicit_date(anchor['date']), evidence_text=compact[anchor.start():ends[0].start()].strip(),
                   evidence_locator='PDF, página 1; bloque fechado ACCIÓN DE CLASIFICACIÓN anterior a tabla')
        direct = MOODY_ENTITY.fullmatch(lead)
        program = MOODY_PROGRAM.fullmatch(lead)
        match = direct or program
        if match and len(re.findall(r'S\.A(?:\.A|\.C)?\.', match['entity'])) == 1:
            row.update(event_type='rating_withdrawal_announced', entity_name=match['entity'], extraction_status='extracted',
                       withdrawal_scope=('Entidad: ' + match['entity_rating'] + '; Emisor: ' + match['issuer_rating']) if direct else match['scope'],
                       withdrawal_reason='' if direct else match['reason'])
    elif agency == 'Apoyo & Asociados':
        anchors = list(re.finditer(r'Apoyo & Asociados Internacionales \(A&A\) [–-] Lima, (?P<date>' + DATE + r'): ', compact, re.I))
        ends = list(re.finditer(r'Contactos:', compact, re.I))
        if len(anchors) != 1 or len(ends) != 1 or ends[0].start() <= anchors[0].end():
            return row
        anchor = anchors[0]
        lead = compact[anchor.end():ends[0].start()].strip()
        row.update(announcement_date=explicit_date(anchor['date']), evidence_text=compact[anchor.start():ends[0].start()].strip(),
                   evidence_locator='PDF, página 1; comunicado fechado anterior a Contactos')
        match = AAI.fullmatch(lead)
        if match:
            row.update(event_type='rating_withdrawal_announced', entity_name=match['entity'], withdrawal_scope=match['scope'],
                       withdrawal_reason=match['reason'], extraction_status='extracted')
    return row


def parse_document(content, *, reference, retrieved_at):
    from pypdf import PdfReader
    if not content.startswith(b'%PDF-') or b'%%EOF' not in content[-2048:] or len(content) > 20_000_000:
        raise SourceUnavailableError('Respuesta no PDF, incompleta o mayor de 20 MB.')
    try:
        reader = PdfReader(BytesIO(content), strict=True)
        if reader.is_encrypted or not 1 <= len(reader.pages) <= 200:
            raise ValueError('PDF cifrado o número de páginas fuera de rango.')
        texts = [page.extract_text() or '' for page in reader.pages]
    except Exception as exc:
        raise SchemaChangedError('PDF ilegible, cifrado o estructura dañada.') from exc
    row = extract_withdrawal(texts[0], agency=reference['rating_agency'])
    if not texts[0].strip():
        row['extraction_status'] = 'needs_ocr'
    digest = sha256(content).hexdigest()
    record = dict(**reference, **row, page_number=1, pdf_sha256=digest, source=reference['rating_agency'], retrieved_at=retrieved_at)
    data = pd.DataFrame([record])
    for col in data:
        data[col] = data[col].astype('int64' if col == 'page_number' else 'string')
    return data, dict(pdf_sha256=digest, page_count=len(texts), text_pages=sum(bool(t.strip()) for t in texts),
                      extraction_scope='dated_first_page_action', extraction_status=row['extraction_status'])


class RatingWithdrawalsProvider(DatasetProvider):
    parser_version = '2026-10-09.1'
    contract_version = '1'

    def __init__(self, spec):
        super().__init__(spec)
        self.pdf_root = self.storage.root / 'documents'
        self.transport = CurlChromeTransport(state_dir=self.storage.state_root, timeout=180)

    def single_request(self, *, url=None, redownload=False, **_):
        reference = document_reference(url)
        return PeriodRequest(period_key=reference['document_id'], partition_key='comunicados',
                             params=dict(reference=reference, redownload=redownload))

    def plan_sync(self, *, urls=None, redownload=False, **_):
        if urls is None:
            raise InvalidQueryError('Indique urls; no se rastrean automáticamente los sitios de las clasificadoras.')
        seen = set()
        for url in [urls] if isinstance(urls, str) else urls:
            request = self.single_request(url=url, redownload=redownload)
            if request.period_key not in seen:
                seen.add(request.period_key)
                yield request

    def pdf_path(self, document_id):
        return self.pdf_root / (document_id + '.pdf')

    def pdf_matches(self, document_id, digest):
        try:
            return bool(digest and sha256(self.pdf_path(document_id).read_bytes()).hexdigest() == digest)
        except OSError:
            return False

    def _refresh_due(self, request, previous, refresh_hours):
        if request.params['redownload'] or not self.pdf_matches(request.period_key, previous.get('metadata', {}).get('pdf_sha256')):
            return True
        return super()._refresh_due(request, previous, refresh_hours)

    def _fetch_period(self, request):
        reference = request.params['reference']
        path = self.pdf_path(request.period_key)
        previous = self.storage.manifest.get(request.period_key) or {}
        expected = previous.get('metadata', {}).get('pdf_sha256')
        cached = self.pdf_matches(request.period_key, expected) and not request.params['redownload']
        if cached:
            content = path.read_bytes()
        else:
            response = self.transport.request('GET', reference['source_url'])
            if document_reference(response.url)['source_url'] != reference['source_url']:
                raise SourceUnavailableError('Redirección a otro comunicado; captura rechazada.')
            content = response.content
        data, metadata = parse_document(content, reference=reference, retrieved_at=utc_now_iso())
        if not cached:
            self.pdf_root.mkdir(parents=True, exist_ok=True)
            temporary = path.with_suffix('.pdf.tmp')
            temporary.write_bytes(content)
            os.replace(temporary, path)
        return FetchResult(self.spec.dataset_id, data, metadata=metadata)

    def filter_loaded(self, data, *, urls=None, **_):
        if urls is not None:
            identifiers = {document_reference(url)['document_id'] for url in ([urls] if isinstance(urls, str) else urls)}
            data = data[data.document_id.isin(identifiers)]
        return data.drop(columns=['_period_key'], errors='ignore')
