"""Verified PDF downloads and conservative, page-backed rating extraction."""
from datetime import date
from hashlib import sha256
from io import BytesIO
import os
import re

import pandas as pd

from fuentes_financieras.exceptions import InvalidQueryError, SchemaChangedError, SourceUnavailableError
from fuentes_financieras.models import FetchResult, PeriodRequest
from fuentes_financieras.provider import DatasetProvider
from fuentes_financieras.storage import utc_now_iso
from fuentes_financieras.transports.curl_chrome import CurlChromeTransport
from ._clasificaciones_parser import period_parts, report_reference

LABELS = {
    'Fortaleza Financiera': 'financial_strength',
    'Depósitos de Corto Plazo': 'short_term_deposits',
    'Depósitos de Mediano y Largo Plazo': 'medium_long_term_deposits',
    'Depósitos de Largo Plazo': 'long_term_deposits',
}
RATING = r'(?:PE)?(?:Categoría\s+[IVX]+|CP[123][+-]?|[ABCDEF]{1,3}[+-]?)'
MONTHS = dict(zip(('enero febrero marzo abril mayo junio julio agosto septiembre octubre noviembre diciembre').split(), range(1,13)))


def normalized_date(raw):
    value = raw.strip().lower()
    match = re.fullmatch(r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})', value)
    if match:
        day, month, year = map(int, match.groups())
    else:
        match = re.fullmatch(r'(\d{1,2}) de ([a-z]+) de (\d{4})', value)
        if not match or match[2] not in MONTHS:
            return ''
        day, month, year = int(match[1]), MONTHS[match[2]], int(match[3])
    try:
        return date(year, month, day).isoformat()
    except ValueError:
        return ''


def extract_cover_fields(text, *, agency_code):
    """Recognize only tested cover structures; never parse narrative ratings."""
    lines = [re.sub(r'\s+', ' ', line).strip() for line in text.splitlines() if line.strip()]
    fields = []
    def add(kind, label, value, evidence, *, role='current', status='extracted'):
        fields.append(dict(field_kind=kind, field_label=label, value_raw=value,
                           normalized_value=value, temporal_role=role,
                           extraction_status=status, evidence_text=evidence))
    if agency_code == '000410' and 'CALIFICACIÓN CREDITICIA' in lines and 'CALIFICACIÓN PERSPECTIVA' in lines:
        index=lines.index('CALIFICACIÓN PERSPECTIVA')
        if index+1<len(lines):
            match=re.fullmatch(r'([A-D][+−-]?) (Estable|Positiva|Negativa)',lines[index+1])
            if match:
                evidence='CALIFICACIÓN CREDITICIA\nCALIFICACIÓN PERSPECTIVA\n'+lines[index+1]
                add('credit_rating','Calificación crediticia',match[1],evidence)
                add('credit_rating_outlook','Perspectiva de calificación crediticia',match[2],evidence)
    if agency_code == '000409':
        # Current PCR cards precede the definitions and historical table.
        stop = next((i for i,line in enumerate(lines) if line.casefold() == 'significado de la calificación'), None)
        if stop is not None:
            for i,line in enumerate(lines[:stop]):
                if line in LABELS and i+1 < stop and re.fullmatch(RATING, lines[i+1]):
                    add(LABELS[line], line, lines[i+1], line+'\n'+lines[i+1])
    elif agency_code in ('000408', '001196'):
        header = 'Ratings Actual Anterior' if agency_code == '000408' else 'Rating Actual* Anterior**'
        if header in lines:
            start = lines.index(header)
            stop = next((i for i in range(start+1,len(lines)) if lines[i].startswith(('Con información', '*Información', 'Metodologías'))), len(lines))
            block = ' '.join(lines[start:stop])
            for label,kind in LABELS.items():
                pattern = re.escape(label)+r'\d*\s+('+RATING+r')\s+('+RATING+r')(?=\s|$)'
                for match in re.finditer(pattern, block):
                    evidence = header+'\n'+match[0]
                    add(kind,label,match[1],evidence)
                    add(kind,label,match[2],evidence,role='previous')
    if agency_code == '000406' and 'CLASIFICACIONES ACTUALES (*)' in lines and 'Clasificación Perspectiva' in lines:
        start=lines.index('Clasificación Perspectiva')
        stop=next((i for i in range(start+1,len(lines)) if lines[i].startswith('(*)')),len(lines))
        block=' '.join(lines[start+1:stop])
        for label,kind,pattern in [('Entidad','entity_rating',r'[A-C][+-]?'),
                                   ('Emisor','issuer_rating',r'[A-C]{1,3}[+-]?\.pe'),
                                   ('Depósitos de Corto Plazo','short_term_deposits',r'ML A-[123]\.pe')]:
            match=re.search(re.escape(label)+r'\s+('+pattern+r')\s+(Estable|Positiva|Negativa|-)(?=\s|$)',block)
            if match:
                add(kind,label,match[1],'CLASIFICACIONES ACTUALES (*)\n'+match[0])
                if match[2]!='-':
                    add(kind+'_outlook','Perspectiva de '+label,match[2],'CLASIFICACIONES ACTUALES (*)\n'+match[0])
        for line in lines:
            match=re.fullmatch(r'Fecha de (comité|publicación): (\d{1,2} de [a-z]+ de \d{4})',line,re.I)
            if match:
                kind='committee_date' if match[1].lower()=='comité' else 'publication_date'
                iso=normalized_date(match[2])
                add(kind,'Fecha de '+match[1],match[2],line,status='extracted' if iso else 'needs_review')
                fields[-1]['normalized_value']=iso
    if agency_code in ('000408','000409','001196'):
        for i,line in enumerate(lines):
            if line == 'Perspectiva' and i+1 < len(lines) and re.fullmatch(r'(?:Estable|Positiva|Negativa)\.?', lines[i+1]):
                add('outlook','Perspectiva',lines[i+1],line+'\n'+lines[i+1])
                fields[-1]['normalized_value']=lines[i+1].rstrip('.')
            elif agency_code == '001196':
                match = re.fullmatch(r'Perspectiva (Estable|Positiva|Negativa) (Estable|Positiva|Negativa)',line)
                if match:
                    add('outlook','Perspectiva',match[1],line)
                    add('outlook','Perspectiva',match[2],line,role='previous')
    if agency_code == '000409':
        for line in lines:
            match = re.search(r'Fecha de Comité:\s*(\d{1,2} de [a-zá]+ de \d{4})',line,re.I)
            if match:
                iso=normalized_date(match[1])
                add('committee_date','Fecha de Comité',match[1],match[0],status='extracted' if iso else 'needs_review')
                fields[-1]['normalized_value']=iso
    elif agency_code == '000408':
        for i,line in enumerate(lines):
            if 'Clasificaciones otorgadas en Comités de fecha' in line and i+1<len(lines):
                for raw in re.findall(r'\d{1,2}/\d{1,2}/\d{4}',lines[i+1]):
                    add('committee_date','Comités de fecha',raw,line+'\n'+lines[i+1],role='unspecified',status='needs_review')
                    fields[-1]['normalized_value']=normalized_date(raw)
    elif agency_code == '001196':
        for i,line in enumerate(lines):
            if line.startswith('*Información') and not line.startswith('**') and i+1<len(lines):
                match=re.fullmatch(r'Aprobado en comité de (\d{1,2}-\d{1,2}-\d{4})\.',lines[i+1])
                if match:
                    iso=normalized_date(match[1])
                    add('committee_date','Aprobado en comité',match[1],line+'\n'+lines[i+1],status='extracted' if iso else 'needs_review')
                    fields[-1]['normalized_value']=iso
    # Repeated values are not silently reconciled across potentially different instruments.
    for field in fields:
        key=(field['field_kind'],field['temporal_role'])
        if sum((f['field_kind'],f['temporal_role'])==key for f in fields)>1:
            field['extraction_status']='needs_review'
    return fields


def parse_document(content, *, reference, retrieved_at):
    from pypdf import PdfReader
    if not content.startswith(b'%PDF-') or b'%%EOF' not in content[-2048:] or len(content)>20_000_000:
        raise SourceUnavailableError('Respuesta no PDF, incompleta o mayor de 20 MB.')
    try:
        reader=PdfReader(BytesIO(content),strict=True)
        if reader.is_encrypted or not 1<=len(reader.pages)<=200:
            raise ValueError('PDF cifrado o número de páginas fuera de rango.')
        texts=[page.extract_text() or '' for page in reader.pages]
    except Exception as exc:
        raise SchemaChangedError('PDF ilegible, cifrado o estructura dañada.') from exc
    digest=sha256(content).hexdigest()
    fields=extract_cover_fields(texts[0],agency_code=reference['report_agency_code'])
    if not fields:
        status='needs_ocr' if not any(t.strip() for t in texts) else 'unsupported_cover'
        fields=[dict(field_kind='document',field_label='',value_raw='',normalized_value='',temporal_role='unspecified',
                     extraction_status=status,evidence_text='')]
    from ._evidencia_informes import extract_body_evidence
    body_fields,coverage=extract_body_evidence(texts,agency_code=reference['report_agency_code'])
    fields=[dict(page_number=1,**field) for field in fields]+body_fields
    for field in fields:
        field.setdefault('unit','')
        field.setdefault('top_depositors',None)
        field.setdefault('observation_period','')
        field.setdefault('denominator_basis','')
    year,semester,_,period=period_parts(reference['report_period_code'])
    records=[dict(**reference,period_code=reference['report_period_code'],period=period,year=year,semester=semester,
        **field,pdf_sha256=digest,source='SBS',source_url=reference['report_url'],retrieved_at=retrieved_at) for field in fields]
    data=pd.DataFrame(records)
    for col in data:data[col]=data[col].astype('Int64' if col=='top_depositors' else 'int64' if col in ('year','semester','page_number') else 'string')
    metadata=dict(pdf_sha256=digest,page_count=len(texts),text_pages=sum(bool(t.strip()) for t in texts),
                  extraction_scope='cover_and_body_evidence',coverage=coverage,field_counts=data.extraction_status.value_counts().to_dict())
    return data,metadata


class RiskDocumentsProvider(DatasetProvider):
    parser_version='2026-10-09.2'
    contract_version='2'

    def __init__(self,spec):
        super().__init__(spec)
        self.pdf_root=self.storage.root/'documents'
        self.transport=CurlChromeTransport(state_dir=self.storage.state_root,timeout=180)

    def single_request(self,*,url=None,redownload=False,**_):
        if not isinstance(url,str) or not url:
            raise InvalidQueryError('Indique una URL oficial de informe SBS.')
        from urllib.parse import parse_qs,urlsplit
        values=parse_qs(urlsplit(url).query).get('codPeriodo',[])
        if len(values)!=1:raise InvalidQueryError('URL sin período SBS único.')
        try:
            reference,_=report_reference(url,values[0],url)
            year,*_=period_parts(values[0])
        except (ValueError,SchemaChangedError) as exc:
            raise InvalidQueryError('URL de informe SBS no válida.') from exc
        return PeriodRequest(period_key=reference['report_id'],partition_key=f'year={year}',
                             params=dict(reference=reference,redownload=redownload))

    def plan_sync(self,*,urls=None,redownload=False,**_):
        if urls is None:raise InvalidQueryError('Indique urls; no se consulta automáticamente el inventario.')
        if isinstance(urls,str):urls=[urls]
        seen=set()
        for url in urls:
            request=self.single_request(url=url,redownload=redownload)
            if request.period_key not in seen:
                seen.add(request.period_key);yield request

    def pdf_path(self,report_id):
        return self.pdf_root/(report_id.replace(':','_')+'.pdf')

    def pdf_matches(self,report_id,digest):
        path=self.pdf_path(report_id)
        try:return bool(digest and path.exists() and sha256(path.read_bytes()).hexdigest()==digest)
        except OSError:return False

    def _refresh_due(self,request,previous,refresh_hours):
        if request.params['redownload'] or not self.pdf_matches(request.period_key,previous.get('metadata',{}).get('pdf_sha256')):
            return True
        return super()._refresh_due(request,previous,refresh_hours)

    def _fetch_period(self,request):
        reference=request.params['reference'];path=self.pdf_path(request.period_key)
        previous=self.storage.manifest.get(request.period_key) or {}
        expected=previous.get('metadata',{}).get('pdf_sha256')
        content=path.read_bytes() if path.exists() else b''
        cached=bool(expected and sha256(content).hexdigest()==expected and not request.params['redownload'])
        if not cached:
            response=self.transport.request('GET',reference['report_url'])
            from urllib.parse import urlsplit
            final=urlsplit(response.url)
            if final.scheme!='https' or final.netloc!='extranet.sbs.gob.pe':
                raise SourceUnavailableError('Redirección del PDF fuera del destino oficial.')
            content=response.content
        data,metadata=parse_document(content,reference=reference,retrieved_at=utc_now_iso())
        if not cached:
            self.pdf_root.mkdir(parents=True,exist_ok=True)
            temporary=path.with_suffix('.pdf.tmp');temporary.write_bytes(content);os.replace(temporary,path)
        return FetchResult(self.spec.dataset_id,data,metadata=metadata)

    def filter_loaded(self,data,*,report_ids=None,**_):
        if report_ids is None:return data.drop(columns=['_period_key'],errors='ignore')
        if isinstance(report_ids,str):report_ids=[report_ids]
        return data[data.report_id.isin(report_ids)].drop(columns=['_period_key'],errors='ignore')
