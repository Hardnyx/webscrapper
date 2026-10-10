"""Dated deposit concentration and reviewable qualitative evidence from PDF text."""
import re

TOPICS = {
    'strategy': r'estrategia|plan estrat[eé]gico|modelo de negocio',
    'ownership_support': r'accionistas?|respaldo de su principal|soporte patrimonial|grupo controlador',
    'funding_cost': r'costo de fondeo|costo de fondos|costos? de captaci[oó]n',
    'risk_drivers': r'limitantes|debilidades|vulnerabilidades|riesgo de (?:liquidez|cr[eé]dito|mercado|concentraci[oó]n)',
    'event_mentions': r'intervenci[oó]n|fusiones?|fusi[oó]n|absorci[oó]n|retiro de (?:rating|clasificaci[oó]n)|sanciones?',
    'deposit_concentration': r'(?:principales|mayores) depositantes|concentraci[oó]n de depositantes',
}
MONTHS = dict(zip('enero febrero marzo abril mayo junio julio agosto septiembre octubre noviembre diciembre'.split(),range(1,13)))
MONTHS['setiembre']=9


def concentration_fields(text, *, agency_code):
    """Extract tested dated sentences, retaining denominator and date limits."""
    found=[]
    patterns=[]
    if agency_code=='001196':
        patterns=[(r'al cierre de ([a-z]+) de (\d{4}),? los (\d+) principales depositantes representaron el ([\d.,]+)% del total de depósitos(?: d\s*el Banco)?, mientras que los (\d+) principales concentraron el ([\d.,]+)%', 'total_deposits')]
    elif agency_code=='000406':
        patterns=[(r'concentración de depositantes \(([\d.,]+)% los (\d+) principales a ([a-z]+) de (\d{4})\)', 'unspecified_in_excerpt')]
    for pattern,basis in patterns:
        for match in re.finditer(pattern,text,re.I):
            if agency_code=='001196':month,year,n1,v1,n2,v2=match.groups();pairs=[(n1,v1),(n2,v2)]
            else:v,n,month,year=match.groups();pairs=[(n,v)]
            if month.lower() not in MONTHS or not 1900<=int(year)<=2100:continue
            for n,raw in pairs:
                # A mixed decimal/thousands separator is ambiguous in a percentage.
                if not re.fullmatch(r'\d+(?:[.,]\d+)?', raw):continue
                value=float(raw.replace(',','.'))
                if not 0<=value<=100 or not 1<=int(n)<=1000:continue
                found.append(dict(field_kind='deposit_concentration',field_label='Principales depositantes',
                    value_raw=raw+'%',normalized_value=str(value),unit='percent',top_depositors=int(n),
                    observation_period=f'{year}-{MONTHS[month.lower()]:02d}',denominator_basis=basis,
                    temporal_role='dated_observation',extraction_status='extracted',evidence_text=match[0]))
    if agency_code == '000406':
        pattern = (r'Cabe indicar que, a nivel de concentración de principales depositantes '
                   r'se evidencia un incremento en la concentración de los (\d+) principales '
                   r'depositantes al pasar a ([\d.,]+)% al cierre de (\d{4}), '
                   r'desde ([\d.,]+)% al término de (\d{4})\.')
        for match in re.finditer(pattern, text, re.I):
            n, current, year, previous, previous_year = match.groups()
            if not 1 <= int(n) <= 1000 or not 1900 <= int(previous_year) < int(year) <= 2100:
                continue
            # Year-end observations explicitly mean December; no document date is borrowed.
            for raw, observed_year in ((current, year), (previous, previous_year)):
                if not re.fullmatch(r'\d+(?:[.,]\d+)?', raw):
                    continue
                value = float(raw.replace(',', '.'))
                if not 0 <= value <= 100:
                    continue
                found.append(dict(field_kind='deposit_concentration', field_label='Principales depositantes',
                    value_raw=raw+'%', normalized_value=str(value), unit='percent', top_depositors=int(n),
                    observation_period=f'{observed_year}-12', denominator_basis='unspecified_in_excerpt',
                    temporal_role='dated_observation', extraction_status='extracted', evidence_text=match[0]))
    if agency_code == '001196':
        pattern = (r'Los (\d+) principales depositantes representaron el ([\d.,]+)% del total de depósitos '
                   r'de la CMAC Huancayo\s*, mientras que los (\d+) principales concentraron el ([\d.,]+)%, '
                   r'consolidando la tendencia decreciente a lo largo del periodo bajo análisis '
                   r'\(([\d.,]+)% y ([\d.,]+)% respectivamente al cierre de ([a-z]+) (?:de )?(\d{4})\s*\)')
        for match in re.finditer(pattern, text, re.I):
            n1, current1, n2, current2, previous1, previous2, month, year = match.groups()
            # "Respectively" binds the dated values to the preceding ordered depositor counts.
            if month.lower() not in MONTHS or not 1900 <= int(year) <= 2100:
                continue
            if not 1 <= int(n1) < int(n2) <= 1000:
                continue
            values = (current1, current2, previous1, previous2)
            if any(not re.fullmatch(r'\d+(?:[.,]\d+)?', raw) for raw in values):
                continue
            if any(not 0 <= float(raw.replace(',', '.')) <= 100 for raw in values):
                continue
            # Current values have no explicit date in this sentence and stay review candidates.
            for n, raw in ((n1, previous1), (n2, previous2)):
                found.append(dict(field_kind='deposit_concentration', field_label='Principales depositantes',
                    value_raw=raw+'%', normalized_value=str(float(raw.replace(',', '.'))), unit='percent',
                    top_depositors=int(n), observation_period=f'{year}-{MONTHS[month.lower()]:02d}',
                    denominator_basis='total_deposits', temporal_role='dated_observation',
                    extraction_status='extracted', evidence_text=match[0]))
    return found


def ownership_support_fields(text, *, agency_code):
    """Recognize complete published blocks without resolving legal ownership or guarantees."""
    found = []
    if agency_code == '001196':
        name = r"([A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúñ(). &'\-]{1,179}?)"
        integer = r'(\d{1,3}(?:,\d{3})+|\d+)'
        percent = r'(\d+(?:\.\d+)?)'
        pattern = (r'Accionistas Acciones Participación \(%\) ' + name + ' ' + integer + ' ' + percent + r'% '
                   + name + ' ' + integer + ' ' + percent + r'% Total ' + integer + ' ' + percent
                   + r'% (?=Miembros del Directorio Cargo Condición)')
        matches = list(re.finditer(pattern, text))
        if len(matches) == 1:
            match = matches[0]
            owner1, count1, percent1, owner2, count2, percent2, total, total_percent = match.groups()
            counts = [int(raw.replace(',', '')) for raw in (count1, count2, total)]
            percentages = [float(raw) for raw in (percent1, percent2, total_percent)]
            valid = (owner1 != owner2 and counts[0] > 0 and counts[1] > 0
                     and sum(counts[:2]) == counts[2] and percentages[2] == 100
                     and all(0 < value <= 100 for value in percentages[:2])
                     and abs(sum(percentages[:2]) - 100) <= 0.011
                     and all(abs(count / counts[2] * 100 - value) <= 0.0051
                             for count, value in zip(counts[:2], percentages[:2])))
            if valid:
                for owner, count, raw_percent in ((owner1, count1, percent1), (owner2, count2, percent2)):
                    for kind, raw, value, unit, basis in (
                        ('shareholder_shares', count, str(int(count.replace(',', ''))), 'shares', ''),
                        ('shareholder_participation', raw_percent+'%', str(float(raw_percent)), 'percent', 'reported_total_shares'),
                    ):
                        found.append(dict(field_kind=kind, field_label=owner, value_raw=raw,
                            normalized_value=value, unit=unit, denominator_basis=basis,
                            temporal_role='unspecified', extraction_status='extracted', evidence_text=match[0]))
    if agency_code == '000406':
        pattern = (r'→ Respaldo de su principal accionista, lo cual se ha visto reflejado en la '
                   r'capitalización de utilidades \(([\d.,]+)% de capitalización en el (\d{4}) '
                   r'correspondiente a las utilidades del ejercicio (\d{4})\)\.')
        for match in re.finditer(pattern, text):
            raw, year, earnings_year = match.groups()
            if not re.fullmatch(r'\d+(?:[.,]\d+)?', raw):
                continue
            value = float(raw.replace(',', '.'))
            if not 0 <= value <= 100 or not 1900 <= int(earnings_year) < int(year) <= 2100:
                continue
            found.append(dict(field_kind='earnings_capitalization', field_label='Utilidades capitalizadas',
                value_raw=raw+'%', normalized_value=str(value), unit='percent',
                observation_period=year, denominator_basis='earnings_year_'+earnings_year,
                temporal_role='annual_observation', extraction_status='extracted', evidence_text=match[0]))
    return found


def extract_body_evidence(texts, *, agency_code):
    fields=[];counts={topic:0 for topic in TOPICS};seen=set()
    for page_number,raw in enumerate(texts,1):
        text=re.sub(r'\s+',' ',raw).strip()
        for field in concentration_fields(text,agency_code=agency_code):
            fields.append(dict(page_number=page_number,**field))
        for field in ownership_support_fields(text,agency_code=agency_code):
            fields.append(dict(page_number=page_number,**field))
        for topic,pattern in TOPICS.items():
            for match in re.finditer(pattern,text,re.I):
                if counts[topic]>=2:break
                start=max(0,match.start()-140);end=min(len(text),match.end()+260)
                excerpt=text[start:end]
                key=(topic,excerpt)
                if key in seen:continue
                seen.add(key);counts[topic]+=1
                fields.append(dict(page_number=page_number,field_kind='qualitative_'+topic,
                    field_label=match[0],value_raw=excerpt,normalized_value='',temporal_role='unspecified',
                    extraction_status='needs_review',evidence_text=excerpt))
    coverage=[]
    has_text=any(t.strip() for t in texts)
    for topic in TOPICS:
        kinds = {'deposit_concentration': {'deposit_concentration'},
                 'ownership_support': {'shareholder_shares', 'shareholder_participation', 'earnings_capitalization'}}
        recognized = sum(f['field_kind'] in kinds.get(topic, set()) for f in fields)
        coverage.append(dict(topic=topic,recognized_count=recognized,candidate_count=counts[topic],
            coverage_status='needs_ocr' if not has_text else 'recognized' if recognized else 'needs_review' if counts[topic] else 'not_found_in_text',
            pages_checked=len(texts),text_pages=sum(bool(t.strip()) for t in texts)))
    return fields,coverage
