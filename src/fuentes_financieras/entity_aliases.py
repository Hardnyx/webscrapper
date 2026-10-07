"""Reviewed SBS source-label equivalences, not authorization records."""
from .providers.sbs._universe_parser import normalized_entity_name

UNIVERSE_URL = 'https://www.sbs.gob.pe/app/pp/empresasweb/Paginas/EmpCaptarDep.aspx'
RATES_URL = 'https://www.sbs.gob.pe/app/pp/EstadisticasSAEEPortal/Paginas/TIPasivaDepositoEmpresa.aspx?tip='

# B/F labels were checked for 2026-10-06, R labels for 2026-08.
# Bounds deliberately cover only the reviewed sample; older/future periods
# require reviewed extensions rather than assuming historical continuity.
REVIEWED_LABELS = {
    'B': {
        'Alfin': 'ALFIN BANCO', 'BCI': 'BANCO BCI', 'BIF': 'BANBIF',
        'Bank of China': 'BANK OF CHINA (PERU)', 'Citibank': 'CITIBANK DEL PERU',
        'Compartamos': 'COMPARTAMOS BANCO', 'Crédito': 'BANCO DE CREDITO',
        'Efectiva': 'BANCO EFECTIVA', 'Falabella': 'BANCO FALABELLA',
        'GNB': 'BANCO GNB', 'ICBC': 'ICBC PERU BANK S.A.',
        'Pichincha': 'BANCO PICHINCHA', 'Ripley': 'BANCO RIPLEY',
        'Santander': 'SANTANDER PERU', 'Santander Cons. Bank': 'BN. SANTANDER CONS.',
        'Scotiabank': 'SCOTIABANK PERU',
    },
    'F': {
        'Confianza': 'FINANCIERA CONFIANZA', 'InFinance XP': 'INFINANCE XP S.A.',
        'Proempresa': 'FINANC. PROEMPRESA', 'Qapaq': 'FINANCIERA QAPAQ',
        'Surgir': 'FINANCIERA SURGIR',
    },
    'R': {
        'Cencosud Scotia': 'CRAC CENCOSUD SCOTIA', 'Incasur': 'CRAC INCASUR',
        'Los Andes': 'CRAC LOS ANDES', 'Prymera': 'CRAC PRYMERA',
    },
}
SUPPLEMENTARY_EVIDENCE = {
    ('B', 'BIF'): 'https://www.banbif.com.pe/Portals/0/PDF/Qui%C3%A9nes%20Somos/MEMORIAANUAL2015.pdf',
    ('B', 'Santander Cons. Bank'): 'https://www.santanderconsumer.com.pe/personas/contactanos',
}


def reviewed_aliases(catalog):
    records = {(r['entity_type_code'], r['normalized_name']): r for r in catalog.records}
    aliases = []
    for type_code, pairs in REVIEWED_LABELS.items():
        for label, target in pairs.items():
            record = records.get((type_code, normalized_entity_name(target)))
            if record is None:
                continue
            evidence = (f'Revisión de etiquetas SBS 2026-10-07: {label} -> {target}; '
                        f'{RATES_URL}{type_code}; {UNIVERSE_URL}')
            if (type_code, label) in SUPPLEMENTARY_EVIDENCE:
                evidence += '; ' + SUPPLEMENTARY_EVIDENCE[type_code, label]
            aliases.append({
                'dataset': 'pe.sbs.tasas_pasivas', 'entity_type_code': type_code,
                'alias': label, 'entity_id': record['entity_id'],
                'valid_from': '2026-08-01' if type_code == 'R' else '2026-10-06',
                'valid_to': '2026-08-31' if type_code == 'R' else '2026-10-07',
                'evidence': evidence,
            })
    return aliases
