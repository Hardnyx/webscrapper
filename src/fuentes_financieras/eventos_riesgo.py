"""Evidence-based changes from explicit consumers of independent SBS datasets."""
import pandas as pd
from .providers.sbs._clasificaciones_parser import period_parts

SUMMARY_COLUMNS = [
    'period_code', 'period_date', 'entity_type', 'entity_name', 'rating_agency',
    'report_agency_code', 'rating_kind', 'rating', 'trend', 'trend_basis',
    'source_change_icon_url', 'source_change_title', 'source_change_alt',
    'data_quality_flags', 'report_id', 'report_url', 'source_url', 'retrieved_at',
    'previous_period_code', 'previous_rating', 'previous_report_id',
    'previous_report_url', 'previous_source_url', 'previous_retrieved_at',
    'comparison_status', 'event_type', 'event_basis',
]
DOCUMENT_COLUMNS = [
    'report_id', 'period_code', 'field_kind', 'previous_value', 'current_value',
    'previous_page_number', 'page_number', 'previous_evidence_text', 'evidence_text',
    'pdf_sha256', 'report_url', 'retrieved_at', 'comparison_status', 'event_type', 'event_basis',
]
RATING_FIELDS = {
    'financial_strength', 'entity_rating', 'issuer_rating', 'credit_rating',
    'short_term_deposits', 'medium_long_term_deposits', 'long_term_deposits',
    'outlook', 'entity_rating_outlook', 'issuer_rating_outlook', 'credit_rating_outlook',
    'short_term_deposits_outlook',
}
DISPLAY_VALUES = {
    'first_observation': 'Primera observación de la selección',
    'gap_not_compared': 'Cortes no consecutivos; no comparados',
    'equal_text': 'Mismo texto; no implica estabilidad crediticia',
    'changed_text': 'Texto distinto', 'missing_value': 'Valor ausente; no comparado',
    'needs_review': 'Requiere revisión', 'no_previous_block': 'Sin bloque anterior reconocido',
    'no_current_block': 'Sin bloque actual reconocido',
    'published_upgrade': 'Subida publicada por SBS', 'published_downgrade': 'Bajada publicada por SBS',
    'observed_rating_change': 'Cambio de texto entre cortes; dirección sin determinar',
    'rating_text_change': 'Cambio de rating entre bloques; dirección sin determinar',
    'outlook_change': 'Cambio de perspectiva entre bloques',
    'published_change_vs_previous_classification': 'Símbolo SBS respecto de la clasificación anterior; no necesariamente del corte anterior',
    'consecutive_summary_cuts': 'Comparación literal de cortes SBS consecutivos',
    'explicit_document_blocks': 'Bloques Actual/Anterior del mismo PDF; sin fecha de evento inferida',
    'institutional_summary': 'Clasificación institucional del resumen', 'up': 'Subió', 'down': 'Bajó',
}


def _semester_index(code):
    year, semester, _, _ = period_parts(str(code))
    return year * 2 + semester


def summary_changes(history):
    """Do not rank rating letters or connect names, categories or agencies by aliases."""
    if history.empty:
        empty = pd.DataFrame(columns=SUMMARY_COLUMNS)
        return empty, empty.copy()
    required = set(SUMMARY_COLUMNS[:18])
    if missing := required - set(history):
        raise ValueError('Faltan columnas del contrato de clasificaciones: ' + ', '.join(sorted(missing)))
    data = history.copy()
    data['_semester'] = data.period_code.map(_semester_index)
    keys = ['entity_type', 'entity_name', 'rating_agency', 'report_agency_code', 'rating_kind']
    if data.duplicated(['period_code', *keys]).any():
        raise ValueError('Clasificaciones duplicadas; comparación ambigua.')
    if not data.rating_kind.eq('institutional_summary').all():
        raise ValueError('Se requieren clasificaciones institucionales del resumen.')
    rows = []
    for _, frame in data.groupby(keys, dropna=False, sort=False):
        previous = None
        for record in frame.sort_values('_semester').to_dict('records'):
            row = {column: record[column] for column in SUMMARY_COLUMNS[:18]}
            row.update({column: '' for column in SUMMARY_COLUMNS[18:]})
            row['comparison_status'] = 'first_observation'
            if previous is not None:
                for dest, src in [('previous_period_code', 'period_code'), ('previous_rating', 'rating'),
                                  ('previous_report_id', 'report_id'), ('previous_report_url', 'report_url'),
                                  ('previous_source_url', 'source_url'), ('previous_retrieved_at', 'retrieved_at')]:
                    row[dest] = previous[src]
                if record['_semester'] - previous['_semester'] != 1:
                    row['comparison_status'] = 'gap_not_compared'
                elif pd.isna(record['rating']) or pd.isna(previous['rating']) or not str(record['rating']).strip() or not str(previous['rating']).strip():
                    row['comparison_status'] = 'missing_value'
                else:
                    row['comparison_status'] = 'equal_text' if record['rating'] == previous['rating'] else 'changed_text'
            # The source icon refers to the previous classification, not our previous selected cut.
            if pd.notna(record['trend']) and record['trend'] in ('up', 'down'):
                row['event_type'] = 'published_upgrade' if record['trend'] == 'up' else 'published_downgrade'
                row['event_basis'] = 'published_change_vs_previous_classification'
            elif row['comparison_status'] == 'changed_text':
                row['event_type'] = 'observed_rating_change'
                row['event_basis'] = 'consecutive_summary_cuts'
            rows.append(row)
            previous = record
    observations = pd.DataFrame(rows, columns=SUMMARY_COLUMNS).sort_values(['period_code', 'entity_type', 'entity_name', 'rating_agency']).reset_index(drop=True)
    return observations, observations[observations.event_type.ne('')].reset_index(drop=True)


def document_changes(fields):
    """Compare only explicit current/previous fields within a single document."""
    rows = []
    if fields.empty:
        return pd.DataFrame(columns=DOCUMENT_COLUMNS)
    selected = fields[fields.field_kind.isin(RATING_FIELDS) & fields.temporal_role.isin(['current', 'previous'])]
    for (report_id, kind), frame in selected.groupby(['report_id', 'field_kind'], sort=True):
        current = frame[frame.temporal_role.eq('current')]
        previous = frame[frame.temporal_role.eq('previous')]
        source = frame.iloc[0]
        row = dict.fromkeys(DOCUMENT_COLUMNS, '')
        row.update(report_id=report_id, field_kind=kind, period_code=source.period_code,
                   pdf_sha256=source.pdf_sha256, report_url=source.report_url, retrieved_at=source.retrieved_at)
        for subset, prefix in [(current, ''), (previous, 'previous_')]:
            if len(subset) == 1:
                record = subset.iloc[0]
                row['current_value' if not prefix else 'previous_value'] = record.value_raw
                row[prefix + 'page_number'] = record.page_number
                row[prefix + 'evidence_text'] = record.evidence_text
        if len(current) > 1 or len(previous) > 1 or frame.pdf_sha256.nunique(dropna=False) != 1 or not frame.extraction_status.eq('extracted').all():
            row['comparison_status'] = 'needs_review'
        elif current.empty:
            row['comparison_status'] = 'no_current_block'
        elif previous.empty:
            row['comparison_status'] = 'no_previous_block'
        elif any(pd.isna(value) or not str(value).strip() for value in
                 [row['current_value'], row['previous_value'], current.iloc[0].normalized_value, previous.iloc[0].normalized_value]):
            row['comparison_status'] = 'missing_value'
        else:
            changed = current.iloc[0].normalized_value != previous.iloc[0].normalized_value
            row['comparison_status'] = 'changed_text' if changed else 'equal_text'
            row['event_basis'] = 'explicit_document_blocks'
            if changed:
                row['event_type'] = 'outlook_change' if 'outlook' in kind else 'rating_text_change'
        rows.append(row)
    return pd.DataFrame(rows, columns=DOCUMENT_COLUMNS)


def display_changes(frame):
    """Translate report semantics without translating original source evidence."""
    out = frame.copy()
    for column in ('comparison_status', 'event_type', 'event_basis', 'trend', 'trend_basis', 'rating_kind'):
        if column in out:
            out[column] = out[column].replace(DISPLAY_VALUES)
    if 'field_kind' in out:
        out['field_kind'] = out.field_kind.replace({
            'financial_strength': 'Fortaleza financiera', 'entity_rating': 'Clasificación de entidad',
            'issuer_rating': 'Clasificación de emisor', 'credit_rating': 'Calificación crediticia',
            'short_term_deposits': 'Depósitos de corto plazo', 'medium_long_term_deposits': 'Depósitos de mediano y largo plazo',
            'long_term_deposits': 'Depósitos de largo plazo', 'outlook': 'Perspectiva del informe',
            'entity_rating_outlook': 'Perspectiva de entidad', 'issuer_rating_outlook': 'Perspectiva de emisor',
            'credit_rating_outlook': 'Perspectiva de calificación crediticia', 'short_term_deposits_outlook': 'Perspectiva de depósitos de corto plazo',
        })
    return out
