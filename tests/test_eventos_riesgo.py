"""Test source-backed direction, conservative chronology and explicit PDF blocks."""
import pandas as pd
import pytest
from fuentes_financieras.eventos_riesgo import (
    SUMMARY_COLUMNS, summary_changes, document_changes, display_changes,
)


def rating(period='202501', value='B', **overrides):
    row = dict.fromkeys(SUMMARY_COLUMNS[:18], '')
    row.update(period_code=period, entity_type='Banco', entity_name='Entidad',
               rating_agency='Agencia', report_agency_code='001196',
               rating_kind='institutional_summary', rating=value,
               trend_basis='published_change_vs_previous_classification',
               report_id='001196:'+period+':1:1', report_url='https://example.test/'+period,
               source_url='https://example.test/resumen', retrieved_at='checked')
    row.update(overrides)
    return row


def test_changes_are_ordered_without_ranking_letters():
    observations, events = summary_changes(pd.DataFrame([
        rating('202601', 'B+', trend='down'), rating('202501', 'B'), rating('202502', 'A'),
    ]))
    assert list(observations.comparison_status) == ['first_observation', 'changed_text', 'changed_text']
    assert list(events.event_type) == ['observed_rating_change', 'published_downgrade']
    assert events.iloc[0].event_basis == 'consecutive_summary_cuts'
    assert events.iloc[1].previous_rating == 'A'
    assert events.iloc[1].event_basis == 'published_change_vs_previous_classification'
    assert events.iloc[1].previous_report_id == '001196:202502:1:1'


def test_gap_is_not_bridged_but_source_signal_is_preserved():
    observations, events = summary_changes(pd.DataFrame([rating(), rating('202601', 'A')]))
    assert observations.iloc[1].comparison_status == 'gap_not_compared' and events.empty
    _, events = summary_changes(pd.DataFrame([rating(), rating('202601', 'A', trend='up')]))
    assert events.iloc[0].event_type == 'published_upgrade'


def test_first_cut_can_have_published_direction_without_previous_selected_cut():
    observations, events = summary_changes(pd.DataFrame([rating(trend='down')]))
    assert observations.iloc[0].comparison_status == 'first_observation'
    assert events.iloc[0].previous_period_code == ''


def test_absent_unknown_and_equal_are_not_stable_or_withdrawal_events():
    _, events = summary_changes(pd.DataFrame([rating(), rating('202502', 'B', trend=pd.NA)]))
    assert events.empty
    observations, events = summary_changes(pd.DataFrame([rating(), rating('202502', None)]))
    assert observations.iloc[1].comparison_status == 'missing_value' and events.empty
    _, events = summary_changes(pd.DataFrame([rating(), rating('202502', 'B', trend='unknown', data_quality_flags='source_change_symbol_unrecognized')]))
    assert events.empty
    out = display_changes(observations)
    assert observations.iloc[1].comparison_status == 'missing_value'
    assert out.iloc[1].comparison_status == 'Valor ausente; no comparado'


@pytest.mark.parametrize('column,value', [('rating_agency','Otra'), ('entity_name','Entidad SA'), ('entity_type','Financiera'), ('report_agency_code','000409')])
def test_names_types_and_agencies_are_not_automatically_joined(column, value):
    observations, events = summary_changes(pd.DataFrame([rating(), rating('202502', 'A', **{column:value})]))
    assert observations.comparison_status.eq('first_observation').all() and events.empty


def test_duplicate_rows_are_rejected():
    with pytest.raises(ValueError, match='duplicadas'):
        summary_changes(pd.DataFrame([rating(), rating()]))


def field(role, value, **overrides):
    row = dict(report_id='id', period_code='202601', field_kind='outlook',
               temporal_role=role, value_raw=value, normalized_value=value,
               extraction_status='extracted', page_number=1, evidence_text=role+': '+value,
               pdf_sha256='digest', report_url='https://example.test/pdf', retrieved_at='checked')
    row.update(overrides)
    return row


def test_outlook_pairs_keep_page_and_do_not_propagate_to_deposits():
    fields = pd.DataFrame([field('previous','Negativa'),field('current','Estable'),
                           field('current','CP1',field_kind='short_term_deposits')])
    comparisons = document_changes(fields)
    outlook = comparisons[comparisons.field_kind.eq('outlook')].iloc[0]
    assert outlook.event_type == 'outlook_change' and outlook.previous_value == 'Negativa'
    assert outlook.previous_page_number == 1 and outlook.evidence_text == 'current: Estable'
    deposits = comparisons[comparisons.field_kind.eq('short_term_deposits')].iloc[0]
    assert deposits.comparison_status == 'no_previous_block' and deposits.event_type == ''


def test_rating_letter_change_has_no_inferred_direction():
    comparisons = document_changes(pd.DataFrame([
        field('previous','B',field_kind='financial_strength'),field('current','A',field_kind='financial_strength')]))
    assert comparisons.iloc[0].event_type == 'rating_text_change'
    assert comparisons.iloc[0].event_basis == 'explicit_document_blocks'


def test_ambiguous_or_reviewed_document_blocks_do_not_emit_event():
    a, b = field('previous','Negativa'), field('current','Estable')
    for records in ([a,b,b], [a,dict(b,extraction_status='needs_review')], [a,dict(b,pdf_sha256='other-digest')]):
        result = document_changes(pd.DataFrame(records))
        assert result.iloc[0].comparison_status == 'needs_review' and result.iloc[0].event_type == ''
    result = document_changes(pd.DataFrame([a,field('current','Negativa')]))
    assert result.iloc[0].comparison_status == 'equal_text' and result.iloc[0].event_type == ''
