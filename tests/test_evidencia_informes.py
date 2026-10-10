"""Keep dated figures separate from unconfirmed narrative candidates."""
import pytest
from fuentes_financieras.providers.sbs._evidencia_informes import (
    concentration_fields, extract_body_evidence,
)
from fuentes_financieras.providers.sbs.documentos_riesgo import extract_cover_fields

JCR = ('Al cierre de diciembre de 2025 los 10 principales depositantes '
       'representaron el 9.7% del total de depósitos, mientras que los 20 '
       'principales concentraron el 13.2%. A junio representaron 14.7% y 17.7%.')
MOODYS_YEAR_END = ('Cabe indicar que, a nivel de concentración de principales depositantes '
                  'se evidencia un incremento en la concentración de los 20 principales '
                  'depositantes al pasar a 22.30% al cierre de 2025, desde 20.72% '
                  'al término de 2024.')
HUANCAYO = ('Los 10 principales depositantes representaron el 1.7% del total de depósitos de la CMAC Huancayo , '
            'mientras que los 20 principales concentraron el 2.4%, consolidando la tendencia decreciente '
            'a lo largo del periodo bajo análisis (2.4% y 3.2% respectivamente al cierre de diciembre 2024 ), '
            'ubicándose por debajo de las señales internas de alerta, apetito y capacidad.')


def test_jcr_historical_respectively_keeps_only_explicitly_dated_values():
    fields, coverage = extract_body_evidence(['Portada', HUANCAYO], agency_code='001196')
    figures = [f for f in fields if f['field_kind'] == 'deposit_concentration']
    assert [(f['top_depositors'], f['normalized_value']) for f in figures] == [(10, '2.4'), (20, '3.2')]
    assert all(f['observation_period'] == '2024-12' and f['page_number'] == 2
               and f['denominator_basis'] == 'total_deposits' for f in figures)
    assert any(f['extraction_status'] == 'needs_review' and '1.7%' in f['evidence_text'] for f in fields)
    assert next(c for c in coverage if c['topic'] == 'deposit_concentration')['recognized_count'] == 2
    assert not concentration_fields(HUANCAYO, agency_code='000406')


@pytest.mark.parametrize('old,new', [
    ('respectivamente', ''), ('diciembre 2024', ''), ('diciembre', 'desconocido'),
    ('2024', '2101'), ('3.2%', '103.2%'), ('1.7%', '1,7.2%'),
    ('Los 10', 'Los 20'), ('20 principales', '5 principales'),
    ('total de depósitos', 'total de activos'), ('representaron', 'no representaron'),
    ('representaron', 'representarían'),
])
def test_jcr_historical_comparison_rejects_missing_binding_and_ambiguity(old, new):
    assert not concentration_fields(HUANCAYO.replace(old, new), agency_code='001196')


def test_historical_jcr_bank_denominator_with_pdf_spacing():
    text = JCR.replace('diciembre', 'junio').replace('del total de depósitos,', 'del total de depósitos d el Banco,')
    fields = concentration_fields(text, agency_code='001196')
    assert len(fields) == 2
    assert all(f['observation_period'] == '2025-06' and f['denominator_basis'] == 'total_deposits' for f in fields)
    assert not concentration_fields(text.replace('del total de depósitos d el Banco', 'del total de créditos'), agency_code='001196')


def test_moodys_year_end_preserves_two_explicit_observations():
    fields = concentration_fields(MOODYS_YEAR_END, agency_code='000406')
    assert [(f['observation_period'], f['normalized_value']) for f in fields] == [('2025-12', '22.3'), ('2024-12', '20.72')]
    assert all(f['top_depositors'] == 20 and f['denominator_basis'] == 'unspecified_in_excerpt'
               and f['evidence_text'] == MOODYS_YEAR_END for f in fields)
    assert not concentration_fields(MOODYS_YEAR_END, agency_code='001196')


@pytest.mark.parametrize('old,new', [
    ('se evidencia', 'no se evidencia'), ('se evidencia', 'se podría evidenciar'),
    ('al cierre de 2025', 'a junio de 2025'), ('al término de 2024', 'anteriormente'),
    ('2024.', '2025.'), ('2024.', '2026.'), ('22.30', '1,2.3'),
])
def test_annual_comparison_rejects_unsupported_or_ambiguous_context(old, new):
    fields = concentration_fields(MOODYS_YEAR_END.replace(old, new), agency_code='000406')
    assert all(f['normalized_value'] != '22.3' for f in fields)


def test_undated_microrate_percent_remains_reviewable():
    fields, coverage = extract_body_evidence([
        'La concentración de los 20 principales depositantes se considera manejable (8.3% del total de depósitos).'
    ], agency_code='000410')
    assert not any(f['field_kind'] == 'deposit_concentration' for f in fields)
    assert next(c for c in coverage if c['topic'] == 'deposit_concentration')['coverage_status'] == 'needs_review'


def test_dated_concentration_excludes_other_comparisons():
    fields = concentration_fields(JCR, agency_code='001196')
    assert [(f['top_depositors'], f['normalized_value']) for f in fields] == [(10, '9.7'), (20, '13.2')]
    assert all(f['observation_period'] == '2025-12' and f['denominator_basis'] == 'total_deposits' for f in fields)
    assert not concentration_fields(JCR, agency_code='000409')
    assert not concentration_fields(JCR.replace('Al cierre de diciembre de 2025 ', ''), agency_code='001196')


def test_moodys_does_not_invent_denominator():
    fields = concentration_fields('concentración de depositantes (22,30% los 20 principales a diciembre de 2025)', agency_code='000406')
    assert len(fields) == 1 and fields[0]['normalized_value'] == '22.3'
    assert fields[0]['denominator_basis'] == 'unspecified_in_excerpt'


@pytest.mark.parametrize('value', ['101', '1,2.3', '..', '.', '1,'])
def test_invalid_percent_is_not_promoted(value):
    fields = concentration_fields(JCR.replace('9.7', value), agency_code='001196')
    assert all(f['top_depositors'] != 10 for f in fields)


def test_page_evidence_and_review_candidates_do_not_claim_events():
    fields, coverage = extract_body_evidence(['Portada', JCR + ' Estrategia con riesgo de liquidez. No hubo intervención.'], agency_code='001196')
    assert all(f['page_number'] == 2 for f in fields)
    candidates = [f for f in fields if f['field_kind'].startswith('qualitative_')]
    assert all(f['extraction_status'] == 'needs_review' and f['normalized_value'] == '' for f in candidates)
    assert any(f['field_kind'] == 'qualitative_event_mentions' and 'No hubo intervención' in f['evidence_text'] for f in candidates)
    statuses = {f['topic']: f['coverage_status'] for f in coverage}
    assert statuses['deposit_concentration'] == 'recognized'
    assert statuses['funding_cost'] == 'not_found_in_text'
    assert all(f['pages_checked'] == 2 for f in coverage)
    _, blank = extract_body_evidence(['', ''], agency_code='001196')
    assert all(f['coverage_status'] == 'needs_ocr' for f in blank)


def test_qualitative_candidates_are_bounded_per_topic():
    fields, _ = extract_body_evidence(['Estrategia ' + str(i) for i in range(20)], agency_code='000410')
    assert len(fields) == 2


def test_microrate_credit_rating_is_not_a_deposit_rating():
    fields = extract_cover_fields('CALIFICACIÓN PERSPECTIVA\nD+ Estable\nCALIFICACIÓN CREDITICIA', agency_code='000410')
    assert [(f['field_kind'], f['value_raw']) for f in fields] == [('credit_rating', 'D+'), ('credit_rating_outlook', 'Estable')]
    assert not extract_cover_fields('La calificación crediticia es D+ con perspectiva Estable', agency_code='000410')


OWNERSHIP_TABLE = ('Accionistas Acciones Participación (%) Municipalidad Provincial de Huancayo '
                   '81,421,830 92.33% Corporación Interamericana de Inversiones (BID Invest) '
                   '6,763,841 7.67% Total 88,185,671 100.00% Miembros del Directorio Cargo Condición')
CAPITALIZATION = ('→ Respaldo de su principal accionista, lo cual se ha visto reflejado en la '
                  'capitalización de utilidades (70% de capitalización en el 2025 '
                  'correspondiente a las utilidades del ejercicio 2024).')


def test_ownership_table_preserves_names_counts_and_published_percentages():
    from fuentes_financieras.providers.sbs._evidencia_informes import ownership_support_fields
    fields = ownership_support_fields(OWNERSHIP_TABLE, agency_code='001196')
    assert [(f['field_kind'], f['normalized_value']) for f in fields] == [
        ('shareholder_shares', '81421830'), ('shareholder_participation', '92.33'),
        ('shareholder_shares', '6763841'), ('shareholder_participation', '7.67')]
    assert fields[0]['field_label'] == 'Municipalidad Provincial de Huancayo'
    assert fields[2]['field_label'] == 'Corporación Interamericana de Inversiones (BID Invest)'
    assert all(f['temporal_role'] == 'unspecified' and 'observation_period' not in f for f in fields)
    assert not ownership_support_fields(OWNERSHIP_TABLE, agency_code='000406')
    assert not ownership_support_fields('Además, el accionista mayoritario es Corporación Coril S.A.C. (99.9%),', agency_code='001196')


@pytest.mark.parametrize('old,new', [
    ('88,185,671', '88,185,670'), ('92.33%', '90.33%'), ('7.67%', '9.67%'),
    ('100.00%', '99.99%'), ('81,421,830', '81.421.830'),
    ('Accionistas Acciones Participación (%)', 'Accionistas Participación (%)'),
    ('Miembros del Directorio Cargo Condición', ''),
    ('Corporación Interamericana de Inversiones (BID Invest)', 'Municipalidad Provincial de Huancayo'),
])
def test_ownership_table_rejects_broken_totals_or_missing_structure(old,new):
    from fuentes_financieras.providers.sbs._evidencia_informes import ownership_support_fields
    assert not ownership_support_fields(OWNERSHIP_TABLE.replace(old,new), agency_code='001196')


def test_repeated_ownership_tables_stay_unresolved():
    from fuentes_financieras.providers.sbs._evidencia_informes import ownership_support_fields
    assert not ownership_support_fields(OWNERSHIP_TABLE+' '+OWNERSHIP_TABLE, agency_code='001196')


def test_capitalization_uses_explicit_annual_date_and_profit_basis():
    fields, coverage = extract_body_evidence(['Portada', CAPITALIZATION], agency_code='000406')
    field = next(f for f in fields if f['field_kind'] == 'earnings_capitalization')
    assert field['normalized_value'] == '70.0' and field['observation_period'] == '2025'
    assert field['denominator_basis'] == 'earnings_year_2024' and field['temporal_role'] == 'annual_observation'
    assert field['page_number'] == 2 and field['evidence_text'] == CAPITALIZATION
    assert next(c for c in coverage if c['topic'] == 'ownership_support')['recognized_count'] == 1
    assert any(f['field_kind'] == 'qualitative_ownership_support' for f in fields)


@pytest.mark.parametrize('old,new', [
    ('se ha visto reflejado', 'se espera que se refleje'), ('Respaldo', 'Sin respaldo'),
    ('70%', '170%'), ('70%', '7,0.0%'), ('2024', '2025'), ('2025', '2101'),
    ('correspondiente a las utilidades del ejercicio 2024', 'de utilidades'),
])
def test_capitalization_rejects_forecasts_invalid_values_and_missing_year(old,new):
    from fuentes_financieras.providers.sbs._evidencia_informes import ownership_support_fields
    assert not ownership_support_fields(CAPITALIZATION.replace(old,new), agency_code='000406')
