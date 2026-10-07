from datetime import date
from types import SimpleNamespace

import pandas as pd
import pytest

from fuentes_financieras.benchmarks import product_benchmarks
from fuentes_financieras.catalog import CATALOG
from fuentes_financieras.exceptions import PeriodUnavailableError, SchemaChangedError
from fuentes_financieras.providers.sbs.passive_market import PassiveMarketProvider, parse_market


def market_html(day='06/10/2026', flow_day=None):
    return '<input name="__VIEWSTATE" value="state"/>' + (
        f'<span id="ctl00_cphContent_lblFecha">Tasa efectiva al {day}</span>'
        f'<span id="ctl00_cphContent_lblFecha2">30 días útiles al {flow_day or day}</span>'
        '<span id="ctl00_cphContent_lblVAL_TIPMN_TASA">1.97</span>'
        '<span id="ctl00_cphContent_lblVAL_TIPMEX_TASA">1,17</span>'
        '<span id="ctl00_cphContent_lblVAL_FTIPMN">2.06</span>'
        '<span id="ctl00_cphContent_lblVAL_FTIPMEX">2.25</span>')


def test_general_references_keep_dates_scopes_and_bases():
    frame = parse_market(market_html(), target=date(2026, 10, 6), retrieved_at='now')
    assert frame.rate.tolist() == [1.97, 1.17, 2.06, 2.25]
    assert frame.basis.tolist() == ['stock', 'stock', 'flow', 'flow']
    assert frame.entity_scope.tolist() == ['B+F', 'B+F', 'B', 'B']
    with pytest.raises(PeriodUnavailableError):
        parse_market(market_html(flow_day='05/10/2026'), target=date(2026, 10, 6), retrieved_at='now')
    with pytest.raises(SchemaChangedError):
        parse_market(market_html().replace('>2.25<', '>-<'), target=date(2026, 10, 6), retrieved_at='now')


def test_historical_post_and_second_sync_uses_cache(tmp_path, monkeypatch):
    monkeypatch.setenv('FINANCIAL_SOURCES_DATA_ROOT', str(tmp_path))
    provider = PassiveMarketProvider(CATALOG['pe.sbs.tasas_pasivas_mercado'])
    calls = []

    def request(method, url, **kwargs):
        calls.append((method, kwargs))
        if method == 'GET':
            return SimpleNamespace(text=market_html())
        assert kwargs['data']['ctl00$cphContent$rdpDate$dateInput'] == '05/10/2026'
        return SimpleNamespace(text=market_html('05/10/2026'))

    monkeypatch.setattr(provider.transport, 'request', request)
    query = dict(desde='2026-10-05', hasta='2026-10-05')
    assert provider.sync(**query).downloaded == 1
    assert [call[0] for call in calls] == ['GET', 'POST']
    assert provider.sync(**query).skipped_existing == 1
    assert len(calls) == 2
    assert len(provider.load(metric='TIPMN')) == 1


def test_products_preserve_published_average_and_window():
    frame = pd.DataFrame([
        dict(entity_type=kind, frequency=freq, period='2026-08', period_date='2026-08-01',
             currency='MN', table_kind='persona', person_type='JURIDICA',
             metric='181-360 días', entity_name=name, rate=rate,
             source='SBS', source_url='https://www.sbs.gob.pe', retrieved_at='now')
        for kind, freq, name, rate in [('B', 'daily', 'Promedio', 4.1),
            ('B', 'daily', 'BANK', 9.0), ('C', 'monthly', 'Promedio', 5.2)]
    ])
    refs = product_benchmarks(frame)
    assert refs.rate.tolist() == [4.1, 5.2]
    assert refs.observation_window.tolist() == ['last_30_business_days', 'calendar_month']
    assert refs.person_type.tolist() == ['JURIDICA', 'JURIDICA']
    with pytest.raises(SchemaChangedError):
        product_benchmarks(pd.concat([frame, frame.iloc[:1]]))
    frame.loc[0, 'frequency'] = 'monthly'
    with pytest.raises(SchemaChangedError):
        product_benchmarks(frame)
