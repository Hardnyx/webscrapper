from dataclasses import replace

import pytest

from fuentes_financieras.catalog import CATALOG
from fuentes_financieras.exceptions import SchemaChangedError
from fuentes_financieras.providers.sbs.tasas_pasivas import PassiveRatesProvider


@pytest.mark.parametrize('tipo', ['C', 'R'])
def test_explicit_absence_is_cached_and_can_be_rechecked(tmp_path, monkeypatch, tipo):
    monkeypatch.setenv('FINANCIAL_SOURCES_DATA_ROOT', str(tmp_path))
    provider = PassiveRatesProvider(CATALOG['pe.sbs.tasas_pasivas'])
    original_request = provider.single_request
    monkeypatch.setattr(provider, 'single_request',
        lambda **query: replace(original_request(**query), mutable=True))
    client = provider._client(tipo)
    monkeypatch.setattr(client, 'open', lambda: setattr(client, 'opened', True))
    monkeypatch.setattr(client, '_monthly_payload', lambda **kwargs: {})
    calls = []
    html = '<p>a Setiembre del 2026</p><p>No existe información para la fecha elegida</p>'

    def response(payload):
        calls.append(payload)
        return html, html

    monkeypatch.setattr(client, '_post', response)
    query = dict(tipos=[tipo], desde='2026-09-01', hasta='2026-09-01')
    first = provider.sync(**query)
    assert first.unavailable == 1 and first.failed == 0
    assert provider.sync(**query).skipped_existing == 1
    assert len(calls) == 1
    assert provider.load().empty
    # Mutable unavailable periods must not remain unavailable indefinitely.
    assert provider.sync(refresh_hours=0, **query).unavailable == 1
    assert len(calls) == 2


def test_missing_tables_without_absence_notice_still_fails(tmp_path, monkeypatch):
    monkeypatch.setenv('FINANCIAL_SOURCES_DATA_ROOT', str(tmp_path))
    provider = PassiveRatesProvider(CATALOG['pe.sbs.tasas_pasivas'])
    client = provider._client('C')
    client.opened = True
    monkeypatch.setattr(client, '_monthly_payload', lambda **kwargs: {})
    html = '<p>a Setiembre del 2026</p><p>Unexpected layout</p>'
    monkeypatch.setattr(client, '_post', lambda payload: (html, html))
    with pytest.raises(SchemaChangedError):
        client.fetch_monthly(year=2026, month=9)
