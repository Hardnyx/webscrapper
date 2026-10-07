import pandas as pd
import pytest

from fuentes_financieras import list_datasets
from fuentes_financieras.catalog import CATALOG
from fuentes_financieras.exceptions import SchemaChangedError
from fuentes_financieras.models import FetchResult, PeriodRequest
from fuentes_financieras.provider import DatasetProvider
from fuentes_financieras.registry import get_provider
from fuentes_financieras.runtime import resolve_data_root


class LocalProvider(DatasetProvider):
    def __init__(self, spec):
        super().__init__(spec)
        self.frame = pd.DataFrame({'value': pd.Series([1], dtype='Int64')})
        self.calls = 0

    def single_request(self, **query):
        return PeriodRequest('202601', 'year=2026', {})

    def plan_sync(self, **query):
        return [self.single_request()]

    def _fetch_period(self, request):
        self.calls += 1
        return FetchResult(self.spec.dataset_id, self.frame.copy())


def test_provider_imports_and_standalone_storage(tmp_path, monkeypatch):
    monkeypatch.setenv('FINANCIAL_SOURCES_DATA_ROOT', str(tmp_path))
    get_provider.cache_clear()
    for dataset in list_datasets():
        provider = get_provider(dataset.dataset_id)
        assert provider.storage.root == tmp_path / CATALOG[dataset.dataset_id].storage_path
    monkeypatch.delenv('FINANCIAL_SOURCES_DATA_ROOT')
    assert resolve_data_root().name == 'sources'
    get_provider.cache_clear()


def test_cache_reuse_and_corrupted_partition_repair(tmp_path, monkeypatch):
    monkeypatch.setenv('FINANCIAL_SOURCES_DATA_ROOT', str(tmp_path))
    provider = LocalProvider(CATALOG['pe.sbs.tasas_pasivas'])
    assert provider.sync().downloaded == 1
    assert provider.sync().skipped_existing == 1
    assert provider.calls == 1
    path = provider.storage.partition_path('year=2026')
    saved = pd.read_parquet(path)
    saved['value'] = 99
    saved.to_parquet(path, index=False)
    assert provider.sync().downloaded == 1
    assert provider.load()['value'].tolist() == [1]


def test_contract_migration_preserves_same_contract_schema_guard(tmp_path, monkeypatch):
    monkeypatch.setenv('FINANCIAL_SOURCES_DATA_ROOT', str(tmp_path))
    provider = LocalProvider(CATALOG['pe.sbs.tasas_pasivas'])
    provider.sync()
    # Legacy manifests without a contract marker must still enforce the old schema.
    provider.storage.manifest.data.pop('last_schema_contract_version')
    provider.storage.manifest.save()
    provider.frame = pd.DataFrame({'value': pd.Series(['A'], dtype='string')})
    with pytest.raises(SchemaChangedError):
        provider.sync(force=True)
    provider.contract_version = '2'
    assert provider.sync().downloaded == 1
    assert provider.storage.manifest.last_schema_contract_version == '2'
    assert str(provider.load()['value'].dtype).startswith('string')
    assert provider.sync().skipped_existing == 1
    provider.frame = pd.DataFrame({'different': [1]})
    with pytest.raises(SchemaChangedError):
        provider.sync(force=True)
