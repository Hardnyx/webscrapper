"""Physical Parquet files must not inherit Hive partition columns from paths."""
import pandas as pd

from fuentes_financieras.storage import DatasetStorage, content_hash, read_parquet_file


def test_partition_directory_does_not_add_columns(tmp_path):
    storage=DatasetStorage(tmp_path,'test.series')
    original=pd.DataFrame({'value':pd.Series([1.5],dtype='float64')})
    storage.upsert_period('serie=abc/year=2026','abc:2026',original)
    assert storage.period_matches('serie=abc/year=2026','abc:2026',content_hash(original))
    assert storage.read_all().columns.tolist()==['value','_period_key']
    assert storage.canonical_has_period('serie=abc/year=2026','abc:2026')


def test_native_year_column_does_not_conflict_with_partition_type(tmp_path):
    storage=DatasetStorage(tmp_path,'test.documents')
    original=pd.DataFrame({'year':pd.Series([2026],dtype='int64'),'value':pd.Series([1],dtype='int64')})
    storage.upsert_period('year=2026','first',original)
    storage.upsert_period('year=2026','second',original.assign(value=2))
    loaded=storage.read_partitions(['year=2026'])
    assert loaded.year.dtype=='int64' and loaded.value.tolist()==[1,2]
    assert storage.period_matches('year=2026','first',content_hash(original))
    assert read_parquet_file(storage.partition_path('year=2026'),columns=['value']).value.tolist()==[1,2]
