"""Verify that moved entry points preserve public behavior and resolve local paths."""
import importlib.util
from pathlib import Path
from types import SimpleNamespace

import pandas as pd

from fuentes_financieras import sbs_tipo_cambio
from fuentes_financieras.cli import clasificaciones_riesgo
from fuentes_financieras.providers.sbs import tipo_cambio_contable
from fuentes_financieras.models import SyncResult

ROOT = Path(__file__).resolve().parents[1]


def test_accounting_compatibility_and_parser():
    assert sbs_tipo_cambio.get_accounting_exchange_rate is tipo_cambio_contable.get_accounting_exchange_rate
    html = b'<table><tr><td>30/09/2026</td><td>3.7500</td></tr></table>'
    result = sbs_tipo_cambio.parse_accounting_exchange_rate_response(html)
    assert result['usd_pen_accounting'].tolist() == [3.75]
    assert result['date'].iloc[0] == pd.Timestamp('2026-09-30')


def test_ratings_reports_use_selected_output_directory(tmp_path, monkeypatch):
    period = {'period_code': '202601', 'label': 'Marzo 2026', 'period_date': '2026-03-31', 'selected': True}
    frame = pd.DataFrame([{
        'period_code': '202601', 'period': '2026-03', 'period_date': '2026-03-31',
        'entity_type_code': 'S', 'entity_type': 'Seguros', 'entity_name': 'Entidad',
        'rating_agency': 'Agencia', 'rating': 'A', 'trend': None,
        'rating_kind': 'institutional_summary', 'trend_basis': 'published_change_vs_previous_classification',
        'data_quality_flags': '',
    }])

    class CachedRatings:
        contract_version = '2'
        storage = SimpleNamespace(manifest=SimpleNamespace(last_schema_contract_version='2', last_schema_hash=None))

        def plan_sync(self, **kwargs):
            return [SimpleNamespace(period_key=period['period_code'])]

        def sync(self, **kwargs):
            return SyncResult('pe.sbs.clasificaciones_riesgo', requested=1, skipped_existing=1)

        def load(self, **kwargs):
            return frame.copy()

    monkeypatch.setattr(clasificaciones_riesgo, 'source', lambda _: CachedRatings())
    monkeypatch.chdir(tmp_path)
    output = tmp_path / 'reports'
    assert clasificaciones_riesgo.main(['--data-root', str(tmp_path / 'data'), '--output-dir', str(output)]) == 0
    assert (output / 'clasificaciones_informes.xlsx').is_file()
    assert not (output / 'resumen_historico_clasificaciones.csv').exists()
    assert not (tmp_path / 'resultado_historico_clasificaciones.json').exists()


def test_site_capture_defaults_resolve_from_config_location(tmp_path, monkeypatch):
    path = ROOT / 'tools/site_capture/site_dump.py'
    spec = importlib.util.spec_from_file_location('repository_site_capture', path)
    module = importlib.util.module_from_spec(spec)
    import sys
    monkeypatch.setitem(sys.modules, spec.name, module)
    spec.loader.exec_module(module)
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(sys, 'argv', ['site_dump.py'])
    args = module.parse_args()
    assert Path(args.config) == path.with_name('site_dump_config.json')
    config = module.load_config_file(args.config)
    assert Path(config.out_dir) == ROOT / 'outputs/site_capture'
