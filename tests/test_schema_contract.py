from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
sys.path.insert(0, str(SRC))

from fuentes_financieras.providers.sbs._clasificaciones_parser import to_long_form
from fuentes_financieras.storage import schema_hash


def _html(trend: bool):
    img = '<img alt="sube" />' if trend else ""
    return f"""
    <html><body><table>
      <tr><th>Tipo de Entidad</th><th>Entidad</th><th>JCR</th></tr>
      <tr><td>Seguros</td><td>Entidad Perú</td><td><a>A</a>{img}</td></tr>
    </table></body></html>
    """


def test_schema_is_stable_when_nullable_trend_changes():
    kwargs = dict(
        period_code="202601",
        type_code_by_label={"Seguros": "S"},
        source_url="https://example.test",
        retrieved_at="2026-10-06T00:00:00+00:00",
    )
    without_trend = to_long_form(_html(False), **kwargs)
    with_trend = to_long_form(_html(True), **kwargs)

    assert str(without_trend["trend"].dtype).startswith("string")
    assert str(with_trend["trend"].dtype).startswith("string")
    assert str(without_trend["year"].dtype) == "Int64"
    assert str(without_trend["semester"].dtype) == "Int64"
    assert schema_hash(without_trend) == schema_hash(with_trend)
