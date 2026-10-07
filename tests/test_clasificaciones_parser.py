from fuentes_financieras.providers.sbs._clasificaciones_parser import (
    entity_type_map,
    period_parts,
    to_long_form,
)


def test_period_parts():
    assert period_parts("202601") == (2026, 1, 3, "2026-03")
    assert period_parts("202602") == (2026, 2, 9, "2026-09")


def test_nested_rating_table_is_not_duplicated():
    html = """
    <html><body>
      <select name="ctl00$MainContent$DdlTiposEntidad">
        <option value="">Todos</option>
        <option value="S">Seguros</option>
      </select>
      <table>
        <tr>
          <th>Tipo de Entidad</th><th>Entidad</th>
          <th>JCR Latino America</th><th>PCR (Pacific Credit Rating)</th>
        </tr>
        <tr class="data">
          <td>Seguros</td><td>Entidad Perú</td>
          <td><table><tr><td><a>A</a></td><td>&nbsp;</td></tr></table></td>
          <td><table><tr><td><a>A</a></td><td>&nbsp;</td></tr></table></td>
        </tr>
      </table>
    </body></html>
    """

    mapping = entity_type_map(html)
    df = to_long_form(
        html,
        period_code="202601",
        type_code_by_label=mapping,
        source_url="https://example.test",
        retrieved_at="2026-10-06T00:00:00+00:00",
    )

    assert len(df) == 2
    assert set(df["rating_agency"]) == {
        "JCR Latino America",
        "PCR (Pacific Credit Rating)",
    }
    assert set(df["entity_type_code"]) == {"S"}
