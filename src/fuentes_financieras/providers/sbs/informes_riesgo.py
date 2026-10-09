"""Independent inventory of report references published in the SBS summary."""
from ._clasificaciones_parser import to_long_form
from .clasificaciones_riesgo import RiskRatingsProvider


def report_inventory(html, **query):
    data = to_long_form(html, **query)
    data = data.drop(columns=['trend', 'rating_kind', 'trend_basis', 'source_change_icon_url',
                             'source_change_title', 'source_change_alt']).rename(columns={'rating': 'summary_rating'})
    data['document_status'] = data.report_url.map(lambda url: 'linked_not_downloaded' if url else 'link_missing').astype('string')
    # A summary period is not the document's publication or rating committee date.
    data['report_date'] = ''
    data['report_date'] = data.report_date.astype('string')
    return data


class RiskReportInventoryProvider(RiskRatingsProvider):
    parser_version = '2026-10-09.1'
    contract_version = '1'
    parse_summary = staticmethod(report_inventory)
