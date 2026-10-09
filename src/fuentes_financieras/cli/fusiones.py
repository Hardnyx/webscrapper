"""Export individually verified merger resolutions without inferring execution."""
import argparse
import os
from pathlib import Path

from fuentes_financieras import source
from fuentes_financieras.registry import get_provider
from .anuncios_regulatorios import verified_selection
from .universo_depositos import write_report


def run(argv=None):
    parser = argparse.ArgumentParser(description='Resoluciones SBS de fusión publicadas en El Peruano; URLs explícitas.')
    parser.add_argument('--urls', nargs='+', required=True)
    parser.add_argument('--data-root', type=Path)
    parser.add_argument('--output-dir', type=Path, default=Path('outputs/fusiones'))
    parser.add_argument('--load-only', action='store_true')
    parser.add_argument('--force', action='store_true')
    args = parser.parse_args(argv)
    if args.load_only and args.force:
        parser.error('--load-only no se combina con --force.')
    if args.data_root:
        os.environ['FINANCIAL_SOURCES_DATA_ROOT'] = str(args.data_root.resolve())
    get_provider.cache_clear()
    provider = source('pe.elperuano.fusiones')
    requests = list(provider.plan_sync(urls=args.urls))
    if not args.load_only:
        result = provider.sync(urls=args.urls, force=args.force, keep_raw=True)
        print(result)
        if result.failed or result.unavailable:
            print('Captura incompleta: no se exporta un reporte nuevo.')
            return 1
    verified_selection(provider, requests)
    data = provider.load(urls=args.urls)
    if data.empty or set(data.norm_id) != {r.period_key for r in requests}:
        raise ValueError('Selección incompleta en caché.')
    report = data.copy()
    translations = {
        'event_type': {'merger_authorized': 'Fusión autorizada; ejecución no acreditada',
                       'merger_date_clarification': 'Aclaración de fecha de vigencia de fusión'},
        'extraction_status': {'extracted': 'Disposición reconocida', 'needs_review': 'Requiere revisión'},
        'roles_status': {'unspecified': 'Roles sin extraer', 'operative_article': 'Roles explícitos en artículo'},
        'effective_basis': {'not_extracted': 'Fecha de vigencia sin extraer',
                            'date_in_clarification_article': 'Fecha estipulada en artículo aclaratorio; no prueba ejecución'},
    }
    for column, values in translations.items():
        report[column] = report[column].replace(values)
    # These headings are specific to gazette resolutions, not SBS news announcements.
    report = report.rename(columns={'resolution_number': 'Número de resolución SBS',
                                    'effective_date': 'Fecha de vigencia estipulada; no prueba ejecución'})
    path = args.output_dir.resolve() / 'fusiones.xlsx'
    write_report(path, {'resoluciones': report, 'disposiciones': report[data.extraction_status.eq('extracted')]})
    pending = int(data.extraction_status.ne('extracted').sum())
    print(f'Resoluciones: {len(data)}; reconocidas: {len(data)-pending}; pendientes: {pending}; reporte: {path}')
    return 2 if pending else 0


def main(argv=None):
    try:
        return run(argv)
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        print(f'Error: {type(exc).__name__}: {exc}')
        return 1
