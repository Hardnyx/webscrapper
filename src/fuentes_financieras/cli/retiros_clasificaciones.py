"""Export explicit rating withdrawals from verified agency documents."""
import argparse
import os
from pathlib import Path

from fuentes_financieras import source
from fuentes_financieras.registry import get_provider
from .anuncios_regulatorios import verified_selection
from .universo_depositos import write_report


def run(argv=None):
    parser = argparse.ArgumentParser(description='Retiros de clasificaciones en PDF oficiales explícitos.')
    parser.add_argument('--urls', nargs='+', required=True)
    parser.add_argument('--data-root', type=Path)
    parser.add_argument('--output-dir', type=Path, default=Path('outputs/retiros_clasificaciones'))
    parser.add_argument('--load-only', action='store_true')
    parser.add_argument('--force', action='store_true', help='Reprocesar PDF local verificado.')
    parser.add_argument('--redownload', action='store_true', help='Volver a descargar los PDF oficiales.')
    args = parser.parse_args(argv)
    if args.load_only and (args.force or args.redownload):
        parser.error('--load-only no se combina con --force ni --redownload.')
    if args.data_root:
        os.environ['FINANCIAL_SOURCES_DATA_ROOT'] = str(args.data_root.resolve())
    get_provider.cache_clear()
    provider = source('pe.clasificadoras.retiros')
    requests = list(provider.plan_sync(urls=args.urls, redownload=args.redownload))
    if not args.load_only:
        result = provider.sync(urls=args.urls, redownload=args.redownload, force=args.force)
        print(result)
        if result.failed or result.unavailable:
            print('Captura incompleta: no se exporta un reporte nuevo.')
            return 1
    verified_selection(provider, requests)
    for request in requests:
        entry = provider.storage.manifest.get(request.period_key)
        if not provider.pdf_matches(request.period_key, entry.get('metadata', {}).get('pdf_sha256')):
            raise ValueError('PDF ausente o dañado; captura no verificada.')
    data = provider.load(urls=args.urls)
    if data.empty or set(data.document_id) != {r.period_key for r in requests}:
        raise ValueError('Selección incompleta en caché.')
    report = data.copy()
    report['event_type'] = report.event_type.replace({'rating_withdrawal_announced': 'Retiro explícito de las clasificaciones indicadas'})
    report['extraction_status'] = report.extraction_status.replace({
        'extracted': 'Retiro reconocido en comunicado oficial', 'needs_review': 'Formato sin retiro reconocido; requiere revisión',
        'needs_ocr': 'Página inicial sin texto; requiere OCR'})
    report = report.rename(columns={'evidence_locator': 'Ubicación de evidencia en PDF'})
    path = args.output_dir.resolve() / 'retiros_clasificaciones.xlsx'
    write_report(path, {'comunicados': report, 'retiros': report[data.extraction_status.eq('extracted')]})
    pending = int(data.extraction_status.ne('extracted').sum())
    print(f'Comunicados: {len(data)}; retiros reconocidos: {len(data)-pending}; pendientes: {pending}; reporte: {path}')
    return 2 if pending else 0


def main(argv=None):
    try:
        return run(argv)
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        print(f'Error: {type(exc).__name__}: {exc}')
        return 1
