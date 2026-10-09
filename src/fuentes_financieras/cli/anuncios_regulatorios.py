"""Export explicit SBS news selections and their extraction coverage."""
import argparse
import os
from pathlib import Path
from fuentes_financieras import source
from fuentes_financieras.registry import get_provider
from .universo_depositos import write_report


def run(argv=None):
    parser = argparse.ArgumentParser(description='Anuncios regulatorios explícitos de noticias oficiales SBS.')
    parser.add_argument('--urls', nargs='+', required=True)
    parser.add_argument('--data-root', type=Path)
    parser.add_argument('--output-dir', type=Path, default=Path('outputs/anuncios_regulatorios'))
    parser.add_argument('--load-only', action='store_true')
    parser.add_argument('--force', action='store_true')
    args = parser.parse_args(argv)
    if args.load_only and args.force:
        parser.error('--load-only no se combina con --force.')
    if args.data_root:
        os.environ['FINANCIAL_SOURCES_DATA_ROOT'] = str(args.data_root.resolve())
    get_provider.cache_clear()
    provider = source('pe.sbs.anuncios_regulatorios')
    requests = list(provider.plan_sync(urls=args.urls))
    if not args.load_only:
        result = provider.sync(urls=args.urls, force=args.force, keep_raw=True)
        print(result)
        if result.failed or result.unavailable:
            print('Captura incompleta: no se exporta un reporte nuevo.')
            return 1
    for request in requests:
        entry = provider.storage.manifest.get(request.period_key)
        if not entry or entry.get('status') != 'validated' or entry.get('contract_version') != provider.contract_version or not provider.storage.period_matches(request.partition_key, request.period_key, entry.get('content_hash')):
            raise ValueError('Noticia ausente, incompatible o dañada en caché.')
    data = provider.load(urls=args.urls)
    if data.empty or set(data.announcement_id) != {r.period_key for r in requests}:
        raise ValueError('Selección incompleta en caché.')
    report = data.copy()
    report['event_type'] = report.event_type.replace({
        'intervention_announced': 'Intervención anunciada',
        'dissolution_liquidation_announced': 'Disolución e inicio de liquidación anunciados'})
    report['extraction_status'] = report.extraction_status.replace({
        'extracted': 'Disposición reconocida en anuncio oficial',
        'needs_review': 'Sin disposición reconocida; requiere revisión'})
    path = args.output_dir.resolve() / 'anuncios_regulatorios.xlsx'
    write_report(path, {'anuncios': report, 'eventos': report[data.extraction_status.eq('extracted')]})
    pending = int(data.extraction_status.ne('extracted').sum())
    print(f'Noticias: {len(data)}; disposiciones reconocidas: {len(data)-pending}; pendientes: {pending}; reporte: {path}')
    return 2 if pending else 0


def main(argv=None):
    try:
        return run(argv)
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        print(f'Error: {type(exc).__name__}: {exc}')
        return 1
