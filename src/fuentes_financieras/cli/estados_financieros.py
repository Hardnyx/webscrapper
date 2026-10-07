"""Sync monthly statistical balance and income statements."""
import argparse
import os
from pathlib import Path

from fuentes_financieras import source
from fuentes_financieras.registry import get_provider
from .universo_depositos import write_report


def run(argv=None):
    parser = argparse.ArgumentParser(description='Estados financieros estadísticos SBS mensuales.')
    parser.add_argument('--desde', required=True, help='YYYY-MM')
    parser.add_argument('--hasta', help='YYYY-MM; por defecto, --desde')
    parser.add_argument('--tipos', nargs='+', choices=list('BFCR'), default=list('BFCR'))
    parser.add_argument('--data-root', type=Path)
    parser.add_argument('--output-dir', type=Path, default=Path('outputs/estados_financieros'))
    parser.add_argument('--load-only', action='store_true')
    parser.add_argument('--force', action='store_true')
    args = parser.parse_args(argv)
    if args.data_root:
        os.environ['FINANCIAL_SOURCES_DATA_ROOT'] = str(args.data_root.resolve())
    get_provider.cache_clear()
    provider = source('pe.sbs.estados_financieros')
    requests = list(provider.plan_sync(desde=args.desde, hasta=args.hasta, tipos=args.tipos))
    if not args.load_only:
        result = provider.sync(desde=args.desde, hasta=args.hasta, tipos=args.tipos, force=args.force, keep_raw=True)
        print(result)
        if result.failed or result.unavailable:
            print('Rango incompleto: no se exporta el reporte.')
            return 1
    data = provider.load(desde=args.desde, hasta=args.hasta or args.desde, tipos=args.tipos)
    if data.empty or {(r.entity_type, r.period) for r in data.itertuples()} != {(r.params['tipo'], r.params['periodo']) for r in requests}:
        raise ValueError('La caché no contiene todos los grupos y meses solicitados.')
    path = args.output_dir.resolve() / 'estados_financieros.xlsx'
    write_report(path, {'Balance': data[data.statement == 'balance'], 'Resultados': data[data.statement == 'income']})
    print(f'Filas: {len(data)}; reporte: {path}')
    return 0


def main(argv=None):
    try:
        return run(argv)
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        print(f'Error: {type(exc).__name__}: {exc}')
        return 1
