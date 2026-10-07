"""Sync general references and export published product averages from local rates."""
import argparse
import os
from pathlib import Path

from fuentes_financieras import source
from fuentes_financieras.referencias_tasas import product_benchmarks
from fuentes_financieras.registry import get_provider
from .universo_depositos import write_report


def run(argv=None):
    parser = argparse.ArgumentParser(description='Referencias SBS de tasas pasivas generales y por producto.')
    parser.add_argument('--desde')
    parser.add_argument('--hasta')
    parser.add_argument('--data-root', type=Path)
    parser.add_argument('--output-dir', type=Path, default=Path('outputs/referencias_tasas_pasivas'))
    parser.add_argument('--load-only', action='store_true')
    parser.add_argument('--force', action='store_true')
    args = parser.parse_args(argv)
    if not args.load_only and not args.desde:
        parser.error('--desde es obligatorio para sincronizar.')
    if args.data_root:
        os.environ['FINANCIAL_SOURCES_DATA_ROOT'] = str(args.data_root.resolve())
    get_provider.cache_clear()
    provider = source('pe.sbs.tasas_pasivas_mercado')
    if not args.load_only:
        result = provider.sync(desde=args.desde, hasta=args.hasta, force=args.force, keep_raw=True)
        print(result)
        if result.failed or result.unavailable:
            print('Rango incompleto: no se exporta el reporte como si estuviera completo.')
            return 1
    general = provider.load(desde=args.desde, hasta=args.hasta)
    if general.empty:
        raise ValueError('No hay referencias generales para el rango seleccionado.')
    # Product periods remain explicit: do not force monthly references into
    # the daily window selected for general references.
    products = product_benchmarks(source('pe.sbs.tasas_pasivas').load())
    path = args.output_dir.resolve() / 'referencias_tasas_pasivas.xlsx'
    write_report(path, {'Referencias generales': general, 'Promedios por producto': products})
    print(f'Referencias generales: {len(general)}; promedios locales por producto: {len(products)}')
    if products.empty:
        print('Sin tasas locales: promedios por producto no evaluados.')
    print(f'Reporte: {path}')
    return 0


def main(argv=None):
    try:
        return run(argv)
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        print(f'Error: {type(exc).__name__}: {exc}')
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
