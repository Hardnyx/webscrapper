"""Sync independent SBS solvency datasets and export their published values."""
import argparse
import os
from pathlib import Path

from fuentes_financieras import source
from fuentes_financieras.registry import get_provider
from .universo_depositos import write_report

DATASETS = {'Ratios y requerimientos': 'pe.sbs.solvencia', 'Patrimonio efectivo': 'pe.sbs.patrimonio_efectivo'}


def run(argv=None):
    parser = argparse.ArgumentParser(description='Solvencia y patrimonio efectivo SBS mensuales.')
    parser.add_argument('--desde', required=True, help='YYYY-MM')
    parser.add_argument('--hasta', help='YYYY-MM; por defecto --desde')
    parser.add_argument('--tipos', nargs='+', choices=list('BFCR'), default=list('BFCR'))
    parser.add_argument('--data-root', type=Path)
    parser.add_argument('--output-dir', type=Path, default=Path('outputs/solvencia'))
    parser.add_argument('--load-only', action='store_true')
    parser.add_argument('--force', action='store_true')
    args = parser.parse_args(argv)
    if args.data_root:
        os.environ['FINANCIAL_SOURCES_DATA_ROOT'] = str(args.data_root.resolve())
    get_provider.cache_clear()
    sheets = {}
    flags = 0
    for name, dataset in DATASETS.items():
        provider = source(dataset)
        requests = list(provider.plan_sync(desde=args.desde, hasta=args.hasta, tipos=args.tipos))
        if not args.load_only:
            result = provider.sync(desde=args.desde, hasta=args.hasta, tipos=args.tipos, force=args.force, keep_raw=True)
            print(result)
            if result.failed or result.unavailable:
                print('Rango incompleto: no se exporta el reporte.')
                return 1
        data = provider.load(desde=args.desde, hasta=args.hasta or args.desde, tipos=args.tipos)
        expected = {(r.params['tipo'], r.params['periodo']) for r in requests}
        if data.empty or {(r.entity_type, r.period) for r in data.itertuples()} != expected:
            raise ValueError('La caché no contiene todos los grupos y meses solicitados.')
        flags += int((data.data_quality_flags != '').sum())
        report = data.copy()
        report['unit'] = report.unit.replace({'percent': 'Porcentaje', 'thousands_PEN': 'Miles de soles',
            'unspecified_by_source': 'Unidad no especificada por la fuente'})
        report['data_quality_flags'] = report.data_quality_flags.str.replace(
            'published_components_sum_mismatch', 'Componentes publicados no coinciden con el total', regex=False).str.replace(
            'source_unit_unspecified', 'Unidad del total no especificada', regex=False)
        sheets[name] = report
    path = args.output_dir.resolve() / 'solvencia.xlsx'
    write_report(path, sheets)
    print(f'Filas: {sum(len(d) for d in sheets.values())}; filas con avisos de fuente: {flags}; reporte: {path}')
    if flags:
        print('Revisar avisos: se preservan inconsistencias publicadas y unidades no explícitas. No se corrigen ni convierten.')
    return 0


def main(argv=None):
    try:
        return run(argv)
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        print(f'Error: {type(exc).__name__}: {exc}')
        return 1
