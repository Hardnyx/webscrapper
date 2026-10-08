"""Explicitly sync selected independent liquidity datasets."""
import argparse
import os
from pathlib import Path

from fuentes_financieras import source
from fuentes_financieras.registry import get_provider
from .universo_depositos import write_report

DATASETS = {'liquidez': 'pe.sbs.liquidez', 'cobertura': 'pe.sbs.cobertura_liquidez', 'financiacion': 'pe.sbs.financiacion_neta_estable'}


def run(argv=None):
    parser = argparse.ArgumentParser(description='Liquidez, cobertura y financiación neta estable SBS.')
    parser.add_argument('--datasets', nargs='+', choices=list(DATASETS), default=['liquidez'])
    parser.add_argument('--desde', required=True, help='YYYY-MM')
    parser.add_argument('--hasta', help='YYYY-MM; por defecto --desde')
    parser.add_argument('--tipos', nargs='+', choices=list('BFCR'), default=list('BFCR'))
    parser.add_argument('--data-root', type=Path)
    parser.add_argument('--output-dir', type=Path, default=Path('outputs/liquidez'))
    parser.add_argument('--load-only', action='store_true')
    parser.add_argument('--force', action='store_true')
    args = parser.parse_args(argv)
    if args.data_root:
        os.environ['FINANCIAL_SOURCES_DATA_ROOT'] = str(args.data_root.resolve())
    get_provider.cache_clear()
    sheets = {}
    flags = 0
    for name in dict.fromkeys(args.datasets):
        dataset = DATASETS[name]
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
        report['unit'] = report.unit.replace({'percent': 'Porcentaje', 'thousands_PEN': 'Miles de soles', 'thousands_USD': 'Miles de dólares',
            'unspecified_by_source': 'Unidad no especificada por la fuente'})
        for code, label in {'source_period_differs_from_index': 'Período declarado distinto del mes del índice',
            'source_value_missing': 'Valor ausente en la fuente',
            'published_ratio_mismatch': 'Ratio publicado distinto del cociente de importes'}.items():
            report['data_quality_flags'] = report.data_quality_flags.str.replace(code, label, regex=False)
        report = report.rename(columns={'period': 'Mes del índice SBS'})
        sheets[name] = report
    filenames = {'liquidez': 'liquidez.xlsx', 'cobertura': 'cobertura_liquidez.xlsx',
                 'financiacion': 'financiacion_neta_estable.xlsx'}
    filename = filenames[next(iter(sheets))] if len(sheets) == 1 else 'liquidez_completa.xlsx'
    path = args.output_dir.resolve() / filename
    write_report(path, sheets)
    print(f'Filas: {sum(len(d) for d in sheets.values())}; filas con avisos de fuente: {flags}; reporte: {path}')
    if flags:
        print('Revisar avisos: el trimestre del RCL puede diferir del mes del índice. Se preservan ausencias y discrepancias publicadas.')
    return 0


def main(argv=None):
    try:
        return run(argv)
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        print(f'Error: {type(exc).__name__}: {exc}')
        return 1
