"""Sync independent ratings and report inventories; export verified selections."""
import argparse
import os
from pathlib import Path

from fuentes_financieras import source
from fuentes_financieras.registry import get_provider
from .universo_depositos import write_report

DATASETS = {'clasificaciones': 'pe.sbs.clasificaciones_riesgo', 'informes': 'pe.sbs.informes_riesgo'}


def run(argv=None):
    parser = argparse.ArgumentParser(description='Clasificaciones institucionales e inventario de informes SBS.')
    parser.add_argument('--datasets', nargs='+', choices=list(DATASETS), default=['clasificaciones'])
    parser.add_argument('--periodos', nargs='+', help='Códigos SBS YYYY01 (marzo) / YYYY02 (septiembre).')
    parser.add_argument('--desde', help='YYYY-MM o YYYY-MM-DD')
    parser.add_argument('--hasta', help='YYYY-MM o YYYY-MM-DD')
    parser.add_argument('--data-root', type=Path)
    parser.add_argument('--output-dir', type=Path, default=Path('outputs/clasificaciones_riesgo'))
    parser.add_argument('--load-only', action='store_true')
    parser.add_argument('--force', action='store_true')
    parser.add_argument('--keep-raw', action='store_true')
    parser.add_argument('--no-second-sync', action='store_true')
    args = parser.parse_args(argv)
    if args.periodos and (args.desde or args.hasta):
        parser.error('Seleccione --periodos o un rango --desde/--hasta.')
    if args.load_only and args.force:
        parser.error('--load-only no se combina con --force.')
    if args.data_root:
        os.environ['FINANCIAL_SOURCES_DATA_ROOT'] = str(args.data_root.resolve())
    get_provider.cache_clear()
    query = {k: v for k, v in {'periodos': args.periodos, 'desde': args.desde, 'hasta': args.hasta}.items() if v is not None}
    sheets = {}
    for name in dict.fromkeys(args.datasets):
        provider = source(DATASETS[name])
        if not args.load_only:
            requests = list(provider.plan_sync(**query))
            expected = {r.period_key for r in requests}
            if not expected:
                raise ValueError('La selección no contiene períodos publicados.')
            result = provider.sync(**query, force=args.force, keep_raw=args.keep_raw)
            print(result)
            if result.failed or result.unavailable:
                print('Captura incompleta: no se exporta el reporte.')
                return 1
            if not args.no_second_sync:
                second = provider.sync(**query, keep_raw=args.keep_raw)
                if second.skipped_existing != len(expected) or second.downloaded or second.failed or second.unavailable:
                    raise ValueError('La segunda sincronización no reutilizó toda la captura.')
        else:
            entries = provider.storage.manifest.data['entries']
            valid_codes = [k for k, v in entries.items() if v.get('status') == 'validated'
                           and v.get('contract_version', '1') == provider.contract_version
                           and provider.storage.period_matches(v['partition_key'], k, v.get('content_hash'))]
            from fuentes_financieras.providers.sbs._clasificaciones_parser import period_date
            import pandas as pd
            catalog = pd.DataFrame({'period_code': valid_codes, 'period_date': [period_date(k) for k in valid_codes]})
            expected = set(provider.filter_loaded(catalog, **query).period_code.astype(str))
            if args.periodos and expected != set(args.periodos):
                raise ValueError('Faltan períodos solicitados o requieren migración en caché.')
        data = provider.load(**query)
        if not expected or data.empty or set(data.period_code.astype(str)) != expected:
            raise ValueError('La caché no contiene toda la selección validada bajo el contrato actual.')
        report = data.copy()
        if 'trend' in report:
            report['trend'] = report.trend.replace({'up': 'Subió', 'down': 'Bajó'})
            report['rating_kind'] = report.rating_kind.replace({'institutional_summary': 'Clasificación institucional del resumen'})
            report['trend_basis'] = report.trend_basis.replace({'published_change_vs_previous_classification': 'Cambio respecto de la clasificación anterior'})
        if 'document_status' in report:
            report['document_status'] = report.document_status.replace({'linked_not_downloaded': 'Enlace publicado; documento pendiente de descargar', 'link_missing': 'Enlace ausente'})
        report['data_quality_flags'] = report.data_quality_flags.str.replace('source_report_link_missing', 'Enlace al informe ausente', regex=False).str.replace('source_change_symbol_unrecognized', 'Símbolo de cambio sin interpretación', regex=False)
        sheets[name] = report
    path = args.output_dir.resolve() / 'clasificaciones_informes.xlsx'
    write_report(path, sheets)
    print(f'Filas: {sum(len(d) for d in sheets.values())}; reporte: {path}')
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
