"""Explicitly export independently cached funding datasets."""
import argparse
import os
from pathlib import Path

from fuentes_financieras import source
from fuentes_financieras.registry import get_provider
from .universo_depositos import write_report

DATASETS = {'personas': 'pe.sbs.depositos_persona', 'escalas': 'pe.sbs.depositos_escalas',
            'adeudos': 'pe.sbs.adeudos', 'plazos': 'pe.sbs.depositos_plazo'}
NOTICE_LABELS = {
    'source_auxiliary_entity_name': 'Nombre auxiliar de la fuente; equivalencia legal sin establecer',
    'source_value_missing': 'Valor ausente o marcador original',
    'source_entity_placeholder': 'Fila sin entidad identificable; nombre original numérico',
    'source_date_not_month_end': 'Fecha original distinta del cierre de mes',
    'published_components_mismatch': 'Total publicado distinto de sus componentes',
    'published_bands_sum_mismatch': 'Total publicado distinto de sus tramos',
    'published_shares_sum_mismatch': 'Participaciones publicadas no suman 100 %',
    'published_percentage_out_of_range': 'Porcentaje publicado fuera de rango',
}


def run(argv=None, *, datasets=None, report_name='fondeo', description='Cuadros mensuales de fondeo o castigos SBS.'):
    choices = DATASETS if datasets is None else datasets
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument('--datasets', nargs='+', choices=list(choices),
                        default=['personas', 'escalas', 'adeudos'] if datasets is None else list(choices))
    parser.add_argument('--desde', required=True, help='YYYY-MM')
    parser.add_argument('--hasta', help='YYYY-MM; por defecto --desde')
    parser.add_argument('--tipos', nargs='+', choices=list('BFCR'), default=list('BFCR'))
    parser.add_argument('--data-root', type=Path)
    parser.add_argument('--output-dir', type=Path, default=Path('outputs') / report_name)
    parser.add_argument('--load-only', action='store_true')
    parser.add_argument('--force', action='store_true')
    args = parser.parse_args(argv)
    if 'plazos' in args.datasets and set(args.tipos)-set('BF'):
        parser.error('Plazos solo admite B/F: indique --tipos B F. No se omiten cajas silenciosamente.')
    if args.data_root:
        os.environ['FINANCIAL_SOURCES_DATA_ROOT'] = str(args.data_root.resolve())
    get_provider.cache_clear()
    sheets = {}; flags = 0
    for name in dict.fromkeys(args.datasets):
        provider = source(choices[name])
        requests = list(provider.plan_sync(desde=args.desde, hasta=args.hasta, tipos=args.tipos))
        if not args.load_only:
            result = provider.sync(desde=args.desde, hasta=args.hasta, tipos=args.tipos, force=args.force, keep_raw=True)
            print(result)
            if result.failed or result.unavailable:
                for detail in result.details:print(detail)
                print('Rango incompleto: no se exporta el reporte.')
                return 1
        data = provider.load(desde=args.desde, hasta=args.hasta or args.desde, tipos=args.tipos)
        expected = {(r.params['tipo'], r.params['periodo']) for r in requests}
        if data.empty or {(r.entity_type, r.period) for r in data.itertuples()} != expected:
            raise ValueError('La caché no contiene todos los grupos y meses solicitados.')
        flags += int((data.data_quality_flags != '').sum())
        report = data.copy()
        report['unit'] = report.unit.replace({'percent': 'Porcentaje', 'thousands_PEN': 'Miles de soles',
                                             'thousands_USD': 'Miles de dólares', 'count': 'Número publicado', 'thousands_PEN_per_person': 'Miles de soles por persona',
                                             'thousands_PEN_per_employee': 'Miles de soles por empleado',
                                             'thousands_PEN_per_office': 'Miles de soles por oficina'})
        for code, label in NOTICE_LABELS.items():
            report['data_quality_flags'] = report.data_quality_flags.str.replace(code, label, regex=False)
        sheets[name] = report
    path = args.output_dir.resolve() / (report_name+'.xlsx')
    write_report(path, sheets)
    print(f'Filas: {sum(len(d) for d in sheets.values())}; filas con avisos: {flags}; reporte: {path}')
    return 0


def main(argv=None):
    try:
        return run(argv)
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        print(f'Error: {type(exc).__name__}: {exc}')
        return 1
