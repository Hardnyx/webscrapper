"""Export a bounded native series using the shared provider contracts."""
import argparse
import os
from pathlib import Path

import pandas as pd

from fuentes_financieras import source
from fuentes_financieras.catalog import SERIES_DATASETS
from fuentes_financieras.registry import get_provider
from fuentes_financieras.ejecucion import PlannedJob, execute_jobs
from .universo_depositos import write_report


def run(argv=None):
    p = argparse.ArgumentParser(description='Descarga series públicas con escala y frecuencia explícitas.')
    p.add_argument('--dataset', required=True, choices=sorted(SERIES_DATASETS))
    p.add_argument('--codigo', required=True)
    p.add_argument('--frecuencia', required=True, choices=['daily','monthly','quarterly','annual'])
    p.add_argument('--desde', required=True);p.add_argument('--hasta', required=True)
    p.add_argument('--data-root', type=Path, required=True)
    p.add_argument('--output-dir', type=Path, default=Path('outputs/series_macro'))
    p.add_argument('--load-only', action='store_true');p.add_argument('--force', action='store_true')
    a = p.parse_args(argv)
    if a.load_only and a.force:
        p.error('--load-only y --force son excluyentes.')
    os.environ['FINANCIAL_SOURCES_DATA_ROOT'] = str(a.data_root.resolve());get_provider.cache_clear()
    provider = source(a.dataset)
    query = dict(codigo=a.codigo,frecuencia=a.frecuencia,desde=a.desde,hasta=a.hasta)
    job = PlannedJob(a.codigo,provider,query,{'force':a.force,'keep_raw':True},list(provider.plan_offline(**query)))
    summary,coverage,code = execute_jobs([job],cache_only=a.load_only)
    if code:
        for row in coverage: print(row['Período solicitado'],row['Estado de captura'],row['Error'])
        return code
    data = provider.load(**query)
    if data.empty:
        print('Sin observaciones dentro del rango seleccionado.');return 2
    labels = {'series_code':'Código serie','title':'Título fuente','period':'Período original','period_date':'Fecha dato',
        'frequency':'Frecuencia declarada','value':'Valor nativo','source_value':'Valor fuente','unit':'Unidad fuente',
        'source_unit_multiplier':'Exponente de escala fuente','observation_status':'Estado fuente','data_quality_flags':'Avisos',
        'source':'Fuente','source_url':'URL fuente','retrieved_at':'Consultado UTC'}
    path = a.output_dir.resolve()/(a.dataset.replace('.','_')+'.xlsx')
    write_report(path,{'serie':data.rename(columns=labels),'capturas':pd.DataFrame(coverage)})
    print(f'Filas: {len(data)}; reporte: {path}')
    return 0


def main(argv=None):
    try:return run(argv)
    except KeyboardInterrupt:return 130
    except Exception as exc:
        print(f'Error: {type(exc).__name__}: {exc}');return 1

if __name__=='__main__':raise SystemExit(main())
