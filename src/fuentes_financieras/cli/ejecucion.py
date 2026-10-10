"""Run a declarative source profile and export requested capture status."""
import argparse
import json
import os
from pathlib import Path
from uuid import uuid4

import pandas as pd

from fuentes_financieras.ejecucion import read_profile, plan_jobs, execute_jobs
from fuentes_financieras.registry import get_provider
from fuentes_financieras.storage import utc_now_iso
from .universo_depositos import write_report


def run(argv=None):
    parser = argparse.ArgumentParser(description='Ejecuta un perfil de fuentes y verifica las capturas solicitadas.')
    parser.add_argument('--perfil', type=Path, required=True)
    parser.add_argument('--data-root', type=Path, required=True)
    parser.add_argument('--output-dir', type=Path, default=Path('outputs/ejecuciones'))
    parser.add_argument('--max-solicitudes', type=int, default=1000)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument('--solo-plan', action='store_true')
    mode.add_argument('--solo-cache', action='store_true')
    args = parser.parse_args(argv)
    jobs = read_profile(args.perfil)
    os.environ['FINANCIAL_SOURCES_DATA_ROOT'] = str(args.data_root.resolve())
    get_provider.cache_clear()
    planned = plan_jobs(jobs, max_requests=args.max_solicitudes, offline=args.solo_plan or args.solo_cache)
    summary, coverage, code = execute_jobs(planned, cache_only=args.solo_cache, plan_only=args.solo_plan)
    # Each run has its own directory; previous reports are never overwritten.
    run_id = utc_now_iso().replace(':', '').replace('+', '_')+'_'+uuid4().hex[:12]
    folder = args.output_dir.resolve()/run_id
    write_report(folder/'ejecucion.xlsx', {'trabajos': pd.DataFrame(summary), 'capturas': pd.DataFrame(coverage)})
    payload = {'version': 1, 'run_id': run_id, 'modo': 'plan' if args.solo_plan else 'cache' if args.solo_cache else 'sync',
               'codigo_salida': code, 'trabajos': summary, 'capturas': coverage}
    temporary = folder/'ejecucion.json.tmp'
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')
    temporary.replace(folder/'ejecucion.json')
    for job in summary:
        print(f"{job['Trabajo']}: {job['Estado']} ({job['Capturas validadas']}/{job['Solicitudes']})")
    print(f'Reporte: {folder / "ejecucion.xlsx"}')
    return code


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
