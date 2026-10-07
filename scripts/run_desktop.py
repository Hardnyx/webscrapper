"""Launch one SBS desktop application with checked dependencies."""
import argparse
import runpy
from _bootstrap import prepare

APPS = {
    'average-exchange-rate': 'tipo_cambio_promedio.py',
    'weighted-exchange-rate': 'tipo_cambio_ponderado.py',
}

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Aplicaciones SBS de escritorio.')
    parser.add_argument('--app', required=True, choices=sorted(APPS))
    args = parser.parse_args()
    root = prepare('desktop')
    runpy.run_path(str(root / 'apps' / 'sbs' / APPS[args.app]), run_name='__main__')
