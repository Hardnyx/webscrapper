"""Export independent FX exposures and lagged capital ratios."""
from .fondeo import run


def main(argv=None):
    try:
        return run(argv, datasets={'posicion': 'pe.sbs.posicion_cambiaria',
                   'capital': 'pe.sbs.posicion_cambiaria_capital'}, default_datasets=['posicion'],
                   report_name='riesgo_cambiario', description='Posición cambiaria mensual y ratio sobre capital SBS.')
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        print(f'Error: {type(exc).__name__}: {exc}')
        return 1
