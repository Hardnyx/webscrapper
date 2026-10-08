"""Export published size and participation rankings."""
from .fondeo import run


def main(argv=None):
    try:
        return run(argv,datasets={'participacion':'pe.sbs.participacion'},report_name='participacion',
                   description='Ranking mensual de créditos, depósitos y patrimonio SBS.')
    except KeyboardInterrupt:return 130
    except Exception as exc:
        print(f'Error: {type(exc).__name__}: {exc}')
        return 1
