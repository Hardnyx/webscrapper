"""Export independently cached profitability and efficiency indicators."""
from .fondeo import run

DATASETS = {'rentabilidad': 'pe.sbs.rentabilidad', 'eficiencia': 'pe.sbs.eficiencia'}


def main(argv=None):
    try:
        return run(argv, datasets=DATASETS, report_name='rentabilidad_eficiencia',
                   description='Rentabilidad, eficiencia y gestión por entidad SBS.')
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        print(f'Error: {type(exc).__name__}: {exc}')
        return 1
