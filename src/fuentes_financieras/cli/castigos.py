"""Export independently cached monthly credit write-off flows."""
from .fondeo import run


def main(argv=None):
    try:
        return run(argv, datasets={'castigos': 'pe.sbs.castigos'}, report_name='castigos')
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        print(f'Error: {type(exc).__name__}: {exc}')
        return 1
