"""Bootstrap dependencies before running the native series command."""
from _bootstrap import prepare
prepare()
from fuentes_financieras.cli.series_macro import main
if __name__ == '__main__':
    raise SystemExit(main())
