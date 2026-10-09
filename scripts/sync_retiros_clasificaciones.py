"""Launch explicit rating withdrawal extraction with checked PDF dependencies."""
from _bootstrap import prepare
if __name__ == '__main__':
    prepare('pdf')
    from fuentes_financieras.cli.retiros_clasificaciones import main
    raise SystemExit(main())
