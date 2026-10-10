"""Launch configurable batch execution with checked dependencies."""
from _bootstrap import prepare

if __name__ == '__main__':
    prepare('pdf', 'browser')
    from fuentes_financieras.cli.ejecucion import main
    raise SystemExit(main())
