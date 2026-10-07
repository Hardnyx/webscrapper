"""Launch solvency with checked dependencies."""
from _bootstrap import prepare

if __name__ == '__main__':
    prepare()
    from fuentes_financieras.cli.solvencia import main
    raise SystemExit(main())
