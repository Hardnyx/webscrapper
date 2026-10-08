"""Launch size and participation with checked dependencies."""
from _bootstrap import prepare
if __name__=='__main__':
    prepare()
    from fuentes_financieras.cli.participacion import main
    raise SystemExit(main())
