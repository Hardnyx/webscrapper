"""Launch foreign exchange risk with checked dependencies."""
from _bootstrap import prepare
if __name__ == '__main__':
    prepare()
    from fuentes_financieras.cli.riesgo_cambiario import main
    raise SystemExit(main())
