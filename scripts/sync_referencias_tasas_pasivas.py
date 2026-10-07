"""Launch passive referencias_tasas with checked dependencies."""
from _bootstrap import prepare


if __name__ == '__main__':
    prepare()
    from fuentes_financieras.cli.referencias_tasas_pasivas import main
    raise SystemExit(main())
