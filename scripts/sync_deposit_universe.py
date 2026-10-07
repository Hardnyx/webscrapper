"""Launch the deposit universe command with checked dependencies."""
from _bootstrap import prepare


if __name__ == '__main__':
    prepare()
    from fuentes_financieras.cli.deposit_universe import main
    raise SystemExit(main())
