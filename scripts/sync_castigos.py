"""Launch monthly write-offs with checked dependencies."""
from _bootstrap import prepare

if __name__ == '__main__':
    prepare()
    from fuentes_financieras.cli.castigos import main
    raise SystemExit(main())
