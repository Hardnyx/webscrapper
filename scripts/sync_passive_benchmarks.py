"""Launch passive benchmarks with checked dependencies."""
from _bootstrap import prepare


if __name__ == '__main__':
    prepare()
    from fuentes_financieras.cli.passive_benchmarks import main
    raise SystemExit(main())
