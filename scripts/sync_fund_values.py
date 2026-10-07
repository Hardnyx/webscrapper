"""Launch the fund_values command with checked dependencies."""
from _bootstrap import prepare


def main():
    prepare()
    from fuentes_financieras.cli.fund_values import main as run
    return run()


if __name__ == '__main__':
    raise SystemExit(main())
