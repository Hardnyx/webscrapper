"""Launch the risk_ratings command with checked dependencies."""
from _bootstrap import prepare


def main():
    prepare()
    from fuentes_financieras.cli.risk_ratings import main as run
    return run()


if __name__ == '__main__':
    raise SystemExit(main())
