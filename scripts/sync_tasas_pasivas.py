"""Launch the tasas_pasivas command with checked dependencies."""
from _bootstrap import prepare


def main():
    prepare()
    from fuentes_financieras.cli.tasas_pasivas import main as run
    return run()


if __name__ == '__main__':
    raise SystemExit(main())
