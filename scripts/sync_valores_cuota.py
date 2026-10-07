"""Launch the valores_cuota command with checked dependencies."""
from _bootstrap import prepare


def main():
    prepare()
    from fuentes_financieras.cli.valores_cuota import main as run
    return run()


if __name__ == '__main__':
    raise SystemExit(main())
