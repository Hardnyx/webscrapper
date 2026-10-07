"""Compatibility entry point; prefer fuentes_financieras.cli.tasas_pasivas."""
from .cli.tasas_pasivas import main

if __name__ == "__main__":
    raise SystemExit(main())
