"""Compatibility entry point; prefer fuentes_financieras.cli.fund_values."""
from .cli.fund_values import main

if __name__ == "__main__":
    raise SystemExit(main())
