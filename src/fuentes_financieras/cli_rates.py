"""Compatibility entry point; prefer fuentes_financieras.cli.passive_rates."""
from .cli.passive_rates import main

if __name__ == "__main__":
    raise SystemExit(main())
