"""Compatibility entry point; prefer fuentes_financieras.cli.valores_cuota."""
from .cli.valores_cuota import main

if __name__ == "__main__":
    raise SystemExit(main())
