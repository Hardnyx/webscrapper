"""Launch PDF risk reports with checked dependencies."""
from _bootstrap import prepare
if __name__ == '__main__':
    prepare('pdf')
    from fuentes_financieras.cli.documentos_riesgo import main
    raise SystemExit(main())
