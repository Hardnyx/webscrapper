"""Capture the SBS universe and report local dataset name correspondences."""
import argparse
import json
import os
from pathlib import Path

import pandas as pd
from openpyxl.worksheet.table import Table, TableStyleInfo

from fuentes_financieras import source
from fuentes_financieras.entities import EntityCatalog, correspondence_report, observation_changes
from fuentes_financieras.entity_aliases import reviewed_aliases
from fuentes_financieras.registry import get_provider
from fuentes_financieras.runtime import resolve_data_root

LABELS = {
    'previous_period_code': 'Código del corte anterior observado', 'previous_rating': 'Clasificación del corte anterior observado',
    'previous_report_id': 'Identificador del informe anterior', 'previous_report_url': 'URL informe anterior',
    'previous_source_url': 'URL fuente anterior', 'previous_retrieved_at': 'Fecha extracción anterior',
    'comparison_status': 'Estado de comparación', 'event_type': 'Tipo de cambio', 'event_basis': 'Evidencia del cambio',
    'previous_value': 'Valor del bloque anterior', 'current_value': 'Valor del bloque actual',
    'previous_page_number': 'Página de evidencia anterior', 'previous_evidence_text': 'Texto de evidencia anterior',
    'top_depositors': 'Número de principales depositantes', 'observation_period': 'Mes observado',
    'topic': 'Tema', 'recognized_count': 'Cifras reconocidas', 'candidate_count': 'Pasajes candidatos',
    'coverage_status': 'Estado de cobertura', 'pages_checked': 'Páginas revisadas', 'text_pages': 'Páginas con texto',
    'field_kind': 'Tipo de campo', 'field_label': 'Etiqueta original', 'value_raw': 'Valor original',
    'normalized_value': 'Valor normalizado', 'temporal_role': 'Vigencia indicada en el bloque',
    'extraction_status': 'Estado de extracción', 'evidence_text': 'Texto de evidencia',
    'page_number': 'Página del PDF', 'pdf_sha256': 'SHA-256 del PDF',

    'period_code': 'Código período SBS', 'year': 'Año', 'semester': 'Semestre',
    'rating_agency': 'Clasificadora', 'rating': 'Clasificación publicada',
    'summary_rating': 'Clasificación del resumen', 'rating_kind': 'Tipo de clasificación',
    'trend': 'Cambio publicado', 'trend_basis': 'Base del cambio',
    'source_change_icon_url': 'URL símbolo de cambio', 'source_change_title': 'Texto del cambio',
    'source_change_alt': 'Texto alternativo del símbolo',
    'report_url': 'URL informe', 'report_agency_code': 'Código clasificadora SBS',
    'report_period_code': 'Código período del informe', 'report_file_number': 'Número de archivo SBS',
    'report_version': 'Versión del informe', 'report_id': 'Identificador del informe',
    'document_status': 'Estado del documento', 'report_date': 'Fecha del informe',

    'denominator_period': 'Mes del denominador', 'denominator_date': 'Fecha del denominador',
    'ranking_product': 'Producto del ranking', 'published_rank': 'Posición publicada',
    'comparison_scope': 'Universo de comparación',
    'numerator_basis': 'Base del numerador', 'denominator_basis': 'Base del denominador',
    'annualization_definition': 'Definición de anualización',
    'source_auxiliary_entity_name': 'Nombre auxiliar publicado',
    'deposit_type': 'Tipo de depósito', 'term_bucket': 'Tramo de plazo',
    'funding_origin': 'Origen del acreedor', 'credit_type': 'Tipo de crédito',
    'band_kind': 'Tipo de tramo de montos', 'band_lower_PEN': 'Límite inferior (soles)',
    'band_upper_PEN': 'Límite superior (soles)', 'source_band_label': 'Tramo original',
    'boundary_convention': 'Inclusión de límites',
    'value': 'Valor publicado', 'unit_multiplier': 'Multiplicador a unidades monetarias',
    'source_value': 'Valor almacenado en Excel', 'source_number_format': 'Formato numérico original',
    'observation_start': 'Inicio del período observado', 'observation_end': 'Fin del período observado',
    'amount_basis': 'Base del importe',
    'source_column': 'Columna fuente', 'source_value_token': 'Marcador original de la fuente',
    'risk_category': 'Categoría de riesgo', 'credit_scope': 'Cobertura crediticia',
    'arrears_threshold_days': 'Umbral de atraso, mayor que (días)',
    'unit_evidence': 'Evidencia de unidad',
    'source_metric': 'Indicador original', 'data_quality_flags': 'Avisos de calidad de fuente',
    'source_sheet': 'Hoja fuente', 'source_row': 'Fila fuente',
    'source_auxiliary_date': 'Fecha auxiliar del archivo', 'source_entity_name': 'Encabezado original',
    'statement': 'Estado', 'section': 'Sección', 'account': 'Cuenta',
    'source_account': 'Cuenta original', 'account_code': 'Código de cuenta',
    'amount': 'Importe publicado', 'measurement_basis': 'Base de medición',
    'period': 'Fecha observación', 'period_date': 'Fecha dato',
    'entity_type_code': 'Tipo entidad', 'entity_type': 'Categoría',
    'entity_name': 'Nombre fuente', 'normalized_name': 'Nombre normalizado',
    'entity_id': 'Identificador interno', 'source_order': 'Orden fuente',
    'source': 'Fuente', 'source_url': 'URL fuente', 'retrieved_at': 'Fecha extracción',
    'dataset': 'Dataset', 'match_status': 'Correspondencia', 'evidence': 'Evidencia',
    'first_observed': 'Primera observación', 'last_observed': 'Última observación',
    'previous_period': 'Observación anterior', 'change': 'Cambio observado',
    'alias': 'Nombre equivalente', 'valid_from': 'Desde', 'valid_to': 'Hasta',
    'frequency': 'Frecuencia', 'metric': 'Indicador o producto', 'currency': 'Moneda',
    'rate': 'Tasa (%)', 'unit': 'Unidad', 'basis': 'Base',
    'observation_window': 'Ventana de observación', 'entity_scope': 'Grupo de entidades',
    'reference_kind': 'Tipo de referencia', 'methodology_url': 'Metodología',
    'table_kind': 'Tipo de cuadro', 'person_type': 'Tipo de persona',
}


def write_report(path, sheets):
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        for i, (name, frame) in enumerate(sheets.items(), start=1):
            frame = frame.drop(columns=['_period_key'], errors='ignore').rename(columns=LABELS)
            frame.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            ws.freeze_panes = 'A2'
            if not frame.empty:
                table = Table(displayName=f'Universe{i}', ref=ws.dimensions)
                table.tableStyleInfo = TableStyleInfo(name='TableStyleLight9', showRowStripes=True)
                ws.add_table(table)
            else:
                ws.auto_filter.ref = ws.dimensions


def run(argv=None):
    parser = argparse.ArgumentParser(description='Universo SBS y correspondencias con datos locales.')
    parser.add_argument('--data-root', type=Path)
    parser.add_argument('--output-dir', type=Path, default=Path('outputs/universo_depositos'))
    parser.add_argument('--load-only', action='store_true')
    parser.add_argument('--force', action='store_true')
    parser.add_argument('--html', type=Path, help='Importar una captura local en lugar de consultar SBS.')
    parser.add_argument('--observed-on', help='Fecha real de la captura local, YYYY-MM-DD.')
    parser.add_argument('--aliases', type=Path, help='JSON de equivalencias revisadas y fechadas.')
    args = parser.parse_args(argv)
    if bool(args.html) != bool(args.observed_on) or (args.load_only and args.html):
        parser.error('--html requiere --observed-on y no se combina con --load-only.')
    if args.data_root:
        os.environ['FINANCIAL_SOURCES_DATA_ROOT'] = str(args.data_root.resolve())
    get_provider.cache_clear()
    provider = source('pe.sbs.universo_depositos')
    if not args.load_only:
        result = (provider.import_capture(args.html, observed_on=args.observed_on, force=args.force)
                  if args.html else provider.sync(force=args.force, keep_raw=True))
        if result.failed or result.unavailable:
            for detail in result.details:
                print(detail)
            print('La captura falló. No se genera un reporte vigente a partir de datos antiguos.')
            return 1
    universe = provider.load()
    if universe.empty:
        raise ValueError('No hay capturas locales del universo.')
    catalog = EntityCatalog(resolve_data_root() / 'reference' / 'entity_catalog.json')
    aliases = json.loads(args.aliases.read_text(encoding='utf-8')) if args.aliases else []
    if not isinstance(aliases, list):
        raise ValueError('El archivo de alias debe ser una lista JSON.')
    catalog.observe(universe, aliases=aliases)
    frames = {key: source(dataset).load() for key, dataset in (
        ('rates', 'pe.sbs.tasas_pasivas'), ('ratings', 'pe.sbs.clasificaciones_riesgo'))}
    matches = correspondence_report(catalog, **frames, aliases=aliases)
    summary = pd.DataFrame([{
        'Dataset': key, 'Filas locales': len(frame),
        'Estado': 'Disponible' if not frame.empty else 'Sin datos locales; no evaluado',
    } for key, frame in frames.items()])
    current = universe[universe.period == universe.period.max()]
    report = args.output_dir.resolve() / 'universo_correspondencias.xlsx'
    write_report(report, {
        'Universo vigente observado': current, 'Historial capturas': universe,
        'Identidades': pd.DataFrame(catalog.records), 'Correspondencias': matches,
        'Equivalencias revisadas': pd.DataFrame([*reviewed_aliases(catalog), *aliases],
            columns=['dataset', 'entity_type_code', 'alias', 'entity_id', 'valid_from', 'valid_to', 'evidence']),
        'Cambios observados': observation_changes(universe, catalog, aliases), 'Cobertura local': summary,
    })
    catalog.save()
    print(f'Universo: {len(current)} entidades, observación {current.period.iloc[0]}')
    print(summary.to_string(index=False))
    print(f'Reporte: {report}')
    return 0


def main(argv=None):
    try:
        return run(argv)
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        print(f'Error: {type(exc).__name__}: {exc}')
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
