"""Export explicit rating withdrawals from verified agency documents."""
import argparse
import os
from pathlib import Path

from fuentes_financieras import source
from fuentes_financieras.registry import get_provider
from .anuncios_regulatorios import verified_selection
from .universo_depositos import write_report


def discover(args):
    from datetime import date
    import pandas as pd
    start = date.fromisoformat(args.desde).isoformat() if args.desde else None
    end = date.fromisoformat(args.hasta).isoformat() if args.hasta else None
    if start and end and start > end: raise ValueError('Rango de fechas invertido.')
    index = source('pe.moodys.indice_comunicados')
    if not args.load_only:
        result = index.sync(force=args.force, keep_raw=True)
        if result.failed or result.unavailable: raise ValueError('Índice no validado.')
    verified_selection(index, list(index.plan_sync()))
    data = index.load()
    if data.empty: raise ValueError('Índice vacío.')
    selected = data.copy()
    if start: selected = selected[selected.listed_date.ge(start)]
    if end: selected = selected[selected.listed_date.le(end)]
    mask = pd.Series(False, index=selected.index)
    for word in args.palabras:
        if not word.strip(): raise ValueError('Palabra de selección vacía.')
        mask |= selected.source_title.str.casefold().str.contains(word.strip().casefold(), regex=False, na=False)
    selected = selected[mask].reset_index(drop=True)
    coverage = pd.DataFrame([dict(source='Moody’s Local Perú', source_url=data.source_url.iloc[0],
        retrieved_at=data.retrieved_at.iloc[0], coverage_status='Registros entregados en HTML; no certifica archivo completo',
        recognized_count=len(data), candidate_count=len(selected))])
    sheets = {'indice': data, 'seleccion': selected, 'cobertura_indice': coverage}
    if args.solo_indice or selected.empty: return [], sheets
    if len(selected) > args.max_comunicados: raise ValueError('Selección supera --max-comunicados; acote fechas o palabras.')
    references = source('pe.moodys.referencias_comunicados')
    urls = selected.article_url.tolist()
    if not args.load_only:
        result = references.sync(urls=urls, force=args.force, keep_raw=True)
        if result.failed or result.unavailable: raise ValueError('Referencias incompletas; no se descargan PDF.')
    verified_selection(references, list(references.plan_sync(urls=urls)))
    refs = references.load(urls=urls)
    if set(refs.action_id) != set(selected.action_id): raise ValueError('Referencias incompletas.')
    joined = selected.merge(refs[['action_id','source_title','listed_date']], on='action_id', suffixes=('_indice','_pagina'), validate='one_to_one')
    if not (joined.source_title_indice.eq(joined.source_title_pagina) & joined.listed_date_indice.eq(joined.listed_date_pagina)).all():
        raise ValueError('Título o fecha contradictorios entre índice y acción; recapture ambas fuentes.')
    sheets['referencias'] = refs
    return refs.pdf_url.drop_duplicates().tolist(), sheets


def run(argv=None):
    parser = argparse.ArgumentParser(description='Retiros de clasificaciones en PDF oficiales explícitos.')
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument('--urls', nargs='+')
    mode.add_argument('--descubrir', action='store_true')
    parser.add_argument('--desde')
    parser.add_argument('--hasta')
    parser.add_argument('--palabras', nargs='+')
    parser.add_argument('--solo-indice', action='store_true')
    parser.add_argument('--max-comunicados', type=int, default=20)
    parser.add_argument('--data-root', type=Path)
    parser.add_argument('--output-dir', type=Path, default=Path('outputs/retiros_clasificaciones'))
    parser.add_argument('--load-only', action='store_true')
    parser.add_argument('--force', action='store_true', help='Reprocesar PDF local verificado.')
    parser.add_argument('--redownload', action='store_true', help='Volver a descargar los PDF oficiales.')
    args = parser.parse_args(argv)
    if not args.descubrir and any((args.desde, args.hasta, args.palabras, args.solo_indice)):
        parser.error('Los filtros requieren --descubrir.')
    if args.max_comunicados < 1: parser.error('--max-comunicados debe ser positivo.')
    args.palabras = args.palabras or ['retira']
    if args.load_only and (args.force or args.redownload):
        parser.error('--load-only no se combina con --force ni --redownload.')
    if args.data_root:
        os.environ['FINANCIAL_SOURCES_DATA_ROOT'] = str(args.data_root.resolve())
    get_provider.cache_clear()
    urls, sheets = discover(args) if args.descubrir else (args.urls, {})
    if args.descubrir and (args.solo_indice or not urls):
        path = args.output_dir.resolve() / 'retiros_clasificaciones.xlsx'
        write_report(path, sheets)
        print(f'Índice y selección exportados: {path}')
        return 0
    provider = source('pe.clasificadoras.retiros')
    requests = list(provider.plan_sync(urls=urls, redownload=args.redownload))
    if not args.load_only:
        result = provider.sync(urls=urls, redownload=args.redownload, force=args.force)
        print(result)
        if result.failed or result.unavailable:
            print('Captura incompleta: no se exporta un reporte nuevo.')
            return 1
    verified_selection(provider, requests)
    for request in requests:
        entry = provider.storage.manifest.get(request.period_key)
        if not provider.pdf_matches(request.period_key, entry.get('metadata', {}).get('pdf_sha256')):
            raise ValueError('PDF ausente o dañado; captura no verificada.')
    data = provider.load(urls=urls)
    if data.empty or set(data.document_id) != {r.period_key for r in requests}:
        raise ValueError('Selección incompleta en caché.')
    report = data.copy()
    report['event_type'] = report.event_type.replace({'rating_withdrawal_announced': 'Retiro explícito de las clasificaciones indicadas'})
    report['extraction_status'] = report.extraction_status.replace({
        'extracted': 'Retiro reconocido en comunicado oficial', 'needs_review': 'Formato sin retiro reconocido; requiere revisión',
        'needs_ocr': 'Página inicial sin texto; requiere OCR'})
    report = report.rename(columns={'evidence_locator': 'Ubicación de evidencia en PDF'})
    path = args.output_dir.resolve() / 'retiros_clasificaciones.xlsx'
    sheets.update({'comunicados': report, 'retiros': report[data.extraction_status.eq('extracted')]})
    write_report(path, sheets)
    pending = int(data.extraction_status.ne('extracted').sum())
    print(f'Comunicados: {len(data)}; retiros reconocidos: {len(data)-pending}; pendientes: {pending}; reporte: {path}')
    return 2 if pending else 0


def main(argv=None):
    try:
        return run(argv)
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        print(f'Error: {type(exc).__name__}: {exc}')
        return 1
