"""Export explicit SBS news selections and their extraction coverage."""
import argparse
import os
from pathlib import Path
from fuentes_financieras import source
from fuentes_financieras.registry import get_provider
from .universo_depositos import write_report


def verified_selection(provider, requests):
    for request in requests:
        entry = provider.storage.manifest.get(request.period_key)
        if not entry or entry.get('status') != 'validated' or entry.get('contract_version') != provider.contract_version or not provider.storage.period_matches(request.partition_key, request.period_key, entry.get('content_hash')):
            raise ValueError('Captura ausente, incompatible o dañada en caché.')


def discovery_selection(data, *, desde=None, hasta=None, palabras=None):
    from datetime import date
    import pandas as pd
    start = date.fromisoformat(desde).isoformat() if desde else None
    end = date.fromisoformat(hasta).isoformat() if hasta else None
    if start and end and start > end:
        raise ValueError('Rango de fechas invertido.')
    if data.groupby('announcement_id')[['article_url','source_title','listed_date']].nunique(dropna=False).gt(1).any().any():
        raise ValueError('Referencias contradictorias entre páginas; vuelva a capturarlas.')
    out = data.copy()
    if start: out = out[out.listed_date.ge(start)]
    if end: out = out[out.listed_date.le(end)]
    if palabras:
        mask = pd.Series(False, index=out.index)
        for word in palabras:
            if not word.strip(): raise ValueError('Palabra de selección vacía.')
            mask |= out.source_title.str.casefold().str.contains(word.strip().casefold(), regex=False, na=False)
        out = out[mask]
    return out.drop_duplicates('announcement_id').reset_index(drop=True)


def discover(args):
    import pandas as pd
    index = source('pe.sbs.indice_noticias')
    requests = list(index.plan_sync(paginas=args.paginas))
    if not args.load_only:
        result = index.sync(paginas=args.paginas, force=args.force, keep_raw=True)
        print(result)
        if result.failed or result.unavailable:
            raise ValueError('Índice incompleto; no se exporta una selección parcial.')
    verified_selection(index, requests)
    data = index.load(paginas=args.paginas)
    if data.empty or set(data.index_page) != {int(r.period_key) for r in requests}:
        raise ValueError('Faltan páginas solicitadas del índice.')
    selected = discovery_selection(data, desde=args.desde, hasta=args.hasta, palabras=args.palabras)
    rows = []
    for page, group in data.groupby('index_page', sort=True):
        rows.append(dict(index_page=page, listed_pages_total=int(group.listed_pages_total.iloc[0]),
                         news_count=len(group), selected_count=int(group.announcement_id.isin(selected.announcement_id).sum()),
                         index_status='Página capturada; no certifica todo el histórico ni todos los eventos',
                         source_url=group.source_url.iloc[0], retrieved_at=group.retrieved_at.iloc[0]))
    sheets = {'indice': data, 'seleccion': selected, 'cobertura_indice': pd.DataFrame(rows)}
    return selected.article_url.tolist(), sheets


def run(argv=None):
    parser = argparse.ArgumentParser(description='Anuncios regulatorios explícitos de noticias oficiales SBS.')
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument('--urls', nargs='+')
    mode.add_argument('--descubrir', action='store_true')
    parser.add_argument('--paginas', nargs='+', type=int)
    parser.add_argument('--desde', help='Fecha mínima del índice, YYYY-MM-DD.')
    parser.add_argument('--hasta', help='Fecha máxima del índice, YYYY-MM-DD.')
    parser.add_argument('--palabras', nargs='+', help='Subcadenas del título; se selecciona cualquier coincidencia.')
    parser.add_argument('--solo-indice', action='store_true')
    parser.add_argument('--data-root', type=Path)
    parser.add_argument('--output-dir', type=Path, default=Path('outputs/anuncios_regulatorios'))
    parser.add_argument('--load-only', action='store_true')
    parser.add_argument('--force', action='store_true')
    args = parser.parse_args(argv)
    if not args.descubrir and any((args.paginas, args.desde, args.hasta, args.palabras, args.solo_indice)):
        parser.error('Los filtros de índice requieren --descubrir.')
    args.paginas = args.paginas or [1]
    if args.load_only and args.force:
        parser.error('--load-only no se combina con --force.')
    if args.data_root:
        os.environ['FINANCIAL_SOURCES_DATA_ROOT'] = str(args.data_root.resolve())
    get_provider.cache_clear()
    path = args.output_dir.resolve() / 'anuncios_regulatorios.xlsx'
    urls, sheets = discover(args) if args.descubrir else (args.urls, {})
    if args.descubrir and (args.solo_indice or not urls):
        write_report(path, sheets)
        print(f'Noticias seleccionadas: {len(urls)}; alcance limitado a las páginas solicitadas; reporte: {path}')
        return 0
    provider = source('pe.sbs.anuncios_regulatorios')
    requests = list(provider.plan_sync(urls=urls))
    if not args.load_only:
        result = provider.sync(urls=urls, force=args.force, keep_raw=True)
        print(result)
        if result.failed or result.unavailable:
            print('Captura incompleta: no se exporta un reporte nuevo.')
            return 1
    verified_selection(provider, requests)
    data = provider.load(urls=urls)
    if data.empty or set(data.announcement_id) != {r.period_key for r in requests}:
        raise ValueError('Selección incompleta en caché.')
    report = data.copy()
    report['event_type'] = report.event_type.replace({
        'intervention_announced': 'Intervención anunciada',
        'dissolution_liquidation_announced': 'Disolución e inicio de liquidación anunciados'})
    report['extraction_status'] = report.extraction_status.replace({
        'extracted': 'Disposición reconocida en anuncio oficial',
        'needs_review': 'Sin disposición reconocida; requiere revisión'})
    sheets.update({'anuncios': report, 'eventos': report[data.extraction_status.eq('extracted')]})
    write_report(path, sheets)
    pending = int(data.extraction_status.ne('extracted').sum())
    print(f'Noticias: {len(data)}; disposiciones reconocidas: {len(data)-pending}; pendientes: {pending}; reporte: {path}')
    return 2 if pending else 0


def main(argv=None):
    try:
        return run(argv)
    except KeyboardInterrupt:
        return 130
    except Exception as exc:
        print(f'Error: {type(exc).__name__}: {exc}')
        return 1
