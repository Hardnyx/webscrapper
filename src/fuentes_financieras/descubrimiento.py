"""Explicit consumer adapters for references in verified publication inventories."""
import pandas as pd

SOURCES = {
    'pe.sbs.indice_noticias': ('announcement_id', 'article_url', 'source_title', 'listed_date', 'pe.sbs.anuncios_regulatorios'),
    'pe.moodys.indice_comunicados': ('action_id', 'article_url', 'source_title', 'listed_date', 'pe.moodys.referencias_comunicados'),
    'pe.moodys.referencias_comunicados': ('document_id', 'pdf_url', 'source_title', 'listed_date', 'pe.clasificadoras.retiros'),
    'pe.sbs.informes_riesgo': ('report_id', 'report_url', None, 'report_date', 'pe.sbs.documentos_riesgo'),
}


def reference_observations(planned, coverage):
    """Read only verified requested periods; never infer removals or legal events."""
    valid = {(r['Trabajo'], r['Dataset'], r['Período solicitado'])
             for r in coverage if r['Estado de captura'] == 'Captura validada'}
    found = {}
    for job in planned:
        dataset = job.provider.spec.dataset_id
        if dataset not in SOURCES:
            continue
        requests = [r for r in job.requests if (job.name, dataset, r.period_key) in valid]
        if not requests:
            continue
        data = job.provider.storage.read_partitions({r.partition_key for r in requests})
        data = data[data['_period_key'].isin({r.period_key for r in requests})]
        key_col, url_col, title_col, date_col, target = SOURCES[dataset]
        required = {key_col, url_col, date_col} | ({title_col} if title_col else set())
        if not required.issubset(data.columns):
            raise ValueError(f'Inventario incompatible para descubrimiento: {dataset}.')
        for row in data.to_dict('records'):
            def value(column):
                val = row.get(column)
                return '' if pd.isna(val) else str(val)
            identifier, url = value(key_col), value(url_col)
            # Missing SBS download links remain coverage issues, not publications.
            if not identifier and not url and dataset == 'pe.sbs.informes_riesgo':
                continue
            if not identifier or not url:
                raise ValueError(f'Referencia sin identidad o enlace: {dataset}.')
            payload = {'referencia': identifier, 'url': url, 'titulo': value(title_col) if title_col else '',
                       'fecha_publicada': value(date_col), 'proveedor_destino': target}
            key = (dataset, identifier)
            if key in found and found[key] != payload:
                raise ValueError(f'Referencia contradictoria entre capturas: {dataset}/{identifier}.')
            found[key] = payload
    return found
