"""Explicitly consume the report inventory and export PDF evidence."""
import argparse
import os
from pathlib import Path

from fuentes_financieras import source
from fuentes_financieras.registry import get_provider
from .universo_depositos import write_report


def run(argv=None):
    parser=argparse.ArgumentParser(description='Descarga íntegra y extracción conservadora de informes PDF SBS.')
    parser.add_argument('--periodos',nargs='+',required=True)
    parser.add_argument('--entidad')
    parser.add_argument('--clasificadora')
    parser.add_argument('--tipo-entidad')
    parser.add_argument('--data-root',type=Path)
    parser.add_argument('--output-dir',type=Path,default=Path('outputs/documentos_riesgo'))
    parser.add_argument('--load-only',action='store_true')
    parser.add_argument('--force',action='store_true',help='Volver a analizar los PDF íntegros guardados.')
    parser.add_argument('--redownload',action='store_true',help='Volver a descargar los documentos seleccionados.')
    args=parser.parse_args(argv)
    if args.load_only and (args.force or args.redownload):parser.error('--load-only no admite --force ni --redownload.')
    if args.data_root:os.environ['FINANCIAL_SOURCES_DATA_ROOT']=str(args.data_root.resolve())
    get_provider.cache_clear()
    inventory=source('pe.sbs.informes_riesgo')
    if not args.load_only:
        result=inventory.sync(periodos=args.periodos,keep_raw=True)
        print(result)
        if result.failed or result.unavailable:return 1
    for code in set(args.periodos):
        entry=inventory.storage.manifest.get(code)
        if not entry or entry.get('status')!='validated' or entry.get('contract_version')!=inventory.contract_version or not inventory.storage.period_matches(entry['partition_key'],code,entry.get('content_hash')):
            raise ValueError('Inventario incompleto, incompatible o dañado en caché.')
    refs=inventory.load(periodos=args.periodos,entidad=args.entidad,clasificadora=args.clasificadora,tipo_entidad=args.tipo_entidad)
    if refs.empty or refs.report_url.eq('').any():raise ValueError('Selección vacía o con enlaces ausentes.')
    documents=source('pe.sbs.documentos_riesgo')
    urls=refs.report_url.drop_duplicates().tolist()
    requests=list(documents.plan_sync(urls=urls))
    if not args.load_only:
        result=documents.sync(urls=urls,force=args.force or args.redownload,redownload=args.redownload)
        print(result)
        if result.failed or result.unavailable:
            print('Descarga incompleta: no se exporta el reporte.');return 1
    for request in requests:
        entry=documents.storage.manifest.get(request.period_key)
        if not entry or entry.get('status')!='validated' or entry.get('contract_version')!=documents.contract_version or not documents.storage.period_matches(request.partition_key,request.period_key,entry.get('content_hash')) or not documents.pdf_matches(request.period_key,entry.get('metadata',{}).get('pdf_sha256')):
            raise ValueError('Documento ausente o caché PDF/canónica dañada.')
    expected={r.period_key for r in requests}
    fields=documents.load(report_ids=list(expected))
    if fields.empty or set(fields.report_id)!=expected:raise ValueError('No se recuperaron todos los documentos seleccionados.')
    # Published associations are explicit consumer-level joins, not legal identity matches.
    associations=refs[['report_id','entity_type','entity_name','rating_agency']].drop_duplicates()
    report=fields.merge(associations,on='report_id',how='left')
    pending=int((fields.extraction_status!='extracted').sum())
    report['extraction_status']=report.extraction_status.replace({'extracted':'Extraído del formato reconocido',
        'needs_review':'Requiere revisión', 'needs_ocr':'Sin texto; requiere OCR', 'unsupported_cover':'Portada sin formato reconocido'})
    report['temporal_role']=report.temporal_role.replace({'current':'Actual según el bloque fuente','previous':'Anterior según el bloque fuente','unspecified':'Sin asignación temporal'})
    report['field_kind']=report.field_kind.replace({'financial_strength':'Fortaleza financiera','entity_rating':'Clasificación de entidad',
        'issuer_rating':'Clasificación de emisor','short_term_deposits':'Depósitos de corto plazo','medium_long_term_deposits':'Depósitos de mediano y largo plazo',
        'long_term_deposits':'Depósitos de largo plazo','outlook':'Perspectiva del informe','entity_rating_outlook':'Perspectiva de entidad',
        'issuer_rating_outlook':'Perspectiva de emisor','committee_date':'Fecha de comité','publication_date':'Fecha de publicación','document':'Documento pendiente de extracción'})
    path=args.output_dir.resolve()/'documentos_riesgo.xlsx'
    write_report(path,{'campos':report,'referencias':associations})
    print(f'Documentos: {len(expected)}; campos: {len(fields)}; pendientes de revisión/extracción: {pending}; reporte: {path}')
    return 0 if pending==0 else 2


def main(argv=None):
    try:return run(argv)
    except KeyboardInterrupt:return 130
    except Exception as exc:
        print(f'Error: {type(exc).__name__}: {exc}');return 1


if __name__=='__main__':raise SystemExit(main())
