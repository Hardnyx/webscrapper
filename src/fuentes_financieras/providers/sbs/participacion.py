"""Published size, rank and market shares without recomputing entity rankings."""
import math
import pandas as pd
from fuentes_financieras.exceptions import SchemaChangedError
from ._excel_mensual import MonthlyExcelProvider, clean
from .calidad_cartera import frame_for, validate_date, closing, number, source_notes, guard_percentage_formats

CODES = dict(zip('BFCR', ('B-2332', 'B-3243', 'C-1205', 'C-2205')))
PRODUCTS = {'Créditos Directos': 'direct_credit', 'Depósitos Totales': 'total_deposits', 'Patrimonio': 'equity'}


def parse_workbook(content, *, entity_type, period, source_url, retrieved_at):
    sheet, frame = frame_for(content, 'Ranking de Créditos, Depósitos y Patrimonio')
    cols = (2, 3, 4, 5) if entity_type in 'BF' else (1, 2, 3, 4)
    name_col, amount_col, share_col, cumulative_col = cols
    if frame.shape[1] != cumulative_col+1:raise SchemaChangedError('Dimensiones del ranking cambiadas.')
    starts = [(i, clean(v).rstrip(' *')) for i,v in frame.iloc[:,0].items() if clean(v).rstrip(' *') in PRODUCTS]
    if [label for _,label in starts] != list(PRODUCTS):raise SchemaChangedError('Bloques de ranking ausentes o duplicados.')
    validate_date(frame, period, starts[0][0])
    if not any(clean(v).replace(' ','').lower()=='(enmilesdesoles)' for v in frame.iloc[:starts[0][0],0]):
        raise SchemaChangedError('Unidad monetaria del ranking ausente.')
    footer = next((i for i in range(starts[-1][0]+1,len(frame)) if clean(frame.iloc[i,0]).lower().startswith(('fuente:', 'nota:'))), None)
    if footer is None:raise SchemaChangedError('Cierre de fuente ausente.')
    notes=source_notes(frame,footer)
    if entity_type=='B' and not any(n['text']=='No incluye sucursales en el exterior' for n in notes):
        raise SchemaChangedError('Cobertura territorial del ranking bancario ausente o cambiada.')
    rows=[]
    for pos,(start,label) in enumerate(starts):
        end=starts[pos+1][0] if pos+1<len(starts) else footer
        heads=[i for i in range(start+1,end) if clean(frame.iloc[i,0 if entity_type in 'BF' else 1])=='Empresas']
        if len(heads)!=1:raise SchemaChangedError('Encabezado de ranking ausente o ambiguo.')
        head=heads[0]
        if [clean(frame.iloc[head,j]) for j in (amount_col,share_col,cumulative_col)]!=['Monto','Participación','Porcentaje'] or clean(frame.iloc[head+1,cumulative_col])!='Acumulado' or clean(frame.iloc[head+1,share_col]).replace(' ','')!='(%)':
            raise SchemaChangedError('Columnas de ranking cambiadas.')
        names=set(); block=[]; cumulative=0.; complete=True
        for row in range(head+2,end):
            if not frame.iloc[row].notna().any():continue
            name=clean(frame.iloc[row,name_col]);rank=number(frame.iloc[row,0])
            if not isinstance(frame.iloc[row,name_col],str) or not name or name in names or not math.isfinite(rank) or rank<1 or not rank.is_integer():
                raise SchemaChangedError('Entidad o posición de ranking no reconocida.')
            if entity_type in 'BF' and clean(frame.iloc[row,1]):raise SchemaChangedError('Datos en columna auxiliar de ranking.')
            names.add(name); values=[number(frame.iloc[row,j]) for j in (amount_col,share_col,cumulative_col)]; flags=[]
            if rank!=len(names):flags.append('published_rank_sequence_review')
            if any(math.isfinite(v) and not 0<=v<=100.000001 for v in values[1:]):flags.append('published_percentage_out_of_range')
            if math.isfinite(values[1]):cumulative+=values[1]
            else:complete=False
            if complete and math.isfinite(values[2]) and abs(cumulative-values[2])>.0001:flags.append('published_cumulative_mismatch')
            for j,(metric,value,col,unit) in enumerate(zip(('amount','market_share','cumulative_market_share'),values,(amount_col,share_col,cumulative_col),('thousands_PEN','percent','percent'))):
                raw=frame.iloc[row,col]
                block.append(dict(period=period,period_date=closing(period).isoformat(),frequency='monthly',entity_type=entity_type,
                    entity_name=name,entity_scope='entity',ranking_product=PRODUCTS[label],published_rank=rank,metric=metric,value=value,
                    unit=unit,unit_multiplier=1000. if unit=='thousands_PEN' else float('nan'),currency='TOTAL',
                    comparison_scope='domestic_banking_excluding_foreign_branches' if entity_type=='B' else 'published_entity_group',
                    measurement_basis='published_month_end_ranking',source_metric=clean(frame.iloc[head,col]),
                    source_value_token=clean(raw) if isinstance(raw,str) else '',
                    data_quality_flags=';'.join(flags+(['source_value_missing'] if math.isnan(value) else [])),
                    source_sheet=str(sheet),source_row=row+1,source_column=col+1,source='SBS',source_url=source_url,retrieved_at=retrieved_at))
        if not block:raise SchemaChangedError('Bloque de ranking vacío.')
        shares=[r['value'] for r in block if r['metric']=='market_share']
        if all(math.isfinite(v) for v in shares) and abs(sum(shares)-100)>.0001:
            for r in block:r['data_quality_flags']=';'.join(filter(None,(r['data_quality_flags'],'published_shares_sum_mismatch')))
        rows.extend(block)
    guard_percentage_formats(content,rows)
    notes.append({'scope':'Shares and ranks refer to the published group, not all financial institutions together. Municipal rankings include CMCP when listed.',
                  'rank_contract':'Preserve published rank and cumulative share; do not rerank, fill missing shares or infer unique legal identities.'})
    data=pd.DataFrame(rows)
    for col in data:data[col]=data[col].astype('float64' if col in ('value','unit_multiplier','published_rank') else 'int64' if col in ('source_row','source_column') else 'string')
    return data,notes


class MarketParticipationProvider(MonthlyExcelProvider):
    codes=CODES
    parser_version='2026-10-08.1'
    parse_workbook=staticmethod(parse_workbook)
