"""Published foreign exchange exposures and explicitly lagged capital ratios."""
from datetime import date, timedelta
import re
import pandas as pd
from fuentes_financieras.exceptions import SchemaChangedError
from ._excel_mensual import MonthlyExcelProvider, clean
from .calidad_cartera import frame_for, validate_date, source_notes, normalized, guard_percentage_formats
from .fondeo import header, footer, finish, record, mark_sum

POSITION_CODES = dict(zip('BFCR', ('B-2368', 'B-3266', 'C-1260', 'C-2370')))
CAPITAL_CODES = {'C':'C-1301','R':'C-2301'}
POSITION_LABELS = (
    'Posición de Cambio de Balance en M. E. (a)',
    'Posición Neta en Derivados de M. E. * (b)',
    'Delta de las Posiciones Netas en Opciones sobre M. E. (c)',
    'Posición Global en M.E. (a)+(b)+(c)',
)
METRICS = ('balance_fx_position','net_fx_derivatives_position','net_fx_options_delta','global_fx_position')
GROUP_NAMES = dict(zip('BFCR', ('Empresa Bancaria','Empresa Financiera','Caja Municipal','Caja Rural de Ahorro y Crédito')))
RATIO_LABEL = 'Posición Global en M.E. / Patrimonio Efectivo'
CAPITAL_NOTE = 'La información del patrimonio efectivo corresponde a la del mes anterior.'


def parse_positions(content, *, entity_type, period, source_url, retrieved_at):
    sheet,frame=frame_for(content,'Posición Global en Moneda Extranjera')
    if not any(clean(v)=='Posición Global en Moneda Extranjera por '+GROUP_NAMES[entity_type] for v in frame.iloc[:4,0]):
        raise SchemaChangedError('Tipo de entidad de posición cambiaria incompatible.')
    head=header(frame);validate_date(frame,period,head)
    if frame.shape[1]!=5 or [clean(v) for v in frame.iloc[head,1:]]!=list(POSITION_LABELS):
        raise SchemaChangedError('Componentes de posición cambiaria cambiados.')
    if not any(clean(v).replace(' ','').lower()=='(enmilesdesoles)' for v in frame.iloc[:head,0]):
        raise SchemaChangedError('Unidad monetaria cambiada.')
    stop,notes=footer(frame,head+1);rows=[];names=set()
    for row in range(head+1,stop):
        if not frame.iloc[row].notna().any():continue
        name=clean(frame.iloc[row,0])
        if not isinstance(frame.iloc[row,0],str) or not name or name in names:
            raise SchemaChangedError('Entidad ausente o duplicada en posición cambiaria.')
        names.add(name);start=len(rows)
        for col,metric in enumerate(METRICS,1):
            r=record(entity_type=entity_type,period=period,sheet=sheet,row=row,col=col,name=name,
                metric=metric,label=POSITION_LABELS[col-1],raw=frame.iloc[row,col],source_url=source_url,
                retrieved_at=retrieved_at,currency='ME',measurement_basis='published_month_end_exposure',
                denominator_period='',denominator_date='')
            rows.append(r)
        mark_sum(rows,start+3,[start,start+1,start+2],'published_components_mismatch')
    notes.append({'unit_contract':'All FX components are published in thousands of PEN, not USD; preserve signed values and source markers.',
                  'derivatives_contract':'Preserve derivatives and options delta separately; missing components are not inferred as zero.'})
    return finish(rows,notes)


def parse_capital_ratio(content, *, entity_type, period, source_url, retrieved_at):
    sheet,frame=frame_for(content,'Indicadores Financieros por')
    titles=[(i,j) for i in range(min(4,len(frame))) for j,v in enumerate(frame.iloc[i]) if clean(v).startswith('Indicadores Financieros por')]
    if entity_type not in CAPITAL_CODES or any(clean(frame.iloc[i,j]) != 'Indicadores Financieros por '+GROUP_NAMES[entity_type] for i,j in titles):
        raise SchemaChangedError('Tipo de entidad del ratio cambiario incompatible.')
    label_cols=sorted({j for _,j in titles})
    starts=[i for i,v in frame.iloc[:,0].items() if clean(v)=='POSICIÓN EN MONEDA EXTRANJERA']
    if len(starts)!=1:raise SchemaChangedError('Sección de posición cambiaria ausente o ambigua.')
    start=starts[0]
    stop=next((i for i in range(start+1,len(frame)) if any(clean(v).startswith('Nota:') for v in frame.iloc[i])),None)
    if stop is None:raise SchemaChangedError('Notas de posición cambiaria ausentes.')
    notes=source_notes(frame,stop)
    if not any(re.sub(r'^\*+\s*','',n['text'])==CAPITAL_NOTE for n in notes):
        raise SchemaChangedError('Fecha relativa del patrimonio efectivo ausente o cambiada.')
    heads=[i for i in range(min(10,len(frame))) if not clean(frame.iloc[i,0]) and isinstance(frame.iloc[i,1],str) and clean(frame.iloc[i,1])]
    if len(heads)!=1:raise SchemaChangedError('Encabezado de entidades ambiguo.')
    head=heads[0];previous=date.fromisoformat(period+'-01')-timedelta(days=1);rows=[];names=set()
    for pos,label_col in enumerate(label_cols):
        validate_date(frame,period,head,label_col)
        if clean(frame.iloc[start,label_col])!='POSICIÓN EN MONEDA EXTRANJERA':raise SchemaChangedError('Bloques cambiarios desalineados.')
        pairs=[(i,normalized(frame.iloc[i,label_col])) for i in range(start+1,stop) if clean(frame.iloc[i,label_col])]
        if len(pairs)!=1 or pairs[0][1]!=RATIO_LABEL:raise SchemaChangedError('Ratio cambiario ausente, duplicado o cambiado.')
        row=pairs[0][0]
        if not re.search(r'\(\s*%\s*\)', clean(frame.iloc[row,label_col])):
            raise SchemaChangedError('Unidad porcentual del ratio cambiario ausente.')
        end_col=label_cols[pos+1] if pos+1<len(label_cols) else frame.shape[1]
        if any(frame.iloc[i,label_col+1:end_col].notna().any() for i in range(start+1,stop) if i != row):
            raise SchemaChangedError('Valores cambiarios fuera del indicador reconocido.')
        for col in range(label_col+1,end_col):
            name=clean(frame.iloc[head,col])
            if not name:
                if not pd.isna(frame.iloc[row,col]):raise SchemaChangedError('Ratio cambiario sin entidad.')
                continue
            if name in names or not isinstance(frame.iloc[head,col],str):raise SchemaChangedError('Entidad cambiaria duplicada o no reconocida.')
            names.add(name)
            rows.append(record(entity_type=entity_type,period=period,sheet=sheet,row=row,col=col,name=name,
                metric='global_fx_position_effective_capital_ratio',label=clean(frame.iloc[row,label_col]),
                raw=frame.iloc[row,col],source_url=source_url,retrieved_at=retrieved_at,unit='percent',currency='ME',
                measurement_basis='current_exposure_previous_month_effective_capital',
                denominator_period=previous.strftime('%Y-%m'),denominator_date=previous.isoformat()))
    guard_percentage_formats(content,rows)
    notes.append({'denominator_contract':'The ratio uses previous-month effective capital; no same-month capital join or implied exposure amount.'})
    return finish(rows,notes)


class ForeignExchangePositionProvider(MonthlyExcelProvider):
    codes=POSITION_CODES
    parser_version='2026-10-09.1'
    parse_workbook=staticmethod(parse_positions)


class ForeignExchangeCapitalRatioProvider(MonthlyExcelProvider):
    codes=CAPITAL_CODES
    parser_version='2026-10-09.1'
    parse_workbook=staticmethod(parse_capital_ratio)

    def plan_sync(self,*,tipos=('C','R'),**query):
        return super().plan_sync(tipos=tipos,**query)
