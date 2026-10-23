"""Published Philadelphia SPF vintage series, without modelled forecasts."""
from datetime import date
import io
import re

import pandas as pd

from fuentes_financieras.exceptions import InvalidQueryError, SchemaChangedError
from ._series import SeriesProvider, numeric


class PhiladelphiaSpfProvider(SeriesProvider):
    source_name = 'Federal Reserve Bank of Philadelphia - SPF'
    endpoint = 'https://www.philadelphiafed.org/-/media/frbp/assets/surveys-and-data/survey-of-professional-forecasters/data-files/files'
    frequencies = {'quarterly'}
    files = {'CPI':'median_cpi_level.xlsx','DRGDP':'median_rgdp_growth.xlsx',
             'TBILL':'median_tbill_level.xlsx','UNEMP':'median_unemp_level.xlsx'}

    def validate_code(self, code):
        if not isinstance(code, str) or not re.fullmatch(r'(CPI|DRGDP|TBILL|UNEMP)([1-6]|[ABC])', code):
            raise InvalidQueryError('Seleccione una columna SPF individual, por ejemplo CPIA o TBILL2.')

    def observations(self, code, frequency, year):
        prefix = next(k for k in self.files if code.startswith(k))
        url = self.endpoint+'/'+self.files[prefix]
        content = self.get(url).content
        if not content.startswith(b'PK\x03\x04'):
            raise SchemaChangedError('SPF no devolvió un XLSX válido.')
        try:
            frame = pd.read_excel(io.BytesIO(content), sheet_name=0)
        except Exception:
            raise SchemaChangedError('Libro SPF ilegible.') from None
        frame.columns = [str(c).strip().upper() for c in frame.columns]
        if not {'YEAR','QUARTER',code}.issubset(frame.columns) or frame.columns.duplicated().any():
            raise SchemaChangedError('Columnas SPF cambiadas o código no publicado en este archivo.')
        rows = []
        for row in frame[['YEAR','QUARTER',code]].to_dict('records'):
            if pd.isna(row['YEAR']) and pd.isna(row['QUARTER']):
                continue
            try:
                y, q = int(row['YEAR']), int(row['QUARTER'])
                if y != row['YEAR'] or q != row['QUARTER'] or not 1 <= q <= 4:
                    raise ValueError
                observed = date(y, (q-1)*3+1, 1)
            except (ValueError, TypeError):
                raise SchemaChangedError('Año o trimestre SPF inválido.') from None
            raw = row[code]
            rows.append(dict(date=observed.isoformat(), period=f'{y}-Q{q}', value=float('nan') if pd.isna(raw) else numeric(raw),
                source_value='' if pd.isna(raw) else str(raw), status='survey_vintage_not_target_period'))
        return rows, f'SPF - columna publicada {code}, mediana', url
