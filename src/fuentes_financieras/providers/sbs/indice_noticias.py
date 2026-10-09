"""Independent paginated SBS press-note inventory; no article downloads."""
import re
from datetime import date
from hashlib import sha256
from urllib.parse import urljoin, urlsplit
import pandas as pd
from bs4 import BeautifulSoup
from fuentes_financieras.exceptions import InvalidQueryError, SchemaChangedError, SourceUnavailableError
from fuentes_financieras.models import FetchResult, PeriodRequest
from fuentes_financieras.provider import DatasetProvider
from fuentes_financieras.storage import utc_now_iso
from fuentes_financieras.transports.curl_chrome import CurlChromeTransport
from .anuncios_regulatorios import MONTHS, announcement_reference

URL = 'https://www.sbs.gob.pe/noticia'


def news_link(href):
    parsed = urlsplit(urljoin(URL, href))
    if parsed.netloc != 'www.sbs.gob.pe' or parsed.scheme not in ('http', 'https'):
        raise SchemaChangedError('Enlace de noticia fuera de la SBS.')
    try:
        return announcement_reference(parsed._replace(scheme='https').geturl())
    except InvalidQueryError as exc:
        raise SchemaChangedError('Enlace de noticia ambiguo o incompatible.') from exc


def parse_index(html, *, expected_page, source_url, retrieved_at):
    if not re.search(r'</html\s*>', html, re.I) or len(html.encode('utf-8')) > 2_000_000:
        raise SchemaChangedError('Índice HTML incompleto o mayor de 2 MB.')
    soup = BeautifulSoup(html, 'lxml')
    lists = soup.select('ul#SBS_PressNote')
    if len(lists) != 1:
        raise SchemaChangedError('Listado de noticias ausente o ambiguo.')
    modules, current, last = set(), set(), set()
    for anchor in soup.select('ul.pagination a[href]'):
        parsed = urlsplit(urljoin(URL, anchor['href']))
        match = re.match(r'/noticia/moduleid/(\d+)/pagina/(\d+)/controller/', parsed.path, re.I)
        if parsed.netloc != 'www.sbs.gob.pe' or parsed.scheme not in ('http', 'https') or not match:
            raise SchemaChangedError('Enlace de paginación cambiado.')
        modules.add(match[1])
        page = int(match[2])
        if 'active' in (anchor.parent.get('class') or []):
            current.add(page)
        if anchor.get_text(' ', strip=True).casefold() in ('ultima', 'última'):
            last.add(page)
    if len(modules) != 1 or current != {expected_page} or len(last) != 1 or not 1 <= expected_page <= next(iter(last)):
        raise SchemaChangedError('Paginación ausente, contradictoria o página distinta de la solicitada.')
    rows = []
    for item in lists[0].select('li.list-news__item'):
        anchors = item.select('h3.list-news__item__header__title a[href]')
        dates = item.select('header .date')
        if len(anchors) != 1 or len(dates) != 1:
            raise SchemaChangedError('Tarjeta de noticia sin enlace o fecha únicos.')
        identifier, link = news_link(anchors[0]['href'])
        title = re.sub(r'\s+', ' ', anchors[0].get('title', '')).strip()
        token = re.sub(r'\s+', ' ', dates[0].get_text(' ', strip=True)).strip()
        match = re.fullmatch(r'(\d{1,2}) ([a-z]+) (\d{4})', token, re.I)
        if not title or not match:
            raise SchemaChangedError('Título completo o fecha de tarjeta ausentes.')
        try:
            listed_date = date(int(match[3]), MONTHS[match[2].lower()], int(match[1])).isoformat()
        except (ValueError, KeyError):
            raise SchemaChangedError('Fecha de tarjeta inválida.') from None
        rows.append(dict(announcement_id=identifier, article_url=link, source_title=title,
                         listed_date=listed_date, index_page=expected_page, listed_pages_total=next(iter(last)),
                         source='SBS', source_url=source_url, retrieved_at=retrieved_at))
    data = pd.DataFrame(rows)
    if data.empty or data.announcement_id.duplicated().any():
        raise SchemaChangedError('Índice vacío o noticias duplicadas dentro de una página.')
    for col in data:
        data[col] = data[col].astype('int64' if col in ('index_page', 'listed_pages_total') else 'string')
    return data, dict(module_id=next(iter(modules)), index_page=expected_page,
                      listed_pages_total=next(iter(last)), html_sha256=sha256(html.encode('utf-8')).hexdigest())


class NewsIndexProvider(DatasetProvider):
    parser_version = '2026-10-09.1'
    contract_version = '1'

    def __init__(self, spec):
        super().__init__(spec)
        self.transport = CurlChromeTransport(state_dir=self.storage.state_root, timeout=180)
        self.initial = None

    def single_request(self, *, pagina=1, **_):
        if isinstance(pagina, bool) or not str(pagina).isdigit() or not 1 <= int(pagina) <= 1000:
            raise InvalidQueryError('Indique una página entera entre 1 y 1000.')
        page = int(pagina)
        return PeriodRequest(period_key=str(page), partition_key='indice', params={'pagina': page}, mutable=True)

    def plan_sync(self, *, paginas=None, **_):
        pages = [1] if paginas is None else [paginas] if isinstance(paginas, (int, str)) else paginas
        seen = set()
        for page in pages:
            request = self.single_request(pagina=page)
            if request.period_key not in seen:
                seen.add(request.period_key)
                yield request

    def _response(self, url):
        response = self.transport.request('GET', url)
        final = urlsplit(response.url)
        if final.scheme != 'https' or final.netloc != 'www.sbs.gob.pe' or final.path.casefold().rstrip('/') != urlsplit(url).path.casefold().rstrip('/') or final.query or final.fragment:
            raise SourceUnavailableError('Redirección fuera del índice solicitado.')
        return response

    def _fetch_period(self, request):
        if self.initial is None:
            response = self._response(URL)
            data, metadata = parse_index(response.text, expected_page=1, source_url=URL, retrieved_at=utc_now_iso())
            self.initial = (data, metadata, response.text)
        page = request.params['pagina']
        if page > self.initial[1]['listed_pages_total']:
            raise InvalidQueryError('Página mayor que la última publicada en el índice.')
        if page == 1:
            data, metadata, raw = self.initial
        else:
            url = f"{URL}/moduleId/{self.initial[1]['module_id']}/pagina/{page}/controller/Item/action/Index"
            response = self._response(url)
            data, metadata = parse_index(response.text, expected_page=page, source_url=url, retrieved_at=utc_now_iso())
            raw = response.text
        return FetchResult(self.spec.dataset_id, data.copy(), raw=raw, metadata=metadata)

    def sync(self, **query):
        self.initial = None
        return super().sync(**query)

    def fetch(self, **query):
        self.initial = None
        return super().fetch(**query)

    def filter_loaded(self, data, *, paginas=None, **_):
        if paginas is not None:
            pages = [paginas] if isinstance(paginas, (int, str)) else paginas
            wanted = {int(self.single_request(pagina=p).period_key) for p in pages}
            data = data[data.index_page.isin(wanted)]
        return data.drop(columns=['_period_key'], errors='ignore')
