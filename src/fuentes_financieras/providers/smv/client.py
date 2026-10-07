from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
import time
from typing import Literal
from urllib.parse import urlencode
from urllib.request import urlopen

import requests
from requests import Response

from .errors import SMVRequestError

Transport = Literal["auto", "direct", "requests", "browser"]


@dataclass(frozen=True)
class FetchResult:
    html: str
    transport: Literal["direct", "requests", "browser"]
    status_code: int


class SMVClient:
    """Cliente HTTP para consultar valor cuota de fondos mutuos en SMV.

    El modo ``auto`` usa primero una sesión HTTP normal. Si la página responde
    con un bloqueo o la consulta falla, intenta nuevamente mediante Chromium
    headless. El navegador se abre de forma diferida y se reutiliza durante toda
    la sesión para evitar iniciar un proceso por cada fecha consultada.
    """

    HOME_URL = "https://www.smv.gob.pe/SIMV/Frm_ValorCuota"
    DETAIL_URL = "https://www.smv.gob.pe/SIMV/Frm_ValorCuotaDetalle_V2.aspx"
    FALLBACK_STATUS = {401, 403, 429}

    def __init__(
        self,
        *,
        transport: Transport = "auto",
        timeout: float = 30.0,
        max_retries: int = 2,
        retry_wait: float = 1.0,
        headless: bool = True,
    ) -> None:
        if transport not in {"auto", "direct", "requests", "browser"}:
            raise ValueError("transport fuera de dominio: auto, direct, requests o browser.")
        if timeout <= 0:
            raise ValueError("timeout fuera de dominio: valor positivo requerido.")
        if max_retries < 0:
            raise ValueError("max_retries fuera de dominio: valor no negativo requerido.")
        if retry_wait < 0:
            raise ValueError("retry_wait fuera de dominio: valor no negativo requerido.")

        self.transport = transport
        self.timeout = timeout
        self.max_retries = max_retries
        self.retry_wait = retry_wait
        self.headless = headless

        self._session = requests.Session()
        self._session.headers.update(
            {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/152.0.0.0 Safari/537.36"
                ),
                "Accept": (
                    "text/html,application/xhtml+xml,application/xml;q=0.9,"
                    "image/avif,image/webp,*/*;q=0.8"
                ),
                "Accept-Language": "es-PE,es;q=0.9,en;q=0.8",
                "Referer": self.HOME_URL,
            }
        )
        self._session_initialized = False

        self._playwright = None
        self._browser = None
        self._context = None
        self._page = None

    def __enter__(self) -> "SMVClient":
        return self

    def __exit__(self, exc_type, exc, traceback) -> None:
        self.close()

    @staticmethod
    def _format_date(value: str | date | datetime) -> str:
        if isinstance(value, datetime):
            parsed = value.date()
        elif isinstance(value, date):
            parsed = value
        else:
            parsed = datetime.strptime(value, "%d/%m/%Y").date()
        return parsed.strftime("%d/%m/%Y")

    @classmethod
    def _params(cls, value: str | date | datetime) -> dict[str, str]:
        return {
            "in_ac_pre_ope": "O",
            "tip_fon_desc": "FONDO OPERATIVO",
            "in_ad_fecha": cls._format_date(value),
        }

    @classmethod
    def detail_url(cls, value: str | date | datetime) -> str:
        day = cls._format_date(value)
        return (f"{cls.DETAIL_URL}?in_ac_pre_ope=O"
                f"&tip_fon_desc=FONDO%20OPERATIVO&in_ad_fecha={day}")

    def fetch(self, value: str | date | datetime, *, transport: Transport | None = None) -> FetchResult:
        """Recupera el HTML de una fecha mediante el transporte solicitado."""
        selected = transport or self.transport
        if selected == "direct":
            return self._fetch_direct(value)
        if selected == "requests":
            return self._fetch_requests(value)
        if selected == "browser":
            return self._fetch_browser(value)

        try:
            return self._fetch_direct(value)
        except SMVRequestError as direct_error:
            try:
                return self._fetch_requests(value)
            except SMVRequestError as requests_error:
                try:
                    return self._fetch_browser(value)
                except SMVRequestError as browser_error:
                    raise SMVRequestError(
                        f"Consulta directa: {direct_error}; requests: {requests_error}; "
                        f"navegador: {browser_error}"
                    ) from browser_error

    def _fetch_direct(self, value: str | date | datetime) -> FetchResult:
        """Consulta como pd.read_html(url) y conserva el HTML para auditar."""
        url = self.detail_url(value)
        try:
            with urlopen(url, timeout=self.timeout) as response:
                html = response.read().decode(
                    response.headers.get_content_charset() or "utf-8", errors="replace"
                )
                status = response.status
            if not html.strip():
                raise SMVRequestError("SMV devolvió una respuesta vacía.")
            return FetchResult(html, "direct", status)
        except SMVRequestError:
            raise
        except Exception as error:
            raise SMVRequestError(f"No se pudo consultar la URL directa {url}: {error}") from error

    def _initialize_session(self) -> None:
        if self._session_initialized:
            return

        try:
            response = self._session.get(self.HOME_URL, timeout=self.timeout)
            if response.status_code < 500:
                self._session_initialized = True
        except requests.RequestException:
            # The detail request may still work even if the warm-up request fails.
            self._session_initialized = True

    def _fetch_requests(self, value: str | date | datetime) -> FetchResult:
        last_error: Exception | None = None

        for attempt in range(self.max_retries + 1):
            try:
                response = self._session.get(
                    self.DETAIL_URL,
                    params=self._params(value),
                    timeout=self.timeout,
                )
                self._validate_response(response)
                return FetchResult(response.text, "requests", response.status_code)
            except (requests.RequestException, SMVRequestError) as error:
                last_error = error
                if attempt < self.max_retries:
                    time.sleep(self.retry_wait * (attempt + 1))

        raise SMVRequestError(
            f"No se pudo consultar SMV mediante requests: {last_error}"
        ) from last_error

    def _validate_response(self, response: Response) -> None:
        if response.status_code in self.FALLBACK_STATUS:
            raise SMVRequestError(
                f"SMV respondió HTTP {response.status_code}."
            )
        if response.status_code >= 400:
            raise SMVRequestError(
                f"SMV respondió HTTP {response.status_code}."
            )
        if not response.text.strip():
            raise SMVRequestError("SMV devolvió una respuesta vacía.")

    def _ensure_browser(self) -> None:
        if self._page is not None:
            return

        try:
            from playwright.sync_api import sync_playwright
        except ImportError as error:
            raise SMVRequestError(
                "Fallback headless no disponible: dependencia Playwright ausente. Dependencias requeridas: "
                "de fuentes antes de usar transport='browser' o transport='auto'."
            ) from error

        try:
            self._playwright = sync_playwright().start()
            self._browser = self._playwright.chromium.launch(headless=self.headless)
            self._context = self._browser.new_context(
                locale="es-PE",
                timezone_id="America/Lima",
            )
            self._page = self._context.new_page()
            self._page.goto(
                self.HOME_URL,
                wait_until="domcontentloaded",
                timeout=int(self.timeout * 1000),
            )
        except Exception as error:
            self.close()
            raise SMVRequestError(
                "No se pudo iniciar Chromium para consultar SMV. "
                "Verifique que Playwright tenga Chromium instalado."
            ) from error

    def _fetch_browser(self, value: str | date | datetime) -> FetchResult:
        self._ensure_browser()
        assert self._page is not None

        try:
            # The detail page accepts the date in the query string, so no visual
            # interaction is required to select it.
            url = f"{self.DETAIL_URL}?{urlencode(self._params(value))}"
            response = self._page.goto(
                url,
                wait_until="domcontentloaded",
                timeout=int(self.timeout * 1000),
                referer=self.HOME_URL,
            )
            if response is None:
                raise SMVRequestError("SMV no devolvió respuesta al navegador.")
            if response.status >= 400:
                raise SMVRequestError(f"SMV respondió HTTP {response.status}.")
            html = self._page.content()
            if not html.strip():
                raise SMVRequestError("SMV devolvió una respuesta vacía.")
            return FetchResult(html, "browser", response.status)
        except SMVRequestError:
            raise
        except Exception as error:
            raise SMVRequestError(
                f"No se pudo consultar SMV mediante Chromium: {error}"
            ) from error

    def close(self) -> None:
        """Cierra sesión HTTP y recursos del navegador, si fueron abiertos."""
        self._session.close()

        if self._page is not None:
            try:
                self._page.close()
            except Exception:
                pass
        if self._context is not None:
            try:
                self._context.close()
            except Exception:
                pass
        if self._browser is not None:
            try:
                self._browser.close()
            except Exception:
                pass
        if self._playwright is not None:
            try:
                self._playwright.stop()
            except Exception:
                pass

        self._page = None
        self._context = None
        self._browser = None
        self._playwright = None
