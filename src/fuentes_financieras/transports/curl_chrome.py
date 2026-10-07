from __future__ import annotations

import ssl
from pathlib import Path

import certifi
from curl_cffi import requests

from fuentes_financieras.exceptions import WAFBlockedError, SourceUnavailableError


WAF_HARD_MARKERS = (
    "request unsuccessful",
    "access denied",
    "error 15",
    'id="main-iframe"',
    "incident id",
    "powered by incapsula",
    "request blocked",
)


def build_ca_bundle(target: Path) -> Path:
    """
    certifi + ROOT de Windows.

    Sirve para entornos corporativos con inspección HTTPS sin desactivar TLS.
    """
    target.parent.mkdir(parents=True, exist_ok=True)

    source = Path(certifi.where())
    content = source.read_text(
        encoding="ascii",
        errors="ignore",
    )

    existing = set()
    import re
    for block in re.findall(
        r"-----BEGIN CERTIFICATE-----.*?-----END CERTIFICATE-----",
        content,
        flags=re.S,
    ):
        existing.add(block.strip())

    if hasattr(ssl, "enum_certificates"):
        try:
            for der, encoding, trust in ssl.enum_certificates("ROOT"):
                if encoding != "x509_asn":
                    continue
                try:
                    pem = ssl.DER_cert_to_PEM_cert(der).strip()
                except Exception:
                    continue
                if pem not in existing:
                    existing.add(pem)
                    content += "\n" + pem + "\n"
        except Exception:
            pass

    target.write_text(content, encoding="ascii")
    return target


def looks_like_waf(text: str, status: int, *, normal_marker: str | None = None) -> bool:
    if normal_marker and normal_marker in (text or ""):
        return False

    low = (text or "").lower()

    if status in {403, 406, 429, 503}:
        return True

    return any(marker in low for marker in WAF_HARD_MARKERS)


class CurlChromeTransport:
    def __init__(
        self,
        *,
        state_dir: Path,
        normal_marker: str | None = None,
        timeout: float = 60.0,
    ):
        self.normal_marker = normal_marker
        self.timeout = timeout

        self.ca_bundle = build_ca_bundle(
            Path(state_dir) / "ca_windows_mas_certifi.pem"
        )

        self.session = requests.Session(
            impersonate="chrome"
        )

    def reset(self):
        self.session = requests.Session(
            impersonate="chrome"
        )

    def request(self, method: str, url: str, **kwargs):
        headers = dict(kwargs.pop("headers", {}) or {})
        headers.setdefault(
            "X-Requested-With",
            "XMLHttpRequest",
        )

        response = self.session.request(
            method,
            url,
            headers=headers,
            timeout=kwargs.pop("timeout", self.timeout),
            verify=str(self.ca_bundle),
            **kwargs,
        )

        if looks_like_waf(
            response.text,
            response.status_code,
            normal_marker=self.normal_marker,
        ):
            raise WAFBlockedError(
                f"WAF/Imperva: HTTP {response.status_code}, "
                f"bytes={len(response.content)}."
            )

        if response.status_code >= 400:
            raise SourceUnavailableError(
                f"HTTP {response.status_code} en {url}"
            )

        return response
