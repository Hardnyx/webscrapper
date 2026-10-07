from __future__ import annotations

import re
from bs4 import BeautifulSoup, FeatureNotFound


def _make_soup(html: str):
    try:
        return BeautifulSoup(html, "lxml")
    except FeatureNotFound:
        # Fallback robusto: la lógica WebForms no depende de lxml.
        return BeautifulSoup(html, "html.parser")


def successful_controls(html: str) -> dict[str, str]:
    """
    Emula los controles que un navegador envía en un submit HTML.

    No incluye todos los type=submit: el provider añade solamente el botón
    que originó la acción.
    """
    soup = _make_soup(html)
    form = soup.find("form") or soup
    data: dict[str, str] = {}

    for inp in form.find_all("input"):
        name = inp.get("name")
        if not name or inp.has_attr("disabled"):
            continue

        typ = (inp.get("type") or "text").lower()

        if typ in {"submit", "button", "image", "reset", "file"}:
            continue

        if typ in {"checkbox", "radio"} and not inp.has_attr("checked"):
            continue

        data[name] = inp.get("value", "")

    for sel in form.find_all("select"):
        name = sel.get("name")
        if not name or sel.has_attr("disabled"):
            continue

        opt = (
            sel.find("option", selected=True)
            or sel.find("option")
        )
        data[name] = opt.get("value", "") if opt else ""

    for ta in form.find_all("textarea"):
        name = ta.get("name")
        if name and not ta.has_attr("disabled"):
            data[name] = ta.get_text()

    return data


def set_suffix(
    data: dict[str, str],
    suffix: str,
    value: str,
    default_name: str | None = None,
):
    for key in list(data):
        if key.endswith(suffix):
            data[key] = value
            return key

    if default_name:
        data[default_name] = value
        return default_name

    return None


def parse_delta_response(text: str) -> tuple[str, dict[str, str]]:
    marker = "|updatePanel|"
    pos = text.find(marker)

    if pos < 0:
        # Algunos postbacks pueden devolver HTML completo.
        if "<html" in text.lower() or "<table" in text.lower():
            return text, {}
        raise ValueError("Respuesta sin updatePanel ASP.NET.")

    panel_name_end = text.find("|", pos + len(marker))
    if panel_name_end < 0:
        raise ValueError("Respuesta delta malformada.")

    content_start = panel_name_end + 1

    boundary = re.search(
        r"\|\d+\|hiddenField\|__EVENTTARGET\|",
        text[content_start:],
    )

    if boundary:
        content_end = content_start + boundary.start()
    else:
        boundary = re.search(
            r"\|\d+\|hiddenField\|__VIEWSTATE\|",
            text[content_start:],
        )
        if not boundary:
            raise ValueError(
                "No se encontró el final del updatePanel ASP.NET."
            )
        content_end = content_start + boundary.start()

    html = text[content_start:content_end]

    hidden: dict[str, str] = {}
    for m in re.finditer(
        r"\|\d+\|hiddenField\|([^|]+)\|([^|]*)\|",
        text[content_end:],
    ):
        hidden[m.group(1)] = m.group(2)

    return html, hidden


def merge_hidden_state(
    state: dict[str, str],
    html: str,
    hidden: dict[str, str],
) -> dict[str, str]:
    out = dict(state)

    panel_controls = successful_controls(html)
    for k, v in panel_controls.items():
        if k.startswith("ctl00") or k.startswith("__"):
            out[k] = v

    out.update(hidden)
    return out
