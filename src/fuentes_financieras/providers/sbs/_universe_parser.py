"""SBS deposit universe parser, adapted from the validated v4 capture."""
from dataclasses import dataclass
import re
import unicodedata
from bs4 import BeautifulSoup, Tag

COOPAC_MARKERS = (
    "cooperativas de ahorro y credito",
    "coopac",
)

MIN_TOTAL_ENTITIES = 10

MAX_TOTAL_ENTITIES = 100

@dataclass(frozen=True)
class Entity:
    source_order: int
    type_code: str
    entity_type: str
    sbs_name: str
    normalized_name: str

def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", value.replace("\xa0", " ")).strip()

def clean_entity_display_name(value: str) -> str:
    """Remove visual list bullets added by the SBS accordion without altering the legal name."""
    value = clean_text(value)
    value = re.sub(r"^[\-–—•·]+\s*", "", value)
    return clean_text(value)

def normalized_text(value: str) -> str:
    value = clean_text(value).lower()
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()

def normalized_entity_name(value: str) -> str:
    return normalized_text(value).upper()

SECTION_LABELS = {
    "bancos": ("B", "BANCO"),
    "financieras": ("F", "FINANCIERA"),
    "cajas municipales de ahorro y credito": ("C", "CMAC"),
    "cajas rurales de ahorro y credito": ("R", "CRAC"),
}

SKIP_TEXT_PREFIXES = (
    "conoce a las empresas",
    "ingresando aqui",
    "fondo de seguro",
    "fondo de seguros",
    "para mas informacion",
    "las coopac pueden",
    "relacion de entidades",
)

def classify_heading(text: str) -> tuple[str, str] | None:
    normalized = normalized_text(text)
    if any(marker in normalized for marker in COOPAC_MARKERS):
        return None
    for label, result in SECTION_LABELS.items():
        if normalized == label:
            return result
    return None

def is_coopac_heading(text: str) -> bool:
    normalized = normalized_text(text)
    return (
        normalized == "cooperativas de ahorro y credito coopac"
        or normalized.startswith("cooperativas de ahorro y credito")
        or normalized == "coopac"
    )

def _looks_like_entity_name(text: str) -> bool:
    value = clean_text(text)
    norm = normalized_text(value)
    if not value or len(value) < 3 or len(value) > 180:
        return False
    if not norm or any(norm.startswith(prefix) for prefix in SKIP_TEXT_PREFIXES):
        return False
    if norm in SECTION_LABELS or is_coopac_heading(value):
        return False
    if value.lower().startswith(("http://", "https://", "www.")):
        return False

    letters = [ch for ch in value if ch.isalpha()]
    if len(letters) < 3:
        return False

    # SBS currently publishes the company names in uppercase.  Requiring a high
    # uppercase ratio filters navigation/help prose while still allowing
    # punctuation, numbers and legal suffixes such as S.A.
    uppercase_ratio = sum(ch.isupper() for ch in letters) / len(letters)
    return uppercase_ratio >= 0.72

def _entities_from_text_stream(soup: BeautifulSoup) -> list[Entity]:
    """Parse by semantic text boundaries instead of fixed HTML tags.

    The SBS page can place section titles and company names in div/span/a nodes,
    not necessarily in heading/list tags.  Every occurrence of a target section
    label is treated as a candidate block; for each category the block with the
    largest number of plausible company names is retained.  This also protects
    against duplicated section names in navigation menus.
    """

    tokens = [clean_text(str(x)) for x in soup.stripped_strings]
    tokens = [x for x in tokens if x]

    markers: list[tuple[int, str | None, str]] = []
    for idx, token in enumerate(tokens):
        norm = normalized_text(token)
        if is_coopac_heading(token):
            markers.append((idx, None, "COOPAC"))
            continue
        info = SECTION_LABELS.get(norm)
        if info:
            markers.append((idx, info[0], info[1]))

    blocks: dict[str, list[list[str]]] = {"B": [], "F": [], "C": [], "R": []}
    for marker_idx, (idx, code, _name) in enumerate(markers):
        if code is None:
            continue
        end_idx = markers[marker_idx + 1][0] if marker_idx + 1 < len(markers) else len(tokens)
        candidates: list[str] = []
        for token in tokens[idx + 1 : end_idx]:
            entity_name = clean_entity_display_name(token)
            if _looks_like_entity_name(entity_name):
                candidates.append(entity_name)
        blocks[code].append(candidates)

    selected: dict[str, list[str]] = {}
    for code, candidates in blocks.items():
        if candidates:
            selected[code] = max(candidates, key=len)
        else:
            selected[code] = []

    entities: list[Entity] = []
    order = 0
    for code in ("B", "F", "C", "R"):
        entity_type = {"B": "BANCO", "F": "FINANCIERA", "C": "CMAC", "R": "CRAC"}[code]
        seen_names: set[str] = set()
        for name in selected[code]:
            normalized_name = normalized_entity_name(name)
            if normalized_name in seen_names:
                continue
            seen_names.add(normalized_name)
            order += 1
            entities.append(
                Entity(
                    source_order=order,
                    type_code=code,
                    entity_type=entity_type,
                    sbs_name=name,
                    normalized_name=normalized_name,
                )
            )
    return entities

def _entities_from_heading_lists(soup: BeautifulSoup) -> list[Entity]:
    """Legacy/fallback parser for conventional heading + list markup."""
    entities: list[Entity] = []
    current: tuple[str, str] | None = None
    order = 0

    for tag in soup.find_all(["h1", "h2", "h3", "h4", "h5", "h6", "li"]):
        if not isinstance(tag, Tag):
            continue

        if tag.name in {"h1", "h2", "h3", "h4", "h5", "h6"}:
            heading_text = clean_text(tag.get_text(" ", strip=True))
            if is_coopac_heading(heading_text):
                current = None
                continue
            current = classify_heading(heading_text)
            continue

        if tag.name == "li" and current is not None:
            name = clean_entity_display_name(tag.get_text(" ", strip=True))
            if not _looks_like_entity_name(name):
                continue
            type_code, entity_type = current
            order += 1
            entities.append(
                Entity(
                    source_order=order,
                    type_code=type_code,
                    entity_type=entity_type,
                    sbs_name=name,
                    normalized_name=normalized_entity_name(name),
                )
            )
    return entities

def _dedupe_entities(entities: list[Entity]) -> list[Entity]:
    deduped: list[Entity] = []
    seen: set[tuple[str, str]] = set()
    for entity in entities:
        key = (entity.type_code, entity.normalized_name)
        if key in seen:
            continue
        seen.add(key)
        deduped.append(entity)

    # Re-number after de-duplication to keep a contiguous source order.
    return [
        Entity(
            source_order=i,
            type_code=e.type_code,
            entity_type=e.entity_type,
            sbs_name=e.sbs_name,
            normalized_name=e.normalized_name,
        )
        for i, e in enumerate(deduped, start=1)
    ]

def parse_entities(html: str) -> list[Entity]:
    soup = BeautifulSoup(html, "html.parser")

    strategies = [
        ("flujo_texto_semantico", _entities_from_text_stream),
        ("encabezados_y_listas", _entities_from_heading_lists),
    ]
    failures: list[str] = []

    for name, parser in strategies:
        try:
            entities = _dedupe_entities(parser(soup))
            validate_entities(entities)
            return entities
        except Exception as exc:  # noqa: BLE001
            failures.append(f"{name}: {type(exc).__name__}: {exc}")

    body_preview = clean_text(soup.get_text(" ", strip=True))[:1200]
    raise ValueError(
        "No fue posible identificar el universo SBS con ninguna estrategia.\n"
        + "\n".join(f"  - {item}" for item in failures)
        + f"\nVista previa del texto renderizado: {body_preview!r}"
    )

def validate_entities(entities: list[Entity]) -> None:
    if not (MIN_TOTAL_ENTITIES <= len(entities) <= MAX_TOTAL_ENTITIES):
        raise ValueError(
            f"Conteo inesperado: {len(entities)} entidades. "
            "La estructura de la página SBS pudo haber cambiado."
        )

    counts: dict[str, int] = {code: 0 for code in ("B", "F", "C", "R")}
    for entity in entities:
        counts[entity.type_code] = counts.get(entity.type_code, 0) + 1
        if "COOPERATIVA DE AHORRO Y CREDITO" in entity.normalized_name:
            raise ValueError(
                "Se detectó una COOPAC dentro del universo objetivo. "
                "Se detiene la ejecución para evitar contaminar el maestro."
            )

    missing = [code for code, count in counts.items() if count == 0]
    if missing:
        raise ValueError(
            f"No se encontraron entidades para las categorías: {', '.join(missing)}. "
            "La estructura de la página SBS pudo haber cambiado."
        )

