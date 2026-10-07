from __future__ import annotations

import json
import os
from pathlib import Path

import pandas as pd

from .exceptions import StorageError

ENV_DATA_ROOT = "FINANCIAL_SOURCES_DATA_ROOT"
ENV_AUTOMATIZACIONES_ROOT = "AUTOMATIZACIONES_ROOT"


def _read_automatizaciones_structure(path: Path) -> dict | None:
    marker = path / "estructura.json"
    if not marker.is_file():
        return None
    try:
        data = json.loads(marker.read_text(encoding="utf-8"))
    except Exception:
        return None
    if data.get("formato") != "automatizaciones":
        return None
    return data


def _valid_automatizaciones_root(path: Path) -> bool:
    return _read_automatizaciones_structure(path) is not None


def find_automatizaciones_root(start: Path | None = None) -> Path | None:
    explicit = os.getenv(ENV_AUTOMATIZACIONES_ROOT)
    if explicit:
        root = Path(explicit).expanduser().resolve()
        if not _valid_automatizaciones_root(root):
            raise StorageError(
                f"Raíz inválida en {ENV_AUTOMATIZACIONES_ROOT}: {root}"
            )
        return root

    start = (start or Path.cwd()).resolve()
    for candidate in (start, *start.parents):
        if _valid_automatizaciones_root(candidate):
            return candidate

    here = Path(__file__).resolve()
    for candidate in here.parents:
        if _valid_automatizaciones_root(candidate):
            return candidate
    return None


def _metadata_path(meta) -> str | None:
    if isinstance(meta, str) and meta.strip():
        return meta.strip()
    if isinstance(meta, dict):
        value = meta.get("ruta")
        if isinstance(value, str) and value.strip():
            return value.strip()
    return None


def _declared_data_roots(root: Path, structure: dict) -> list[Path]:
    candidates: list[Path] = []
    for container_key in ("datos", "data", "almacenes", "storage"):
        container = structure.get(container_key)
        if not isinstance(container, dict):
            continue
        for key in ("fuentes", "sources"):
            route = _metadata_path(container.get(key))
            if route:
                candidates.append((root / route).resolve())

    for direct_key in ("ruta_datos_fuentes", "sources_data_root", "data_root"):
        route = structure.get(direct_key)
        if isinstance(route, str) and route.strip():
            candidates.append((root / route.strip()).resolve())

    return list(dict.fromkeys(candidates))


def _internal_package_data_root(root: Path, structure: dict) -> Path:
    libraries = structure.get("librerias", {})
    route = _metadata_path(libraries.get("fuentes")) if isinstance(libraries, dict) else None
    library_root = root / (route or "librerias/fuentes")
    return (library_root / "src" / "fuentes_financieras" / "data" / "sources").resolve()


def _central_data_root(root: Path) -> Path:
    return (root / "datos" / "fuentes").resolve()


def _mutual_fund_dataset_root(data_root: Path) -> Path:
    return data_root / "peru" / "smv" / "fondos_mutuos" / "valores_cuota"


def _has_mutual_fund_store(path: Path) -> bool:
    dataset = _mutual_fund_dataset_root(path)
    return (dataset / "canonical").is_dir() or (dataset / "state" / "manifest.json").is_file()


def _snapshot_store_stats(data_root: Path) -> dict:
    """Read lightweight coverage metadata without scanning the Parquet tree."""
    dataset = _mutual_fund_dataset_root(data_root)
    manifest_path = dataset / "state" / "manifest.json"
    stats = {
        "path": data_root.resolve(),
        "exists": _has_mutual_fund_store(data_root),
        "validated_snapshots": 0,
        "latest_snapshot": None,
        "latest_snapshot_file": None,
        "latest_value_date": None,
    }
    if not manifest_path.is_file():
        return stats

    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except Exception:
        return stats

    validated: list[tuple[str, dict]] = []
    for key, entry in (manifest.get("entries") or {}).items():
        if not str(key).startswith("snapshot:"):
            continue
        if not isinstance(entry, dict) or entry.get("status") != "validated":
            continue
        validated.append((str(key), entry))

    stats["validated_snapshots"] = len(validated)
    if not validated:
        return stats

    # Selección del almacén por fecha efectiva del VC antes que por fecha de
    # consulta. Inspección limitada a las últimas particiones validadas para
    # detectar capturas recientes con valores cuota antiguos.
    latest_value = None
    for key, entry in sorted(validated, key=lambda item: item[0], reverse=True)[:31]:
        canonical_rel = entry.get("canonical_path")
        if canonical_rel:
            candidate = dataset / str(canonical_rel)
        else:
            partition = entry.get("partition_key")
            if not partition:
                continue
            candidate = dataset / "canonical" / str(partition) / "data.parquet"
        if not candidate.is_file():
            continue
        if stats["latest_snapshot"] is None:
            stats["latest_snapshot"] = key.removeprefix("snapshot:")
            stats["latest_snapshot_file"] = candidate
        try:
            dates = pd.read_parquet(candidate, columns=["value_date"])["value_date"]
            effective = pd.to_datetime(dates, errors="coerce").max()
        except Exception:
            effective = pd.NaT
        if pd.notna(effective):
            effective = pd.Timestamp(effective).normalize()
            if latest_value is None or effective > latest_value:
                latest_value = effective

    if latest_value is not None:
        stats["latest_value_date"] = latest_value.strftime("%Y-%m-%d")

    return stats


def describe_data_roots(root: Path | None = None) -> list[dict]:
    automations_root = root or find_automatizaciones_root(Path(__file__).resolve())
    if automations_root is None:
        return []

    structure = _read_automatizaciones_structure(automations_root) or {}
    declared = _declared_data_roots(automations_root, structure)
    central = _central_data_root(automations_root)
    package_data = _internal_package_data_root(automations_root, structure)

    ordered: list[tuple[str, Path, int]] = []
    for path in declared:
        ordered.append(("declarado", path, 3))
    ordered.extend([
        ("central", central, 2),
        ("interno", package_data, 1),
    ])

    seen: set[Path] = set()
    result: list[dict] = []
    for label, path, tie_priority in ordered:
        path = path.resolve()
        if path in seen:
            continue
        seen.add(path)
        stats = _snapshot_store_stats(path)
        stats["label"] = label
        stats["tie_priority"] = tie_priority
        result.append(stats)
    return result


def _store_sort_key(info: dict) -> tuple[str, str, int, int]:
    effective = str(info.get("latest_value_date") or "0000-00-00")
    latest = str(info.get("latest_snapshot") or "0000-00-00")
    count = int(info.get("validated_snapshots") or 0)
    priority = int(info.get("tie_priority") or 0)
    return effective, latest, count, priority


def find_standalone_root() -> Path:
    here = Path(__file__).resolve()
    for candidate in here.parents:
        if (candidate / ".fuentes_financieras_root").is_file():
            return candidate
    raise StorageError("Marcador .fuentes_financieras_root no localizado.")


def resolve_data_root(explicit: str | Path | None = None) -> Path:
    if explicit:
        return Path(explicit).expanduser().resolve()

    env = os.getenv(ENV_DATA_ROOT)
    if env:
        return Path(env).expanduser().resolve()

    automations_root = find_automatizaciones_root(Path(__file__).resolve())
    if automations_root is not None:
        structure = _read_automatizaciones_structure(automations_root) or {}
        candidates = describe_data_roots(automations_root)
        populated = [item for item in candidates if item.get("latest_snapshot")]
        if populated:
            return max(populated, key=_store_sort_key)["path"]

        existing = [item for item in candidates if item.get("exists")]
        if existing:
            return max(existing, key=lambda item: int(item.get("tie_priority") or 0))["path"]

        declared = _declared_data_roots(automations_root, structure)
        if declared:
            return declared[0]

        return _central_data_root(automations_root)

    return (find_standalone_root() / "data" / "sources").resolve()
