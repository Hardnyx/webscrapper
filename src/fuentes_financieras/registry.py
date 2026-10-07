from __future__ import annotations

import importlib
from functools import lru_cache

from .catalog import CATALOG
from .exceptions import UnknownDatasetError


def get_spec(dataset_id: str):
    try:
        return CATALOG[dataset_id]
    except KeyError as exc:
        raise UnknownDatasetError(
            f"Dataset no registrado: {dataset_id!r}. "
            f"Disponibles: {', '.join(sorted(CATALOG))}"
        ) from exc


def _load_symbol(path: str):
    module_name, symbol_name = path.split(":", 1)
    module = importlib.import_module(module_name)
    return getattr(module, symbol_name)


@lru_cache(maxsize=None)
def get_provider(dataset_id: str):
    spec = get_spec(dataset_id)
    cls = _load_symbol(spec.provider_class)
    return cls(spec)


def list_specs():
    return [CATALOG[k] for k in sorted(CATALOG)]
