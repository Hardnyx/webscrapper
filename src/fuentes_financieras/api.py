from __future__ import annotations

from .models import DatasetDescription
from .registry import get_provider, get_spec, list_specs


def source(dataset_id: str):
    return get_provider(dataset_id)


def fetch(dataset_id: str, **query):
    return source(dataset_id).fetch(**query)


def sync(dataset_id: str, **query):
    return source(dataset_id).sync(**query)


def load(dataset_id: str, **query):
    return source(dataset_id).load(**query)


def describe(dataset_id: str) -> DatasetDescription:
    s = get_spec(dataset_id)
    return DatasetDescription(
        dataset_id=s.dataset_id,
        title=s.title,
        provider=s.provider_class,
        country=s.country,
        organization=s.organization,
        frequency=s.frequency,
        storage_path=s.storage_path,
        network_transport=s.network_transport,
        notes=s.notes,
    )


def list_datasets() -> list[DatasetDescription]:
    return [describe(s.dataset_id) for s in list_specs()]
