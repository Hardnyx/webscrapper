from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd


@dataclass(frozen=True)
class DatasetDescription:
    dataset_id: str
    title: str
    provider: str
    country: str | None
    organization: str
    frequency: str
    storage_path: str
    network_transport: str
    notes: str = ""


@dataclass
class FetchResult:
    dataset_id: str
    data: pd.DataFrame
    metadata: dict[str, Any] = field(default_factory=dict)
    raw: str | bytes | None = None


@dataclass
class SyncResult:
    dataset_id: str
    requested: int = 0
    downloaded: int = 0
    unchanged: int = 0
    skipped_existing: int = 0
    unavailable: int = 0
    failed: int = 0
    canonical_root: Path | None = None
    details: list[dict[str, Any]] = field(default_factory=list)


@dataclass(frozen=True)
class PeriodRequest:
    period_key: str
    partition_key: str
    params: dict[str, Any]
    mutable: bool = False
