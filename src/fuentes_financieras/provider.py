from __future__ import annotations

from abc import ABC, abstractmethod
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd

from .exceptions import SchemaChangedError, PeriodUnavailableError
from .models import FetchResult, PeriodRequest, SyncResult
from .runtime import resolve_data_root
from .storage import DatasetStorage, content_hash, schema_hash, utc_now_iso


class DatasetProvider(ABC):
    """
    Contrato de un provider.

    La capa genérica gestiona:
      - caché;
      - manifest;
      - hashes;
      - upsert canónico;
      - fetch/sync/load.

    El provider decide:
      - cómo habla con la fuente;
      - qué es un período;
      - cómo normaliza;
      - qué períodos siguen siendo mutables.
    """

    parser_version = "1"
    contract_version = "1"

    def __init__(self, spec):
        self.spec = spec

        data_root = resolve_data_root()
        self.storage = DatasetStorage(
            data_root / spec.storage_path,
            spec.dataset_id,
        )

    def fetch(self, **query) -> FetchResult:
        request = self.single_request(**query)
        return self._fetch_period(request)

    def sync(
        self,
        *,
        force: bool = False,
        refresh_hours: float = 24.0,
        keep_raw: bool = False,
        allow_schema_change: bool = False,
        **query,
    ) -> SyncResult:
        requests = list(self.plan_sync(**query))

        result = SyncResult(
            dataset_id=self.spec.dataset_id,
            requested=len(requests),
            canonical_root=self.storage.canonical_root,
        )

        for req in requests:
            previous = self.storage.manifest.get(
                req.period_key
            )

            previous_status = (
                previous.get("status")
                if previous else None
            )

            cached_validated = (
                previous_status == "validated"
                and previous.get("contract_version", "1")
                    == self.contract_version
                and self.storage.period_matches(
                    req.partition_key,
                    req.period_key,
                    previous.get("content_hash"),
                )
            )

            cached_unavailable = (
                previous_status == "unavailable"
                and previous.get("contract_version", "1")
                    == self.contract_version
            )

            if (
                not force
                and previous
                and (cached_validated or cached_unavailable)
                and not self._refresh_due(
                    req,
                    previous,
                    refresh_hours,
                )
            ):
                result.skipped_existing += 1
                result.details.append({
                    "period_key": req.period_key,
                    "status": (
                        "skipped_unavailable"
                        if cached_unavailable
                        else "skipped_existing"
                    ),
                })
                continue

            try:
                fetched = self._fetch_period(req)
            except PeriodUnavailableError as exc:
                result.unavailable += 1
                self.storage.manifest.set(
                    req.period_key,
                    {
                        "status": "unavailable",
                        "period_key": req.period_key,
                        "partition_key": req.partition_key,
                        "parser_version": self.parser_version,
                        "contract_version": self.contract_version,
                        "checked_at": utc_now_iso(),
                        "error": str(exc),
                    },
                )
                self.storage.manifest.save()
                result.details.append({
                    "period_key": req.period_key,
                    "status": "unavailable",
                    "error": str(exc),
                })
                continue
            except Exception as exc:
                result.failed += 1
                result.details.append({
                    "period_key": req.period_key,
                    "status": "failed",
                    "error": repr(exc),
                })
                continue

            data = fetched.data.copy()

            sh = schema_hash(data)
            ch = content_hash(data)
            now = utc_now_iso()

            last_schema = self.storage.manifest.last_schema_hash

            if (
                last_schema
                and sh != last_schema
                and not allow_schema_change
            ):
                raise SchemaChangedError(
                    f"Schema hash cambió para {self.spec.dataset_id}: "
                    f"{last_schema[:12]} -> {sh[:12]}. "
                    "allow_schema_change=True requiere revisión previa del cambio de esquema."
                )

            same_content = (
                previous
                and previous.get("content_hash") == ch
                and self.storage.period_matches(
                    req.partition_key,
                    req.period_key,
                    previous.get("content_hash"),
                )
            )

            if same_content:
                result.unchanged += 1
                canonical_path = self.storage.partition_path(
                    req.partition_key
                )
            else:
                canonical_path = self.storage.upsert_period(
                    req.partition_key,
                    req.period_key,
                    data,
                )
                result.downloaded += 1

            raw_path = None
            if keep_raw and fetched.raw is not None:
                raw_path = self.storage.write_raw(
                    req.period_key,
                    fetched.raw,
                )

            entry = {
                "status": "validated",
                "period_key": req.period_key,
                "partition_key": req.partition_key,
                "parser_version": self.parser_version,
                "contract_version": self.contract_version,
                "schema_hash": sh,
                "content_hash": ch,
                "rows": int(len(data)),
                "fetched_at": (
                    previous.get("fetched_at")
                    if same_content and previous
                    else now
                ),
                "checked_at": now,
                "canonical_path": self.storage.relative_path(canonical_path),
                "raw_path": self.storage.relative_path(raw_path) if raw_path else (
                    previous.get("raw_path")
                    if previous else None
                ),
                "metadata": fetched.metadata,
            }

            self.storage.manifest.set(
                req.period_key,
                entry,
            )
            self.storage.manifest.last_schema_hash = sh
            self.storage.manifest.save()

            result.details.append({
                "period_key": req.period_key,
                "status": (
                    "unchanged"
                    if same_content
                    else "downloaded"
                ),
                "rows": int(len(data)),
                "schema_hash": sh,
                "content_hash": ch,
            })

        return result

    def load(self, **query) -> pd.DataFrame:
        partition_keys = self.partition_keys_for_load(**query)
        if partition_keys is None:
            data = self.storage.read_all()
        else:
            data = self.storage.read_partitions(partition_keys)

        if data.empty:
            return data

        return self.filter_loaded(
            data,
            **query,
        ).reset_index(drop=True)

    def partition_keys_for_load(self, **query):
        # None preserves full-history behavior for providers without partition pruning.
        return None

    def _refresh_due(
        self,
        req: PeriodRequest,
        previous: dict[str, Any],
        refresh_hours: float,
    ) -> bool:
        if not req.mutable:
            return False

        checked = previous.get("checked_at")
        if not checked:
            return True

        try:
            checked_dt = datetime.fromisoformat(
                checked.replace("Z", "+00:00")
            )
        except Exception:
            return True

        age = (
            datetime.now(timezone.utc) - checked_dt
        ).total_seconds() / 3600.0

        return age >= float(refresh_hours)

    @abstractmethod
    def single_request(self, **query) -> PeriodRequest:
        ...

    @abstractmethod
    def plan_sync(self, **query):
        ...

    @abstractmethod
    def _fetch_period(
        self,
        request: PeriodRequest,
    ) -> FetchResult:
        ...

    def filter_loaded(
        self,
        data: pd.DataFrame,
        **query,
    ) -> pd.DataFrame:
        return data
