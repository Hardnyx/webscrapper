from __future__ import annotations

import gzip
import hashlib
import json
import os
import re
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import pandas as pd

from .exceptions import StorageError


MANIFEST_VERSION = 1


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _safe_component(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._=-]+", "_", str(value))
    return value.strip("._") or "_"


def schema_hash(df: pd.DataFrame) -> str:
    payload = [
        {"name": str(c), "dtype": str(df[c].dtype)}
        for c in df.columns
    ]
    raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def content_hash(df: pd.DataFrame) -> str:
    if df.empty:
        return hashlib.sha256(b"EMPTY").hexdigest()

    work = df.copy()

    # Metadatos de ejecución no deben hacer parecer que cambió la fuente.
    volatile = {
        "retrieved_at",
        "checked_at",
        "_retrieved_at",
        "_checked_at",
    }
    work = work[
        [c for c in work.columns if c not in volatile]
    ]

    work = work.reindex(sorted(work.columns), axis=1)

    # Hash estable independiente del orden de las filas.
    normalized = work.astype("string").fillna("<NA>")
    sort_cols = list(normalized.columns)
    if sort_cols:
        normalized = normalized.sort_values(
            sort_cols,
            kind="mergesort",
        ).reset_index(drop=True)

    raw = normalized.to_csv(
        index=False,
        lineterminator="\n",
    ).encode("utf-8")

    return hashlib.sha256(raw).hexdigest()


class Manifest:
    def __init__(self, path: Path, dataset_id: str):
        self.path = path
        self.dataset_id = dataset_id
        self.data = self._load()

    def _empty(self) -> dict[str, Any]:
        return {
            "manifest_version": MANIFEST_VERSION,
            "dataset_id": self.dataset_id,
            "last_schema_hash": None,
            "entries": {},
        }

    def _load(self) -> dict[str, Any]:
        if not self.path.exists():
            return self._empty()

        try:
            data = json.loads(
                self.path.read_text(encoding="utf-8")
            )
        except Exception as exc:
            raise StorageError(
                f"Manifest inválido: {self.path}"
            ) from exc

        if data.get("dataset_id") != self.dataset_id:
            raise StorageError(
                f"Manifest pertenece a otro dataset: {self.path}"
            )

        data.setdefault("entries", {})
        data.setdefault("last_schema_hash", None)
        return data

    def get(self, period_key: str) -> dict[str, Any] | None:
        return self.data["entries"].get(period_key)

    def set(self, period_key: str, entry: dict[str, Any]):
        self.data["entries"][period_key] = entry

    @property
    def last_schema_hash(self) -> str | None:
        return self.data.get("last_schema_hash")

    @last_schema_hash.setter
    def last_schema_hash(self, value: str | None):
        self.data["last_schema_hash"] = value

    def save(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)

        raw = json.dumps(
            self.data,
            indent=2,
            ensure_ascii=False,
            sort_keys=True,
        )

        tmp = self.path.with_suffix(
            self.path.suffix + ".tmp"
        )
        tmp.write_text(raw, encoding="utf-8")
        os.replace(tmp, self.path)


class DatasetStorage:
    """
    Fuente de verdad local:
      state/manifest.json
      canonical/<partition>/data.parquet
      raw/... opcional

    El Parquet se actualiza por período completo, no por filas individuales.
    """

    def __init__(
        self,
        root: Path,
        dataset_id: str,
    ):
        self.root = Path(root)
        self.dataset_id = dataset_id

        self.state_root = self.root / "state"
        self.canonical_root = self.root / "canonical"
        self.raw_root = self.root / "raw"

        self.manifest = Manifest(
            self.state_root / "manifest.json",
            dataset_id,
        )

    def partition_path(self, partition_key: str) -> Path:
        pieces = [
            _safe_component(x)
            for x in partition_key.strip("/").split("/")
            if x
        ]
        return self.canonical_root.joinpath(*pieces, "data.parquet")

    def relative_path(self, path: Path) -> str:
        """Store paths relative to the dataset so copied data remains portable."""
        return path.relative_to(self.root).as_posix()

    def canonical_has_period(
        self,
        partition_key: str,
        period_key: str,
    ) -> bool:
        path = self.partition_path(partition_key)

        if not path.exists():
            return False

        try:
            df = pd.read_parquet(
                path,
                columns=["_period_key"],
            )
        except Exception:
            return False

        return bool(
            (df["_period_key"].astype(str) == str(period_key)).any()
        )

    def period_matches(
        self, partition_key: str, period_key: str, expected_hash: str | None,
    ) -> bool:
        """A cached partition is reusable only when its saved rows still match the manifest."""
        if not expected_hash:
            return False
        path = self.partition_path(partition_key)
        try:
            frame = pd.read_parquet(path)
            selected = frame[frame["_period_key"].astype(str).eq(str(period_key))]
            if selected.empty:
                return False
            return content_hash(selected.drop(columns="_period_key")) == expected_hash
        except (OSError, KeyError, ValueError, TypeError):
            return False

    def upsert_period(
        self,
        partition_key: str,
        period_key: str,
        data: pd.DataFrame,
    ) -> Path:
        path = self.partition_path(partition_key)
        path.parent.mkdir(parents=True, exist_ok=True)

        incoming = data.copy()
        incoming["_period_key"] = str(period_key)

        if path.exists():
            existing = pd.read_parquet(path)
            if "_period_key" not in existing.columns:
                raise StorageError(
                    f"Partición antigua sin _period_key: {path}"
                )

            existing = existing[
                existing["_period_key"].astype(str) != str(period_key)
            ]

            out = pd.concat(
                [existing, incoming],
                ignore_index=True,
                sort=False,
            )
        else:
            out = incoming

        # Orden reproducible si están presentes.
        preferred = [
            c for c in (
                "entity_type",
                "period_date",
                "period",
                "currency",
                "table_kind",
                "person_type",
                "entity_name",
                "metric",
            )
            if c in out.columns
        ]
        if preferred:
            out = out.sort_values(
                preferred,
                kind="mergesort",
                na_position="last",
            ).reset_index(drop=True)

        tmp = path.with_suffix(".parquet.tmp")
        out.to_parquet(tmp, index=False)
        os.replace(tmp, path)

        return path

    def _read_files(self, files) -> pd.DataFrame:
        # Centralized Parquet loading for full and partition-scoped reads.
        frames = []
        for path in files:
            try:
                frames.append(pd.read_parquet(path))
            except Exception as exc:
                raise StorageError(
                    f"Lectura Parquet fallida: {path}"
                ) from exc

        if not frames:
            return pd.DataFrame()

        return pd.concat(
            frames,
            ignore_index=True,
            sort=False,
        )

    def read_partitions(self, partition_keys) -> pd.DataFrame:
        # Exact partition paths avoid a recursive scan of the complete history.
        files = []
        seen = set()
        for partition_key in partition_keys:
            path = self.partition_path(str(partition_key))
            if path in seen:
                continue
            seen.add(path)
            if path.is_file():
                files.append(path)

        return self._read_files(files)

    def read_all(self) -> pd.DataFrame:
        files = sorted(
            self.canonical_root.rglob("data.parquet")
        )
        return self._read_files(files)

    def write_raw(
        self,
        period_key: str,
        raw: str | bytes,
        *,
        suffix: str = ".html.gz",
    ) -> Path:
        safe = _safe_component(period_key)
        path = self.raw_root / f"{safe}{suffix}"
        path.parent.mkdir(parents=True, exist_ok=True)

        content = (
            raw.encode("utf-8")
            if isinstance(raw, str)
            else raw
        )

        if suffix.endswith(".gz"):
            with gzip.open(path, "wb") as f:
                f.write(content)
        else:
            path.write_bytes(content)

        return path
