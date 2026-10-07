"""Read-only migration of verified SMV captures from Automatizaciones."""
from __future__ import annotations

from collections import Counter
from hashlib import sha256
from io import StringIO
import json
from pathlib import Path

import pandas as pd

from fuentes_financieras.storage import content_hash, schema_hash, utc_now_iso
from .client import SMVClient
from .fondos_mutuos import keys, normalize_snapshot


def _same_rows(actual: pd.DataFrame, expected: pd.DataFrame) -> bool:
    excluded = {"query_date", "source", "transport"}
    a = actual.drop(columns=excluded, errors="ignore")
    b = expected.drop(columns=excluded, errors="ignore")
    if len(a) != len(b) or set(a.columns) != set(b.columns):
        return False
    columns = sorted(a.columns)
    a = pd.read_csv(StringIO(a[columns].to_csv(index=False, date_format="%Y-%m-%d")))
    b = pd.read_csv(StringIO(b[columns].to_csv(index=False, date_format="%Y-%m-%d")))
    for col in columns:
        if pd.api.types.is_numeric_dtype(b[col]):
            a[col] = pd.to_numeric(a[col], errors="coerce")
    return Counter(pd.util.hash_pandas_object(a, index=False).tolist()) == Counter(
        pd.util.hash_pandas_object(b, index=False).tolist()
    )


def _store(provider, key: str, partition: str, data: pd.DataFrame,
           metadata: dict, raw: bytes | None = None, raw_suffix: str = ".html.gz") -> None:
    path = provider.storage.upsert_period(partition, key, data)
    raw_path = provider.storage.write_raw(key, raw, suffix=raw_suffix) if raw is not None else None
    provider.storage.manifest.set(key, {
        "status": "validated", "period_key": key, "partition_key": partition,
        "parser_version": provider.parser_version,
        "contract_version": provider.contract_version,
        "schema_hash": schema_hash(data), "content_hash": content_hash(data),
        "rows": len(data), "checked_at": utc_now_iso(),
        "canonical_path": provider.storage.relative_path(path),
        "raw_path": provider.storage.relative_path(raw_path) if raw_path else None,
        "metadata": metadata,
    })
    if key.startswith("snapshot:"):
        provider.storage.manifest.last_schema_hash = schema_hash(data)
    provider.storage.manifest.save()


def migrate_legacy(provider, automations_root: str | Path,
                   *, include_exports: bool = True) -> dict[str, int]:
    """Import all rows of verified daily snapshots and any archived range exports.

    Source files remain untouched. Invalid captures are skipped and will be
    fetched on the next sync; existing validated Parquet is reused.
    """
    old = (Path(automations_root).expanduser().resolve()
           / "datos/valor_cuota/peru/smv/fondos_mutuos")
    manifest = old / "_manifest/loads.jsonl"
    records: dict[str, dict] = {}
    if manifest.is_file():
        for line in manifest.read_text(encoding="utf-8").splitlines():
            if line.strip():
                item = json.loads(line)
                if item.get("status") == "ok":
                    records[item.get("partition", "")] = item
    result = {"daily_imported": 0, "daily_reused": 0,
              "daily_invalid": 0, "historical_rows_imported": 0}
    for file in sorted((old / "canonical").rglob("????-??-??.csv")):
        day = pd.Timestamp(file.stem)
        key, partition = keys(day)
        prior = provider.storage.manifest.get(key)
        if prior and provider.storage.period_matches(partition, key, prior.get("content_hash")):
            result["daily_reused"] += 1
            continue
        raw = old / "raw" / f"{day:%Y}" / f"{day:%m}" / f"{day:%Y-%m-%d}.html"
        item = records.get(f"{day:%Y-%m-%d}")
        try:
            if not raw.is_file() or item is None or sha256(raw.read_bytes()).hexdigest() != item.get("sha256"):
                raise ValueError("HTML/manifest missing or inconsistent")
            expected = normalize_snapshot(raw.read_text(encoding="utf-8"), day,
                                          item.get("transport") or "direct")
            actual = pd.read_csv(file, encoding="utf-8-sig")
            if len(actual) != item.get("rows") or not _same_rows(actual, expected):
                raise ValueError("CSV differs from the original HTML")
            _store(provider, key, partition, expected,
                   {"migrated_from": file.relative_to(old.parent.parent.parent.parent.parent).as_posix(),
                    "source_url": SMVClient.detail_url(day.to_pydatetime())},
                   raw.read_bytes())
            result["daily_imported"] += 1
        except (OSError, KeyError, ValueError, TypeError):
            result["daily_invalid"] += 1

    if not include_exports:
        return result
    exports = old / "historico_objetivo/exportaciones"
    for file in sorted((exports / "canonical").glob("*.csv")):
        metadata_file = exports / "manifest" / f"{file.stem}.json"
        raw_file = exports / "raw" / f"{file.stem}.xls"
        try:
            metadata = json.loads(metadata_file.read_text(encoding="utf-8"))
            if (sha256(file.read_bytes()).hexdigest() != metadata["csv_sha256"] or
                    sha256(raw_file.read_bytes()).hexdigest() != metadata["raw_sha256"]):
                continue
            frame = pd.read_csv(file)
            dates = pd.to_datetime(frame["value_date"], errors="raise")
            if len(frame) != metadata["rows"] or dates.nunique() != len(frame):
                continue
            for year, chunk in frame.groupby(dates.dt.year):
                key = f"export:{file.stem}:{year}"
                partition = f"export/year={year}/fund={file.stem}"
                prior = provider.storage.manifest.get(key)
                if prior and provider.storage.period_matches(partition, key, prior.get("content_hash")):
                    continue
                chunk = chunk.copy()
                chunk["value_date"] = pd.to_datetime(chunk["value_date"])
                chunk["query_date"] = pd.to_datetime(chunk["query_date"])
                _store(provider, key, partition, chunk,
                       {"migrated_from": file.relative_to(old.parent.parent.parent.parent.parent).as_posix(),
                        "source_url": metadata["source_url"],
                        "raw_sha256": metadata["raw_sha256"]},
                       raw_file.read_bytes(), raw_suffix=".xls")
                result["historical_rows_imported"] += len(chunk)
        except (OSError, ValueError, KeyError):
            continue
    return result
