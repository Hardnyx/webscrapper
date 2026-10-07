#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Download and validate the complete SBS historical ratings dataset.

Usage:
    python -m fuentes_financieras.cli.clasificaciones_riesgo

Options:
    --data-root PATH
        Storage root for manifest and Parquet files.

    --force
        Re-query every published period.

    --keep-raw
        Keep compressed raw HTML/delta responses for auditing.

    --no-second-sync
        Skip the second cache-validation sync.

Normal execution:
1. uses dependencies declared by the installed package;
2. discovers every period currently published by SBS;
3. downloads every missing historical period;
4. repeats sync and validates downloaded=0;
5. loads the complete local history;
6. writes CSV and JSON validation summaries.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import traceback
from datetime import datetime
from pathlib import Path


def section(title: str):
    print("\n" + "=" * 112)
    print(title)
    print("=" * 112)


def compact_sync(result):
    return {
        "requested": result.requested,
        "downloaded": result.downloaded,
        "unchanged": result.unchanged,
        "skipped_existing": result.skipped_existing,
        "unavailable": result.unavailable,
        "failed": result.failed,
        "canonical_root": str(result.canonical_root) if result.canonical_root else None,
        "details": result.details,
    }


def run(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument("--data-root", default=None)
    parser.add_argument("--force", action="store_true")
    parser.add_argument("--keep-raw", action="store_true")
    parser.add_argument("--no-second-sync", action="store_true")
    parser.add_argument("--output-dir", type=Path, default=Path.cwd() / "outputs" / "clasificaciones_riesgo")
    args = parser.parse_args(argv)

    from fuentes_financieras.runtime import resolve_data_root
    data_root = resolve_data_root(args.data_root)
    data_root.mkdir(parents=True, exist_ok=True)
    os.environ["FINANCIAL_SOURCES_DATA_ROOT"] = str(data_root)
    output_dir = args.output_dir.expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    from fuentes_financieras.registry import get_provider
    get_provider.cache_clear()

    from fuentes_financieras import source

    ratings = source("pe.sbs.clasificaciones_riesgo")

    # v2 migrates safely from a partial v1 run. Existing v1 periods are
    # re-fetched under contract v2 and upserted in place; the user does not
    # need to delete datos_historico manually.
    manifest_contract = (
        ratings.storage.manifest.last_schema_contract_version
    )
    if manifest_contract not in (None, str(ratings.contract_version)):
        print(
            "\n[migración] Se detectó un esquema de contrato anterior "
            f"({manifest_contract}). Se migrará automáticamente a "
            f"contract_version={ratings.contract_version}."
        )
    elif (
        ratings.storage.manifest.last_schema_hash
        and manifest_contract is None
    ):
        print(
            "\n[migración] Se detectó un manifest v1 parcial sin versión "
            "de contrato de esquema. Se migrará automáticamente."
        )

    report = {
        "started_at": datetime.now().isoformat(timespec="seconds"),
        "python": sys.version,
        "executable": sys.executable,
        "data_root": str(data_root),
        "dataset_id": "pe.sbs.clasificaciones_riesgo",
        "tests": {},
        "overall_ok": False,
    }

    section("1. CATÁLOGO DE PERÍODOS SBS")
    periods = ratings.available_periods()
    if not periods:
        raise RuntimeError("SBS no devolvió períodos disponibles.")

    for item in periods:
        print(
            f"{item['period_code']} | {item['label']:<22} | "
            f"fecha={item['period_date']} | selected={item['selected']}"
        )

    print("\nPeríodos disponibles:", len(periods))
    print("Más reciente        :", periods[0]["period_code"], periods[0]["label"])
    print("Más antiguo         :", periods[-1]["period_code"], periods[-1]["label"])

    report["period_catalog"] = periods

    section("2. PRIMER SYNC — TODO EL HISTÓRICO")
    first = ratings.sync(
        force=args.force,
        refresh_hours=24,
        keep_raw=args.keep_raw,
    )

    print(
        f"requested={first.requested} | downloaded={first.downloaded} | "
        f"unchanged={first.unchanged} | skipped={first.skipped_existing} | "
        f"unavailable={first.unavailable} | failed={first.failed}"
    )

    for detail in first.details:
        print(
            f"{detail.get('period_key', ''):<8} | "
            f"{detail.get('status', ''):<20} | "
            f"rows={detail.get('rows', '')}"
        )

    report["tests"]["first_sync"] = {
        "ok": first.failed == 0 and first.unavailable == 0,
        **compact_sync(first),
    }

    if first.failed or first.unavailable:
        raise RuntimeError(
            f"Primer sync incompleto: failed={first.failed}, unavailable={first.unavailable}"
        )

    if not args.no_second_sync:
        section("3. SEGUNDO SYNC IDÉNTICO — VALIDACIÓN DE CACHÉ")
        second = ratings.sync(
            refresh_hours=24,
            keep_raw=args.keep_raw,
        )

        print(
            f"requested={second.requested} | downloaded={second.downloaded} | "
            f"unchanged={second.unchanged} | skipped={second.skipped_existing} | "
            f"unavailable={second.unavailable} | failed={second.failed}"
        )

        cache_ok = (
            second.downloaded == 0
            and second.failed == 0
            and second.unavailable == 0
            and second.skipped_existing == second.requested
        )

        print("\n¿SEGUNDO SYNC SIN REDESCARGAR PERÍODOS? ->", cache_ok)

        report["tests"]["second_sync_cache"] = {
            "ok": cache_ok,
            **compact_sync(second),
        }

        if not cache_ok:
            raise RuntimeError("La segunda sincronización no quedó completamente en caché.")

    section("4. LOAD() — HISTÓRICO LOCAL COMPLETO")
    df = ratings.load()

    if df.empty:
        raise RuntimeError("load() devolvió un DataFrame vacío.")

    print("Filas             :", len(df))
    print("Períodos          :", df["period_code"].nunique())
    print("Entidades únicas  :", df[["entity_type", "entity_name"]].drop_duplicates().shape[0])
    print("Clasificadoras    :", df["rating_agency"].nunique())
    print("Desde             :", df["period_date"].min())
    print("Hasta             :", df["period_date"].max())

    actual_periods = sorted(df["period_code"].astype(str).unique().tolist())
    expected_periods = sorted(x["period_code"] for x in periods)
    period_coverage_ok = actual_periods == expected_periods

    duplicates = int(
        df.duplicated(
            subset=[
                "period_code",
                "entity_type",
                "entity_name",
                "rating_agency",
            ]
        ).sum()
    )

    null_type_codes = int(df["entity_type_code"].isna().sum())

    print("Cobertura 100%     :", period_coverage_ok)
    print("Duplicados lógicos :", duplicates)
    print("Tipos sin código   :", null_type_codes)

    unmapped_labels = sorted(
        df.loc[df["entity_type_code"].isna(), "entity_type"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )
    if unmapped_labels:
        print("Etiquetas históricas sin código actual SBS:")
        for label in unmapped_labels:
            print("  -", label)

    summary = (
        df.groupby(
            ["period_code", "period", "period_date"],
            dropna=False,
        )
        .agg(
            ratings=("rating", "size"),
            entities=("entity_name", "nunique"),
            entity_types=("entity_type", "nunique"),
            agencies=("rating_agency", "nunique"),
        )
        .reset_index()
        .sort_values("period_code")
    )

    summary_path = output_dir / "resumen_historico_clasificaciones.csv"
    summary.to_csv(summary_path, index=False, encoding="utf-8-sig")

    report["tests"]["load_history"] = {
        "ok": period_coverage_ok and duplicates == 0,
        "rows": int(len(df)),
        "periods": int(df["period_code"].nunique()),
        "period_coverage_ok": period_coverage_ok,
        "logical_duplicates": duplicates,
        "null_entity_type_codes": null_type_codes,
        "unmapped_entity_type_labels": unmapped_labels,
        "summary_csv": str(summary_path),
    }

    section("5. MUESTRA")
    cols = [
        "period_code",
        "entity_type_code",
        "entity_type",
        "entity_name",
        "rating_agency",
        "rating",
        "trend",
    ]
    print(df[cols].head(30).to_string(index=False))

    report["overall_ok"] = all(
        item.get("ok") is True
        for item in report["tests"].values()
    )
    report["finished_at"] = datetime.now().isoformat(timespec="seconds")

    report_path = output_dir / "resultado_historico_clasificaciones.json"
    report_path.write_text(
        json.dumps(report, indent=2, ensure_ascii=False, default=str),
        encoding="utf-8",
    )

    section("RESULTADO FINAL")
    for name, result in report["tests"].items():
        print(f"{name:<30} {'OK' if result.get('ok') else 'FALLÓ'}")

    print("\nOVERALL:", report["overall_ok"])
    print("Datos   :", data_root)
    print("Reporte :", report_path)
    print("Resumen :", summary_path)

    if report["overall_ok"]:
        print(
            "\nCONFIRMADO: todo el histórico SBS de clasificaciones quedó "
            "sincronizado y disponible localmente."
        )
        return 0

    return 1


def main(argv=None) -> int:
    try:
        return run(argv)
    except KeyboardInterrupt:
        return 130
    except Exception:
        section("EJECUCIÓN FALLIDA")
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
