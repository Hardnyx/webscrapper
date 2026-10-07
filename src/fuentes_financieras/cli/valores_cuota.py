"""Refresh the shared SMV dataset independently of any report notebook."""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from fuentes_financieras import source


def main() -> None:
    parser = argparse.ArgumentParser(description="Sincroniza todos los fondos y series publicados por SMV.")
    parser.add_argument("--migrate-from", type=Path, help="Raíz anterior de Automatizaciones; lee solo capturas verificadas.")
    parser.add_argument("--desde", default="2025-03-08")
    parser.add_argument("--hasta", default=None, help="Por defecto, ayer en Lima.")
    parser.add_argument("--refresh-recent", action="store_true",
                        help="Revalidar la última semana aunque esté guardada.")
    parser.add_argument("--historical-from", help="Inicio del backfill histórico EVCP para todos los fondos.")
    parser.add_argument("--historical-through", help="Fin del backfill histórico EVCP.")
    args = parser.parse_args()
    dataset = source("pe.smv.fondos_mutuos.valores_cuota")
    if args.migrate_from:
        print("Migración:", json.dumps(dataset.migrate_legacy(args.migrate_from), ensure_ascii=False))
    if args.historical_from or args.historical_through:
        if not args.historical_from or not args.historical_through:
            parser.error("--historical-from y --historical-through deben indicarse juntos")
        print("Histórico:", json.dumps(dataset.sync_historical(
            desde=args.historical_from, hasta=args.historical_through), ensure_ascii=False))
    result = dataset.sync(desde=args.desde, hasta=args.hasta, refresh_recent=args.refresh_recent)
    print("Sincronización:", json.dumps({
        "fechas": result.requested, "nuevas": result.downloaded,
        "reutilizadas": result.skipped_existing, "sin_cambio": result.unchanged,
        "pendientes": result.failed + result.unavailable,
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
