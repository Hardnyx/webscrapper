"""Complete SMV mutual-fund snapshots, independent of project fund selectors."""
from __future__ import annotations

from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd

from fuentes_financieras.exceptions import InvalidQueryError, SourceUnavailableError
from fuentes_financieras.models import FetchResult, PeriodRequest
from fuentes_financieras.provider import DatasetProvider

from .canonical import canonicalize_mutual_fund_values
from .client import SMVClient
from .parser import parse_smv_html


FIRST_DETAIL_DATE = pd.Timestamp("2025-03-08")


def _date(value: object) -> pd.Timestamp:
    return pd.Timestamp(value).normalize()


def keys(day: pd.Timestamp) -> tuple[str, str]:
    return f"snapshot:{day:%Y-%m-%d}", f"year={day:%Y}/month={day:%m}/day={day:%d}"


def normalize_snapshot(html: str, day: pd.Timestamp, transport: str) -> pd.DataFrame:
    parsed = parse_smv_html(html)
    if parsed.empty or parsed["Fondo Mutuo"].isna().any():
        raise SourceUnavailableError(f"SMV: respuesta sin fondos completos para {day:%d/%m/%Y}")
    parsed["Fecha Consulta"] = day
    parsed["Fuente"] = "SMV"
    parsed["Método Consulta"] = transport
    out = canonicalize_mutual_fund_values(parsed)
    if out["fund_name"].isna().any() or out["unit_value"].isna().any():
        raise SourceUnavailableError(f"SMV: cuotas o nombres ilegibles para {day:%d/%m/%Y}")
    return out.reset_index(drop=True)


class MutualFundValuesProvider(DatasetProvider):
    """Fetch and cache every fund/series published on each query date."""

    parser_version = "smv-complete-2"
    contract_version = "2"

    def single_request(self, *, fecha: str | pd.Timestamp, **query) -> PeriodRequest:
        if query:
            raise InvalidQueryError(f"Parámetros desconocidos: {sorted(query)}")
        day = _date(fecha)
        period, partition = keys(day)
        return PeriodRequest(period, partition, {"fecha": day.strftime("%Y-%m-%d")})

    def plan_sync(
        self, *, desde: str | pd.Timestamp = FIRST_DETAIL_DATE,
        hasta: str | pd.Timestamp | None = None, refresh_recent: bool = False,
        **query,
    ):
        if query:
            raise InvalidQueryError(
                f"SMV descarga todos los fondos; filtro posterior requerido en la capa consumidora: {sorted(query)}"
            )
        end = (_date(hasta) if hasta is not None else
               _date(datetime.now(ZoneInfo("America/Lima"))) - timedelta(days=1))
        start = _date(desde)
        if end < start:
            raise InvalidQueryError("Rango inválido: hasta anterior a desde.")
        for day in pd.date_range(start, end):
            period, partition = keys(day)
            yield PeriodRequest(
                period, partition, {"fecha": day.strftime("%Y-%m-%d")},
                mutable=bool(refresh_recent and day >= end - timedelta(days=7)),
            )

    def _fetch_period(self, request: PeriodRequest) -> FetchResult:
        day = _date(request.params["fecha"])
        with SMVClient(transport="auto", timeout=30, max_retries=2) as client:
            response = client.fetch(day.to_pydatetime())
        data = normalize_snapshot(response.html, day, response.transport)
        return FetchResult(
            dataset_id=self.spec.dataset_id, data=data, raw=response.html,
            metadata={"query_date": day.strftime("%Y-%m-%d"), "rows": len(data),
                      "transport": response.transport,
                      "source_url": SMVClient.detail_url(day.to_pydatetime())},
        )

    def sync(self, *, keep_raw: bool = True, **query):
        # The generic result preserves successful dates and exposes failed dates
        # without converting a partial update into an exception.
        return super().sync(keep_raw=keep_raw, **query)

    def partition_keys_for_load(
        self,
        *,
        fechas_consulta=None,
        desde=None,
        hasta=None,
        **query,
    ):
        if query:
            raise InvalidQueryError(
                f"Parámetros de carga no reconocidos: {sorted(query)}"
            )

        if fechas_consulta is not None:
            if isinstance(fechas_consulta, (str, pd.Timestamp, datetime)):
                targets = [_date(fechas_consulta)]
            else:
                targets = [_date(value) for value in fechas_consulta]

            # One nearest cached snapshot is sufficient for each as-of target.
            # A short backward search covers holidays, weekends and publication gaps.
            partitions = []
            seen = set()
            for target in sorted(set(targets)):
                selected = None
                for offset in range(15):
                    candidate = target - timedelta(days=offset)
                    partition = keys(candidate)[1]
                    if self.storage.partition_path(partition).is_file():
                        selected = partition
                        break
                if selected is not None and selected not in seen:
                    seen.add(selected)
                    partitions.append(selected)
            return partitions

        # Full-history semantics remain unchanged when no lower bound is supplied.
        if desde is None:
            return None

        start = _date(desde)
        end = (
            _date(hasta)
            if hasta is not None
            else _date(datetime.now(ZoneInfo("America/Lima")))
        )
        if end < start:
            raise InvalidQueryError("Rango inválido: hasta anterior a desde.")

        return [keys(day)[1] for day in pd.date_range(start, end)]

    def filter_loaded(
        self, data: pd.DataFrame, *, fechas_consulta=None, desde=None, hasta=None, **query
    ) -> pd.DataFrame:
        if query:
            raise InvalidQueryError(f"Filtro de fondos requerido en el proyecto consumidor: {sorted(query)}")
        out = data.copy()
        out["value_date"] = pd.to_datetime(out["value_date"], errors="coerce")
        out["query_date"] = pd.to_datetime(out["query_date"], errors="coerce")
        if desde is not None:
            out = out[out["value_date"] >= _date(desde)]
        if hasta is not None:
            out = out[out["value_date"] <= _date(hasta)]
        # No colapsar aquí nombres repetidos de administradoras distintas ni
        # filas publicadas nuevamente con la misma fecha efectiva. El proyecto
        # decide qué serie/observación usar para su análisis.
        return out.sort_values(["query_date", "fund_name", "value_date"], kind="mergesort")

    def migrate_legacy(self, automations_root, *, include_exports: bool = True):
        from .migration import migrate_legacy
        return migrate_legacy(self, automations_root, include_exports=include_exports)

    def discover_funds(self, *, refresh: bool = False):
        from .historical import discover_funds
        return discover_funds(self, refresh=refresh)

    def sync_historical(self, *, desde, hasta, refresh_catalog: bool = False,
                        force: bool = False, funds=None):
        from .historical import sync_historical
        return sync_historical(self, desde=desde, hasta=hasta,
                               refresh_catalog=refresh_catalog, force=force, funds=funds)
