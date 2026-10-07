from __future__ import annotations

from fuentes_financieras.exceptions import SourceUnavailableError
from fuentes_financieras.provider import DatasetProvider


class SovereignCurveProvider(DatasetProvider):
    """
    Punto de integración.

    El diagnóstico previo ya demostró que esta fuente usa un POST JSON simple.
    No se copia aquí todavía para mantener este artefacto centrado en la arquitectura.
    """

    parser_version = "pending-migration"
    contract_version = "1"

    def single_request(self, **query):
        raise SourceUnavailableError(
            "Adapter de Curva Soberana pendiente de migrar desde la implementación validada."
        )

    def plan_sync(self, **query):
        raise SourceUnavailableError(
            "Adapter de Curva Soberana pendiente de migrar."
        )

    def _fetch_period(self, request):
        raise SourceUnavailableError(
            "Adapter de Curva Soberana pendiente de migrar."
        )
