from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class DatasetSpec:
    dataset_id: str
    provider_class: str
    title: str
    country: str | None
    organization: str
    frequency: str
    storage_path: str
    network_transport: str
    notes: str = ""


CATALOG: dict[str, DatasetSpec] = {
    "pe.sbs.tasas_pasivas": DatasetSpec(
        dataset_id="pe.sbs.tasas_pasivas",
        provider_class="fuentes_financieras.providers.sbs.tasas_pasivas:PassiveRatesProvider",
        title="SBS - Tasas Pasivas por Empresa",
        country="PE",
        organization="SBS",
        frequency="daily+monthly",
        storage_path="peru/sbs/tasas_pasivas",
        network_transport="curl_cffi/chrome + ASP.NET WebForms",
        notes="B/F diarios; C/R mensuales.",
    ),
    "pe.sbs.curva_soberana": DatasetSpec(
        dataset_id="pe.sbs.curva_soberana",
        provider_class="fuentes_financieras.providers.sbs.curva_soberana:SovereignCurveProvider",
        title="SBS - Curva Soberana",
        country="PE",
        organization="SBS",
        frequency="daily",
        storage_path="peru/sbs/curva_soberana",
        network_transport="HTTP POST JSON",
        notes="Adapter para mover la implementación ya validada.",
    ),
    "pe.smv.fondos_mutuos.valores_cuota": DatasetSpec(
        dataset_id="pe.smv.fondos_mutuos.valores_cuota",
        provider_class="fuentes_financieras.providers.smv.fondos_mutuos:MutualFundValuesProvider",
        title="SMV - Valores Cuota de Fondos Mutuos",
        country="PE",
        organization="SMV",
        frequency="daily",
        storage_path="peru/smv/fondos_mutuos/valores_cuota",
        network_transport="provider-specific",
        notes="Todas las filas publicadas por fecha de consulta; el proyecto filtra después.",
    ),
}
