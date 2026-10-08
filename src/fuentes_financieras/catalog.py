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
    "pe.sbs.liquidez": DatasetSpec(
        dataset_id="pe.sbs.liquidez",
        provider_class="fuentes_financieras.providers.sbs.liquidez:LiquidityProvider",
        title="SBS - Liquidez en moneda nacional y extranjera",
        country="PE", organization="SBS", frequency="monthly",
        storage_path="peru/sbs/liquidez", network_transport="curl_cffi/chrome + Excel XLS/XLSX",
        notes="B/F/C/R; importes MN en miles de soles y ME en miles de dólares.",
    ),
    "pe.sbs.cobertura_liquidez": DatasetSpec(
        dataset_id="pe.sbs.cobertura_liquidez",
        provider_class="fuentes_financieras.providers.sbs.liquidez:LiquidityCoverageProvider",
        title="SBS - Cobertura de liquidez",
        country="PE", organization="SBS", frequency="quarterly",
        storage_path="peru/sbs/cobertura_liquidez", network_transport="curl_cffi/chrome + Excel XLS/XLSX",
        notes="B/F/C/R; promedio diario trimestral; mes del índice separado del trimestre observado.",
    ),
    "pe.sbs.financiacion_neta_estable": DatasetSpec(
        dataset_id="pe.sbs.financiacion_neta_estable",
        provider_class="fuentes_financieras.providers.sbs.liquidez:StableFundingProvider",
        title="SBS - Financiación neta estable",
        country="PE", organization="SBS", frequency="monthly",
        storage_path="peru/sbs/financiacion_neta_estable", network_transport="curl_cffi/chrome + Excel XLS/XLSX",
        notes="B/F/C/R; importes ponderados en miles de soles; formato Excel porcentual verificado.",
    ),

    "pe.sbs.solvencia": DatasetSpec(
        dataset_id="pe.sbs.solvencia",
        provider_class="fuentes_financieras.providers.sbs.solvencia:SolvencyProvider",
        title="SBS - Requerimientos patrimoniales, APR y ratios de capital",
        country="PE", organization="SBS", frequency="monthly",
        storage_path="peru/sbs/solvencia", network_transport="curl_cffi/chrome + Excel XLS/XLSX",
        notes="B/F/C/R; requerimientos y APR en miles de soles; ratios en porcentaje.",
    ),
    "pe.sbs.patrimonio_efectivo": DatasetSpec(
        dataset_id="pe.sbs.patrimonio_efectivo",
        provider_class="fuentes_financieras.providers.sbs.solvencia:EffectiveCapitalProvider",
        title="SBS - Patrimonio efectivo y su composición",
        country="PE", organization="SBS", frequency="monthly",
        storage_path="peru/sbs/patrimonio_efectivo", network_transport="curl_cffi/chrome + Excel XLS/XLSX",
        notes="B/F/C/R; unidades publicadas por cuadro; preserva ambigüedades e inconsistencias.",
    ),
    "pe.sbs.estados_financieros": DatasetSpec(
        dataset_id="pe.sbs.estados_financieros",
        provider_class="fuentes_financieras.providers.sbs.estados_financieros:FinancialStatementsProvider",
        title="SBS - Balance y resultados estadísticos por entidad",
        country="PE", organization="SBS", frequency="monthly",
        storage_path="peru/sbs/estados_financieros",
        network_transport="curl_cffi/chrome + Excel XLS/XLSX",
        notes="B/F/C/R; miles de soles; resultados acumulados del ejercicio.",
    ),
    "pe.sbs.tasas_pasivas_mercado": DatasetSpec(
        dataset_id="pe.sbs.tasas_pasivas_mercado",
        provider_class="fuentes_financieras.providers.sbs.tasas_pasivas_mercado:PassiveMarketProvider",
        title="SBS - Tasas pasivas de mercado sobre saldos y flujos",
        country="PE", organization="SBS", frequency="daily",
        storage_path="peru/sbs/tasas_pasivas_mercado",
        network_transport="curl_cffi/chrome + ASP.NET WebForms",
        notes="TIPMN/TIPMEX sobre saldos B+F; FTIPMN/FTIPMEX sobre flujos B de 30 días útiles.",
    ),
    "pe.sbs.universo_depositos": DatasetSpec(
        dataset_id="pe.sbs.universo_depositos",
        provider_class="fuentes_financieras.providers.sbs.universo_depositos:DepositUniverseProvider",
        title="SBS - Universo de entidades autorizadas a captar depósitos",
        country="PE", organization="SBS", frequency="snapshot",
        storage_path="peru/sbs/universo_depositos",
        network_transport="curl_cffi/chrome + HTML",
        notes="Observaciones del universo vigente; no reconstruye autorizaciones históricas.",
    ),
    "pe.sbs.clasificaciones_riesgo": DatasetSpec(
        dataset_id="pe.sbs.clasificaciones_riesgo",
        provider_class="fuentes_financieras.providers.sbs.clasificaciones_riesgo:RiskRatingsProvider",
        title="SBS - Clasificaciones e Informes Semestrales",
        country="PE",
        organization="SBS",
        frequency="semiannual",
        storage_path="peru/sbs/clasificaciones_riesgo",
        network_transport="curl_cffi/chrome + ASP.NET WebForms",
        notes="Clasificaciones semestrales; una consulta por período recupera todos los tipos de entidad.",
    ),
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
