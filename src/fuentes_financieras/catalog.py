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
    "pe.sbs.participacion": DatasetSpec(
        dataset_id="pe.sbs.participacion",
        provider_class="fuentes_financieras.providers.sbs.participacion:MarketParticipationProvider",
        title="SBS - Tamaño, ranking y participación por entidad",
        country="PE", organization="SBS", frequency="monthly",
        storage_path="peru/sbs/participacion", network_transport="curl_cffi/chrome + Excel XLS/XLSX",
        notes="B/F/C/R; posiciones y porcentajes publicados; banca sin sucursales en el exterior.",
    ),
    "pe.sbs.rentabilidad": DatasetSpec(
        dataset_id="pe.sbs.rentabilidad",
        provider_class="fuentes_financieras.providers.sbs.rentabilidad:ProfitabilityProvider",
        title="SBS - Rentabilidad sobre activos y patrimonio",
        country="PE", organization="SBS", frequency="monthly",
        storage_path="peru/sbs/rentabilidad", network_transport="curl_cffi/chrome + Excel XLS/XLSX",
        notes="B/F/C/R; ratios publicados de utilidad anualizada y denominadores promedio de doce meses.",
    ),
    "pe.sbs.eficiencia": DatasetSpec(
        dataset_id="pe.sbs.eficiencia",
        provider_class="fuentes_financieras.providers.sbs.rentabilidad:EfficiencyProvider",
        title="SBS - Eficiencia y gestión",
        country="PE", organization="SBS", frequency="monthly",
        storage_path="peru/sbs/eficiencia", network_transport="curl_cffi/chrome + Excel XLS/XLSX",
        notes="B/F/C/R; denominadores diferenciados, unidades por persona y oficina, períodos explícitos sin homogeneización implícita.",
    ),

    "pe.sbs.depositos_persona": DatasetSpec(
        dataset_id="pe.sbs.depositos_persona",
        provider_class="fuentes_financieras.providers.sbs.fondeo:DepositsByPersonProvider",
        title="SBS - Depósitos por tipo, persona y entidad",
        country="PE", organization="SBS", frequency="monthly",
        storage_path="peru/sbs/depositos_persona", network_transport="curl_cffi/chrome + Excel XLS/XLSX",
        notes="B/F/C/R; miles de soles; monedas agregadas, sin reconstruir columnas ausentes.",
    ),
    "pe.sbs.depositos_escalas": DatasetSpec(
        dataset_id="pe.sbs.depositos_escalas",
        provider_class="fuentes_financieras.providers.sbs.fondeo:DepositSizeBandsProvider",
        title="SBS - Depósitos según escala de montos del sistema",
        country="PE", organization="SBS", frequency="monthly",
        storage_path="peru/sbs/depositos_escalas", network_transport="curl_cffi/chrome + Excel XLS/XLSX",
        notes="B/F/C/R; agregado del sistema, sin concentración por entidad; números y montos publicados.",
    ),
    "pe.sbs.depositos_plazo": DatasetSpec(
        dataset_id="pe.sbs.depositos_plazo",
        provider_class="fuentes_financieras.providers.sbs.fondeo:DepositsByTermProvider",
        title="SBS - Depósitos del público por moneda y plazo",
        country="PE", organization="SBS", frequency="monthly",
        storage_path="peru/sbs/depositos_plazo", network_transport="curl_cffi/chrome + Excel XLS/XLSX",
        notes="Solo B/F; Reporte 6-B; MN miles de soles, ME miles de dólares; fecha original conservada.",
    ),
    "pe.sbs.adeudos": DatasetSpec(
        dataset_id="pe.sbs.adeudos",
        provider_class="fuentes_financieras.providers.sbs.fondeo:FinancialObligationsProvider",
        title="SBS - Estructura de adeudos y obligaciones financieras",
        country="PE", organization="SBS", frequency="monthly",
        storage_path="peru/sbs/adeudos", network_transport="curl_cffi/chrome + Excel XLS/XLSX",
        notes="B/F/C/R; participaciones en porcentaje, total en miles de soles.",
    ),
    "pe.sbs.castigos": DatasetSpec(
        dataset_id="pe.sbs.castigos",
        provider_class="fuentes_financieras.providers.sbs.castigos:CreditWriteoffsProvider",
        title="SBS - Flujo mensual de créditos castigados",
        country="PE", organization="SBS", frequency="monthly",
        storage_path="peru/sbs/castigos", network_transport="curl_cffi/chrome + Excel XLS/XLSX",
        notes="B/F/C/R; miles de soles; flujo mensual, sin reconstruir acumulados ni homogeneizar clasificaciones históricas.",
    ),

    "pe.sbs.calidad_cartera": DatasetSpec(
        dataset_id="pe.sbs.calidad_cartera",
        provider_class="fuentes_financieras.providers.sbs.calidad_cartera:CreditQualityProvider",
        title="SBS - Calidad de activos y cobertura de provisiones",
        country="PE", organization="SBS", frequency="monthly",
        storage_path="peru/sbs/calidad_cartera", network_transport="curl_cffi/chrome + Excel XLS/XLSX",
        notes="B/F/C/R; Solo indicadores de calidad; unidades porcentuales y definiciones originales.",
    ),
    "pe.sbs.categorias_riesgo_cartera": DatasetSpec(
        dataset_id="pe.sbs.categorias_riesgo_cartera",
        provider_class="fuentes_financieras.providers.sbs.calidad_cartera:CreditRiskCategoriesProvider",
        title="SBS - Categorías de riesgo del deudor",
        country="PE", organization="SBS", frequency="monthly",
        storage_path="peru/sbs/categorias_riesgo_cartera", network_transport="curl_cffi/chrome + Excel XLS/XLSX",
        notes="B/F/C/R; Participaciones y exposición total; créditos indirectos/contingentes según cuadro.",
    ),
    "pe.sbs.morosidad_dias": DatasetSpec(
        dataset_id="pe.sbs.morosidad_dias",
        provider_class="fuentes_financieras.providers.sbs.calidad_cartera:CreditArrearsProvider",
        title="SBS - Morosidad por días de incumplimiento",
        country="PE", organization="SBS", frequency="monthly",
        storage_path="peru/sbs/morosidad_dias", network_transport="curl_cffi/chrome + Excel XLS/XLSX",
        notes="B/F/C/R; Umbrales mayores de 30/60/90/120 días y criterio contable SBS separados.",
    ),
    "pe.sbs.saldos_cartera": DatasetSpec(
        dataset_id="pe.sbs.saldos_cartera",
        provider_class="fuentes_financieras.providers.sbs.saldos_cartera:CreditBalancesProvider",
        title="SBS - Saldos de créditos y provisiones",
        country="PE", organization="SBS", frequency="monthly",
        storage_path="peru/sbs/saldos_cartera", network_transport="curl_cffi/chrome + Excel XLS/XLSX",
        notes="B/F/C/R; Bloque crediticio del balance; miles de soles, signos originales y sin ratios calculados.",
    ),

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
