"""
Profile default si reguli auto-ajustare pentru aplicatia interactiva.
"""

# Defaults pentru parcare ~250 locuri RETAIL (asteptat verdict MARGINAL)
RETAIL_250 = {
    "tip_parcare": "RETAIL",
    "tip_investitie": "mix",
    "capex_eur": 40000,
    "pondere_credit": 0.70,
    "durata_credit_ani": 5,
    "rata_dobanda": 0.09,
    "durata_contract_ani": 5,
    "overhead_anual_eur": 12000,
    "numar_locuri": 250,
    "trafic_zilnic": 1250,
    "opex_lunar_eur": 140,
    "gratuitate_min": 60,
    "tarif_ora_1_ron": 5.0,
    "tarif_ora_2_ron": 8.0,
    "tarif_ora_3plus_ron": 12.0,
    "tarif_sesiune_ron": 10.0,
    "rata_colectare": 0.75,
    "cota_impozit": 0.16,
    "inflatie_opex": 0.06,
    "inflatie_tarife": 0.06,
    "discount_rate": 0.12,
    "eur_ron": 5.0,
    "provizion_pct": 0.06,
    "amortizare_ani": 5,
    "rampup_an1": 0.80,
    "asigurare_anuala_eur": 400,
    "marketing_initial_eur": 1000,
    "marketing_lunar_eur": 100,
}

# Defaults pentru parcare ~250 locuri NON-RETAIL (asteptat verdict VIABIL)
NONRETAIL_250 = {
    **RETAIL_250,
    "tip_parcare": "NON-RETAIL",
    "capex_eur": 32000,
    "trafic_zilnic": 625,
    "opex_lunar_eur": 132,
    "gratuitate_min": 0,
    "tarif_sesiune_ron": 10.0,
    "rata_colectare": 0.75,
}


def auto_capex_from_locuri(numar_locuri: int, tip_parcare: str) -> float:
    """Sugereaza CAPEX pe baza dimensiunii parcarii."""
    if tip_parcare == "RETAIL":
        return 15000 + 100 * numar_locuri
    return 12000 + 80 * numar_locuri


def auto_trafic_from_locuri(numar_locuri: int, tip_parcare: str) -> float:
    """Sugereaza trafic baseline (rotatie tipica/zi)."""
    if tip_parcare == "RETAIL":
        return numar_locuri * 5.0
    return numar_locuri * 2.5


def auto_opex_from_capex(capex_eur: float) -> float:
    """Sugereaza OPEX lunar (proportional cu CAPEX, baseline + uzura)."""
    return 100 + capex_eur * 0.012 / 12


# Verdict thresholds (configurabile)
VERDICT_THRESHOLDS = {
    "irr_profitabil": 0.15,
    "payback_profitabil": 4.0,
    "irr_marginal": 0.10,
}
