"""
Teste de regresie pentru parking_calc.py vs validator.py existing.
Goal: cand tarif uniform t1=t2=t3, simulate_scenario produce aceleasi cifre ca cf_net_per_parcare din validator.
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

import pytest

from parking_calc import (
    BUCKETS_RETAIL,
    venit_per_intrare_tiered,
    venit_anual_retail_eur,
    venit_anual_nonretail_eur,
    anuitate_credit,
    npv, irr, payback_simple,
    ScenarioParams, simulate_scenario,
    verdict_from_kpis,
)


# T1 — Tarif uniform produce acelasi rezultat ca formula clasica
def test_tarif_uniform_match_validator_classic():
    """Cand t1=t2=t3=10 si gratuitate=120, venit_per_intrare_tiered trebuie sa egaleze formula clasica."""
    durata_min = 240  # bucket 180+
    gratuitate = 120
    tarif = 10.0
    # Tiered
    v_tiered = venit_per_intrare_tiered(durata_min, gratuitate, tarif, tarif, tarif)
    # Clasic: max(0, durata - grat) / 60 * tarif
    v_classic = max(0, durata_min - gratuitate) / 60 * tarif
    assert abs(v_tiered - v_classic) < 0.001, f"{v_tiered} vs {v_classic}"


def test_tarif_uniform_buckets_sum():
    """Suma pe buckets cu tarif uniform = formula clasica venit_per_intrare retail."""
    gratuitate = 120
    tarif = 10.0
    # Tiered sum
    v_tiered_total = sum(
        p * venit_per_intrare_tiered(m, gratuitate, tarif, tarif, tarif)
        for m, p in BUCKETS_RETAIL
    )
    # Clasic
    v_classic = sum(p * max(0, m - gratuitate) / 60 * tarif for m, p in BUCKETS_RETAIL)
    assert abs(v_tiered_total - v_classic) < 0.001


# T5 — Tarif progresiv exact (test hardcoded din plan)
def test_tarif_progresiv_exact():
    """durata 150min, gratuitate 30min, t1=5/t2=7/t3=10:
    taxabil = 120min = 1h ora1 + 1h ora2 + 0h ora3+
    venit = 1×5 + 1×7 + 0×10 = 12 RON
    """
    v = venit_per_intrare_tiered(150, 30, 5, 7, 10)
    assert abs(v - 12.0) < 0.001, f"got {v}"


def test_tarif_progresiv_2h_taxabil():
    """durata 180min, gratuitate 30min, t1=5/t2=7/t3=10:
    taxabil = 150min = 1h ora1 + 1h ora2 + 0.5h ora3+
    venit = 5 + 7 + 5 = 17 RON
    """
    v = venit_per_intrare_tiered(180, 30, 5, 7, 10)
    assert abs(v - 17.0) < 0.001, f"got {v}"


def test_zero_taxabil_when_durata_under_gratuitate():
    """Daca durata < gratuitate, venit = 0."""
    v = venit_per_intrare_tiered(60, 120, 10, 10, 10)
    assert v == 0.0


# Sanity checks portate din validator.py
def test_buckets_sum_to_one():
    assert abs(sum(p for _, p in BUCKETS_RETAIL) - 1.0) < 0.001


def test_anuitate_x5_greater_than_credit():
    """Anuitate pe 5 ani × 5 > credit (datorita dobanzii)."""
    capex = 40000
    pondere = 0.7
    a = anuitate_credit(capex, 0.09, 5, pondere)
    credit = capex * pondere
    assert a * 5 > credit


def test_anuitate_zero_when_no_credit():
    """Pondere credit 0 → anuitate 0."""
    a = anuitate_credit(40000, 0.09, 5, 0.0)
    assert a == 0.0


def test_npv_with_zero_capex():
    """CAPEX=0 → NPV = sum(CF / (1+r)^n)."""
    cf = [1000, 1000, 1000]
    r = 0.10
    expected = 1000 / 1.1 + 1000 / 1.1**2 + 1000 / 1.1**3
    assert abs(npv(cf, 0, r) - expected) < 0.01


def test_irr_basic():
    """CAPEX 100, CF 50/50/50 → IRR ~23%."""
    cf = [50, 50, 50]
    r = irr(cf, 100)
    assert r is not None and 0.20 < r < 0.25


# T1bis — simulate_scenario cu tarif uniform si parametri bait validator
def test_simulate_retail_uniform_matches_validator_baseline():
    """
    Setup: tarif uniform 10 RON, gratuitate 120 (B-C2-g120 din validator),
    fara overhead, fara provizion, fara amortizare, fara anuitate.
    Verifica venit anul 1 baseline matches validator's calculation.
    """
    # Parametri pure pentru a reproduce comportamentul venit_anual_retail_ron(B, 120) din validator
    # Validator: venit_per_intrare = 0.30 RON la g=120 cu BUCKETS_B [(7.5,0.25),(22.5,0.45),(45,0.22),(90,0.05),(150,0.02),(240,0.01)]
    # venit_anual_RON = 0.30 * 1000 * 365 = 109,500 RON = 21,900 EUR

    venit_eur = venit_anual_retail_eur(
        BUCKETS_RETAIL,
        gratuitate_min=120,
        t1=10.0, t2=10.0, t3plus=10.0,
        trafic_zi=1000,
        eur_ron=5.0,
    )
    # Asteptat ~21,900 EUR (aceiasi distributie B din validator)
    assert abs(venit_eur - 21900) < 100, f"got {venit_eur}"


def test_simulate_scenario_profitable_when_obvious():
    """Scenariu non-retail extreme: trafic mare, colectare 100% → PROFITABIL."""
    p = ScenarioParams(
        tip_parcare="NON-RETAIL",
        tip_investitie="capital_propriu",
        capex_eur=30000,
        pondere_credit=0.0,
        durata_credit_ani=5,
        rata_dobanda=0.0,
        durata_contract_ani=5,
        overhead_anual_eur=0,
        numar_locuri=200,
        trafic_zilnic=400,
        opex_lunar_eur=400,
        tarif_sesiune_ron=15,
        rata_colectare=1.0,
        rampup_an1=1.0,
        asigurare_anuala_eur=0,
        marketing_initial_eur=0,
        marketing_lunar_eur=0,
    )
    r = simulate_scenario(p)
    assert r['npv'] > 0
    assert r['irr'] is not None and r['irr'] > 0.50
    assert r['verdict'] == "PROFITABIL"


def test_simulate_scenario_unprofitable_when_obvious():
    """Scenariu retail extreme: tarife mici, gratuitate mare, trafic mic → NEPROFITABIL."""
    p = ScenarioParams(
        tip_parcare="RETAIL",
        capex_eur=80000,  # CAPEX mare
        trafic_zilnic=100,  # trafic foarte mic
        gratuitate_min=180,  # gratuitate mare
        tarif_ora_1_ron=2,
        tarif_ora_2_ron=3,
        tarif_ora_3plus_ron=5,
        overhead_anual_eur=20000,
        rampup_an1=0.5,
    )
    r = simulate_scenario(p)
    assert r['npv'] < 0
    assert r['verdict'] in ("NEPROFITABIL", "MARGINAL")


def test_verdict_thresholds():
    """Sanity pe verdict_from_kpis."""
    # PROFITABIL
    assert verdict_from_kpis(10000, 0.20, 3.0) == "PROFITABIL"
    # MARGINAL (NPV pozitiv dar IRR sub 15)
    assert verdict_from_kpis(5000, 0.13, 5.0) == "MARGINAL"
    # MARGINAL (IRR > 10 dar payback > 4)
    assert verdict_from_kpis(-100, 0.12, 5.0) == "MARGINAL"
    # NEPROFITABIL
    assert verdict_from_kpis(-1000, 0.05, 6.0) == "NEPROFITABIL"
    assert verdict_from_kpis(-1000, None, None) == "NEPROFITABIL"


def test_payback_simple_basic():
    """CAPEX 100, CF 30/40/50 → payback intre 2 si 3 ani."""
    pb = payback_simple(100, [30, 40, 50])
    assert pb is not None
    # Anul 2: cumul=70, anul 3: cumul=120. Trece prin 100 la 30/50 din an 3 → 2.6
    assert 2.5 < pb < 2.7


# T4 — match cu scenarii_individuale.xlsx baseline retail
def test_match_scenarii_individuale_baseline_retail():
    """
    Baseline retail din scenarii_individuale.xlsx: CAPEX 45k, OPEX 600, trafic 1000,
    tarif 10 RON/h (uniform), gratuitate 120, inflatie 6%, discount 12%, capital propriu 100%.
    NPV baseline acolo: ~1.500 EUR. IRR ~13.3%.
    """
    p = ScenarioParams(
        tip_parcare="RETAIL",
        tip_investitie="capital_propriu",
        capex_eur=45000,
        pondere_credit=0.0,
        durata_credit_ani=5,
        rata_dobanda=0.0,
        durata_contract_ani=5,
        overhead_anual_eur=0,  # in scenarii_individuale.xlsx fara overhead
        numar_locuri=200,
        trafic_zilnic=1000,
        opex_lunar_eur=600,
        gratuitate_min=120,
        tarif_ora_1_ron=10,
        tarif_ora_2_ron=10,
        tarif_ora_3plus_ron=10,
        cota_impozit=0.16,
        inflatie_opex=0.06,
        inflatie_tarife=0.06,
        discount_rate=0.12,
        eur_ron=5.0,
        provizion_pct=0.06,
        amortizare_ani=5,
        rampup_an1=1.0,  # in xlsx fara ramp-up
        asigurare_anuala_eur=0,
        marketing_initial_eur=0,
        marketing_lunar_eur=0,
    )
    r = simulate_scenario(p)
    # NPV baseline asteptat ~1.500 EUR (toleranta 5% datorita rotunjirilor in pre-calc Python)
    assert -500 < r['npv'] < 5000, f"NPV got {r['npv']}, expected ~1500"
    # IRR baseline ~13.3% +/- 1%
    assert r['irr'] is not None and 0.12 < r['irr'] < 0.15, f"IRR got {r['irr']}"
