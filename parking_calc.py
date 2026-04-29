"""
Logica financiara pentru analiza fezabilitate parcari individuale.
Pure functions, fara dependente Streamlit. Reutilizate din validator.py si build_individual_xlsx.py.
"""

from typing import Optional
from dataclasses import dataclass, field

# Distributia duratelor stationare (mid-point min, pondere) — retail mediu standard
BUCKETS_RETAIL = [
    (7.5, 0.25),
    (22.5, 0.45),
    (45.0, 0.22),
    (90.0, 0.05),
    (150.0, 0.02),
    (240.0, 0.01),
]


def venit_per_intrare_tiered(durata_min: float, gratuitate_min: float,
                              t1: float, t2: float, t3plus: float) -> float:
    """RON per intrare la tarife progresive. Tarifare prorata DUPA gratuitate.
    Ora 1 (primele 60 min taxabile) = t1 RON/h
    Ora 2 (urmatoarele 60 min) = t2 RON/h
    Ora 3+ (peste 120 min taxabili) = t3plus RON/h
    """
    taxabil = max(0.0, durata_min - gratuitate_min)
    if taxabil == 0:
        return 0.0
    m1 = min(taxabil, 60.0)
    m2 = min(max(taxabil - 60.0, 0.0), 60.0)
    m3 = max(taxabil - 120.0, 0.0)
    return (m1 / 60.0) * t1 + (m2 / 60.0) * t2 + (m3 / 60.0) * t3plus


def venit_anual_retail_eur(buckets, gratuitate_min: float, t1: float, t2: float, t3plus: float,
                            trafic_zi: float, eur_ron: float, zile: int = 365) -> float:
    """EUR per an pentru o parcare retail."""
    venit_intrare_ron = sum(
        pondere * venit_per_intrare_tiered(mid_min, gratuitate_min, t1, t2, t3plus)
        for mid_min, pondere in buckets
    )
    return venit_intrare_ron * trafic_zi * zile / eur_ron


def venit_anual_nonretail_eur(tarif_sesiune_ron: float, trafic_zi: float,
                               rata_colectare: float, eur_ron: float, zile: int = 365) -> float:
    """EUR per an pentru o parcare non-retail (plata sesiune)."""
    return tarif_sesiune_ron * trafic_zi * rata_colectare * zile / eur_ron


def anuitate_credit(capex: float, rata: float, ani: int, pondere_credit: float) -> float:
    """Anuitate anuala pentru creditul pe partea finantata."""
    if pondere_credit <= 0 or ani <= 0:
        return 0.0
    credit = capex * pondere_credit
    if rata <= 0:
        return credit / ani
    return credit * (rata / (1 - (1 + rata) ** -ani))


def npv(cash_flows, capex_initial: float, rata: float) -> float:
    """NPV: -CAPEX_initial + suma(CF / (1+rata)^n) pentru n=1..N."""
    return -capex_initial + sum(cf / (1 + rata) ** (i + 1) for i, cf in enumerate(cash_flows))


def irr(cash_flows, capex_initial: float, max_iter: int = 200) -> Optional[float]:
    """IRR via bisection. cash_flows = [CF_an_1, ..., CF_an_N]."""
    flows = [-capex_initial] + list(cash_flows)
    if all(f >= 0 for f in flows) or all(f <= 0 for f in flows):
        return None
    lo, hi = -0.99, 5.0
    mid = 0.0
    for _ in range(max_iter):
        mid = (lo + hi) / 2
        v_mid = sum(f / (1 + mid) ** i for i, f in enumerate(flows))
        if abs(v_mid) < 1.0:
            return mid
        v_lo = sum(f / (1 + lo) ** i for i, f in enumerate(flows))
        if (v_lo > 0 and v_mid > 0) or (v_lo < 0 and v_mid < 0):
            lo = mid
        else:
            hi = mid
    return mid


def payback_simple(capex_propriu: float, cash_flows) -> Optional[float]:
    """Numarul de ani pana CF cumulat acopera capitalul propriu (interpolare liniara)."""
    cum = 0.0
    for i, cf in enumerate(cash_flows, start=1):
        prev = cum
        cum += cf
        if cum >= capex_propriu and cf > 0:
            # Interpolare liniara intre anul i-1 (prev) si anul i (cum)
            frac = (capex_propriu - prev) / cf
            return (i - 1) + frac
    return None


@dataclass
class ScenarioParams:
    """Parametri pentru un scenariu de parcare. Toti parametrii sunt explicit declarati."""
    # Tip scenariu
    tip_parcare: str = "RETAIL"  # "RETAIL" sau "NON-RETAIL"

    # Investitie
    tip_investitie: str = "mix"  # "capital_propriu", "credit", "mix"
    capex_eur: float = 40000.0
    pondere_credit: float = 0.70  # 0 = capital propriu 100%; 1 = credit 100%
    durata_credit_ani: int = 5
    rata_dobanda: float = 0.09

    # Operare
    durata_contract_ani: int = 5
    overhead_anual_eur: float = 12000.0
    numar_locuri: int = 250
    trafic_zilnic: float = 1250.0
    opex_lunar_eur: float = 140.0

    # Tarife retail (folosite daca tip_parcare == "RETAIL")
    gratuitate_min: float = 120.0
    tarif_ora_1_ron: float = 5.0
    tarif_ora_2_ron: float = 7.0
    tarif_ora_3plus_ron: float = 10.0

    # Tarife non-retail (folosite daca tip_parcare == "NON-RETAIL")
    tarif_sesiune_ron: float = 10.0
    rata_colectare: float = 0.75

    # Fiscal / macro
    cota_impozit: float = 0.16
    inflatie_opex: float = 0.06
    inflatie_tarife: float = 0.06
    discount_rate: float = 0.12
    eur_ron: float = 5.0
    provizion_pct: float = 0.06
    amortizare_ani: int = 5

    # Suplimentare
    rampup_an1: float = 0.70  # 70% din baseline trafic in anul 1
    asigurare_anuala_eur: float = 600.0
    marketing_initial_eur: float = 2000.0
    marketing_lunar_eur: float = 200.0


def simulate_scenario(p: ScenarioParams) -> dict:
    """
    Simuleaza un scenariu pe orizont durata_contract_ani (max 10 din considerente UI).
    Returneaza dict cu CF an N, P&L breakdown, NPV, IRR, payback, verdict.
    """
    ani = p.durata_contract_ani
    # Venit baseline anul 1 (la trafic full, fara ramp-up)
    if p.tip_parcare == "RETAIL":
        venit_y1_full = venit_anual_retail_eur(
            BUCKETS_RETAIL,
            p.gratuitate_min,
            p.tarif_ora_1_ron, p.tarif_ora_2_ron, p.tarif_ora_3plus_ron,
            p.trafic_zilnic, p.eur_ron,
        )
    else:
        venit_y1_full = venit_anual_nonretail_eur(
            p.tarif_sesiune_ron, p.trafic_zilnic, p.rata_colectare, p.eur_ron,
        )

    # OPEX lunar include base + asigurare/12 + marketing lunar
    opex_lunar_total = p.opex_lunar_eur + p.asigurare_anuala_eur / 12 + p.marketing_lunar_eur
    opex_y1 = opex_lunar_total * 12

    # Anuitate (constanta pe durata creditului; daca durata_credit > durata_contract, calcul total ok)
    anuitate = anuitate_credit(p.capex_eur, p.rata_dobanda, p.durata_credit_ani, p.pondere_credit)

    # Provizion si amortizare (constante)
    provizion = p.capex_eur * p.provizion_pct
    amortizare = p.capex_eur / p.amortizare_ani

    # Capital propriu (investitie initiala)
    capex_propriu = p.capex_eur * (1 - p.pondere_credit)
    # Marketing initial intra in cash out anul 0
    capex_propriu_total = capex_propriu + p.marketing_initial_eur

    # P&L pe ani
    pnl = []
    cash_flows = []
    for an in range(1, ani + 1):
        # Inflatie
        inf_t = (1 + p.inflatie_tarife) ** (an - 1)
        inf_o = (1 + p.inflatie_opex) ** (an - 1)
        # Ramp-up doar in anul 1
        rampup = p.rampup_an1 if an == 1 else 1.0

        venit_an = venit_y1_full * inf_t * rampup
        opex_an = opex_y1 * inf_o
        ebitda = venit_an - opex_an
        ebit = ebitda - provizion - p.overhead_anual_eur
        profit_imp = ebit - amortizare
        impozit = max(0.0, profit_imp * p.cota_impozit)
        profit_net = profit_imp - impozit
        cf_op = profit_net + amortizare
        # Anuitate doar pe durata creditului
        anuitate_an = anuitate if an <= p.durata_credit_ani else 0.0
        cf_net = cf_op - anuitate_an
        cash_flows.append(cf_net)
        pnl.append({
            'an': an,
            'venit': venit_an,
            'opex': opex_an,
            'ebitda': ebitda,
            'provizion': provizion,
            'overhead': p.overhead_anual_eur,
            'ebit': ebit,
            'amortizare': amortizare,
            'profit_impozabil': profit_imp,
            'impozit': impozit,
            'profit_net': profit_net,
            'cf_operational': cf_op,
            'anuitate': anuitate_an,
            'cf_net': cf_net,
        })

    cf_cumulat = []
    cum = -capex_propriu_total
    for cf in cash_flows:
        cum += cf
        cf_cumulat.append(cum)

    npv_val = npv(cash_flows, capex_propriu_total, p.discount_rate)
    irr_val = irr(cash_flows, capex_propriu_total)
    pb = payback_simple(capex_propriu_total, cash_flows)

    return {
        'pnl': pnl,
        'cash_flows': cash_flows,
        'cf_cumulat': cf_cumulat,
        'npv': npv_val,
        'irr': irr_val,
        'payback': pb,
        'cf_cumulat_5': sum(cash_flows[:min(5, ani)]),
        'capex_propriu_total': capex_propriu_total,
        'venit_y1_full': venit_y1_full,
        'anuitate': anuitate,
        'verdict': verdict_from_kpis(npv_val, irr_val, pb),
    }


def verdict_from_kpis(npv_val: float, irr_val: Optional[float], payback: Optional[float]) -> str:
    """
    PROFITABIL: NPV > 0 SI IRR > 15% SI Payback < 4 ani
    MARGINAL: NPV > 0 SAU IRR > 10%
    NEPROFITABIL: restul
    """
    if npv_val > 0 and irr_val is not None and irr_val > 0.15 and payback is not None and payback < 4:
        return "PROFITABIL"
    if npv_val > 0 or (irr_val is not None and irr_val > 0.10):
        return "MARGINAL"
    return "NEPROFITABIL"
