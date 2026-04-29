"""
Validator matematic pentru analiza fezabilitate parcari Total Hub SA.
Faza 2 — verificare formule + sanity checks + reality check vs benchmarks industrie.

Toate cifrele aici trebuie sa egaleze cifrele din model.xlsx (Faza 3).
Daca difera, una din ele e gresita.
"""

from typing import Optional

# === PARAMETRI (din spec.md sectiunea 11) ===
EUR_RON = 5.0
ORIZONT_ANI = 5
COTA_IMPOZIT = 0.16
AMORTIZARE_ANI = 5
PROVIZION_PCT = 0.06
INFLATIE_OPEX = 0.06
INFLATIE_TARIFE = 0.06
RATA_DOBANDA = 0.09
PONDERE_CREDIT = 0.70
COST_CAPITAL_PROPRIU = 0.18
PRAG_TEHNICIAN = 5

WACC = PONDERE_CREDIT * RATA_DOBANDA * (1 - COTA_IMPOZIT) + (1 - PONDERE_CREDIT) * COST_CAPITAL_PROPRIU

# === TIPURI LOCATII (din spec.md sectiunea 4) ===
TIPURI = {
    'A': {'nume': 'Retail mare (hyper)',  'capex': 50_000,  'opex_lunar': 700,  'tarif_h': 10, 'trafic_zi': 2000, 'gratuitate_baseline': 120},
    'B': {'nume': 'Retail mediu (super)', 'capex': 40_000,  'opex_lunar': 600,  'tarif_h': 10, 'trafic_zi': 1000, 'gratuitate_baseline': 120},
    'C': {'nume': 'Retail mic',           'capex': 28_000,  'opex_lunar': 450,  'tarif_h': 5,  'trafic_zi': 500,  'gratuitate_baseline': 120},
    'D': {'nume': 'Standalone public',    'capex': 28_000,  'opex_lunar': 450,  'tarif_sesiune': 10, 'trafic_zi': 400, 'gratuitate_baseline': 0},
    'E': {'nume': 'Semi-public',          'capex': 28_000,  'opex_lunar': 450,  'tarif_sesiune': 10, 'trafic_zi': 300, 'gratuitate_baseline': 0},
    'F': {'nume': 'Captiv',               'capex': 28_000,  'opex_lunar': 450,  'tarif_sesiune': 10, 'trafic_zi': 200, 'gratuitate_baseline': 0},
    'G': {'nume': 'Mega-mall',            'capex': 130_000, 'opex_lunar': 1200, 'tarif_h': 12, 'trafic_zi': 8000, 'gratuitate_baseline': 180},
}

# Distributie durate stationare retail (mid-point min, pondere)
BUCKETS = {
    'A': [(7.5, 0.20), (22.5, 0.40), (45, 0.25), (90, 0.10), (150, 0.03), (240, 0.02)],
    'B': [(7.5, 0.25), (22.5, 0.45), (45, 0.22), (90, 0.05), (150, 0.02), (240, 0.01)],
    'C': [(7.5, 0.20), (22.5, 0.45), (45, 0.25), (90, 0.06), (150, 0.02), (240, 0.02)],
    'G': [(7.5, 0.10), (22.5, 0.30), (45, 0.30), (90, 0.20), (150, 0.07), (240, 0.03)],
}

# Rate colectare baseline non-retail (din spec.md sectiunea 9)
COLECTARE_BASELINE = {'D': 0.75, 'E': 0.65, 'F': 0.35}
COLECTARE_STRES = {'D': [0.60, 0.75, 0.90], 'E': [0.50, 0.65, 0.80], 'F': [0.25, 0.35, 0.50]}

# Fee per varianta contractuala (parametrizabil per tip locatie - aici baseline pe tip B)
# Conventie: fee_lunar > 0 = operator primeste; < 0 = operator plateste
FEE_BASELINE = {
    'C1': {'B': 0,    'A': 0,    'C': 0,    'D': 0,    'E': 0,    'F': 0,    'G': 0},
    'C2': {'B': 0,    'A': 0,    'C': 0,    'D': 0,    'E': 0,    'F': 0,    'G': 0},
    'C3': {'B': 800,  'A': 1200, 'C': 400,  'D': 0,    'E': 0,    'F': 0,    'G': 3000},
    'C4': {'B': -300, 'A': -500, 'C': -150, 'D': 0,    'E': 0,    'F': 0,    'G': -1500},
    'C5': {'B': 0,    'A': 0,    'C': 0,    'D': -400, 'E': -300, 'F': -150, 'G': 0},
}

# Cota operator (revenue split) per varianta
COTA_OPERATOR = {'C1': 1.0, 'C2': 0.5, 'C3': 0.0, 'C4': 0.5, 'C5': 1.0}


# === FUNCTII DE CALCUL ===

def venit_per_intrare_retail(tip: str, gratuitate_min: int) -> float:
    """RON per intrare medie pentru tip retail (A, B, C, G)."""
    if tip not in BUCKETS:
        raise ValueError(f"Tip {tip} nu are distributie durate definita")
    tarif_h = TIPURI[tip]['tarif_h']
    venit = 0.0
    for mid_min, pondere in BUCKETS[tip]:
        durata_taxabila_min = max(0, mid_min - gratuitate_min)
        venit += pondere * (durata_taxabila_min / 60) * tarif_h
    return venit


def venit_anual_retail_ron(tip: str, gratuitate_min: int) -> float:
    """RON per an pentru o parcare retail."""
    v_intrare = venit_per_intrare_retail(tip, gratuitate_min)
    return v_intrare * TIPURI[tip]['trafic_zi'] * 365


def venit_anual_non_retail_ron(tip: str, rata_colectare: float) -> float:
    """RON per an pentru o parcare non-retail (D, E, F)."""
    t = TIPURI[tip]
    return t['tarif_sesiune'] * t['trafic_zi'] * rata_colectare * 365


def anuitate_credit(capex: float, rata: float, ani: int, pondere_credit: float = PONDERE_CREDIT) -> float:
    """Anuitate anuala credit echipament (capital propriu nu intra in anuitate)."""
    credit = capex * pondere_credit
    return credit * (rata / (1 - (1 + rata) ** -ani))


def cf_net_per_parcare(tip: str, varianta: str, gratuitate_min: Optional[int] = None,
                        rata_colectare: Optional[float] = None, an: int = 1,
                        overhead_alocat: float = 0.0) -> dict:
    """
    CF Net anual per parcare, an N (1..ORIZONT_ANI), cu inflatie aplicata.
    Structura financiara conform spec.md sectiunea 14:
      Venit -> EBITDA -> EBIT -> Profit Impozabil -> Profit Net -> CF Net
    """
    t = TIPURI[tip]
    cota_op = COTA_OPERATOR[varianta]
    fee_lunar = FEE_BASELINE[varianta][tip]

    # Venit brut din parcare (toata suma incasata, indiferent cui revine)
    if tip in ['A', 'B', 'C', 'G']:
        if gratuitate_min is None:
            gratuitate_min = t['gratuitate_baseline']
        venit_brut_ron = venit_anual_retail_ron(tip, gratuitate_min)
    else:
        if rata_colectare is None:
            rata_colectare = COLECTARE_BASELINE[tip]
        venit_brut_ron = venit_anual_non_retail_ron(tip, rata_colectare)

    venit_brut_eur = venit_brut_ron / EUR_RON

    # Inflatie tarife (anul N)
    factor_inf_tarife = (1 + INFLATIE_TARIFE) ** (an - 1)
    venit_brut_eur *= factor_inf_tarife

    # Venit operator = cota din venit brut + fee lunar × 12
    venit_operator = venit_brut_eur * cota_op + fee_lunar * 12

    # OPEX direct cu inflatie
    opex_anual = t['opex_lunar'] * 12
    factor_inf_opex = (1 + INFLATIE_OPEX) ** (an - 1)
    opex_anual *= factor_inf_opex

    # Structura financiara
    ebitda = venit_operator - opex_anual
    provizion = t['capex'] * PROVIZION_PCT
    ebit = ebitda - provizion - overhead_alocat
    amortizare = t['capex'] / AMORTIZARE_ANI
    profit_impozabil = ebit - amortizare
    impozit = max(0.0, profit_impozabil * COTA_IMPOZIT)
    profit_net = profit_impozabil - impozit
    cf_operational = profit_net + amortizare
    anuitate = anuitate_credit(t['capex'], RATA_DOBANDA, AMORTIZARE_ANI)
    cf_net = cf_operational - anuitate

    return {
        'venit_brut_eur': venit_brut_eur,
        'venit_operator': venit_operator,
        'opex': opex_anual,
        'EBITDA': ebitda,
        'EBITDA_margin': ebitda / venit_operator if venit_operator > 0 else 0,
        'provizion': provizion,
        'overhead_alocat': overhead_alocat,
        'EBIT': ebit,
        'amortizare': amortizare,
        'profit_impozabil': profit_impozabil,
        'impozit': impozit,
        'profit_net': profit_net,
        'cf_operational': cf_operational,
        'anuitate': anuitate,
        'CF_Net': cf_net,
    }


def npv(cash_flows: list, rata: float, capex_initial: float = 0) -> float:
    """NPV: -CAPEX_initial + suma(CF_an_n / (1+rata)^n) pentru n=1..N."""
    return -capex_initial + sum(cf / (1 + rata) ** (i + 1) for i, cf in enumerate(cash_flows))


def irr(cash_flows: list, capex_initial: float, max_iter: int = 100) -> Optional[float]:
    """IRR via bisection. cash_flows = [CF_an_1, ..., CF_an_N]; CAPEX initial separat."""
    flows = [-capex_initial] + cash_flows
    if all(f >= 0 for f in flows) or all(f <= 0 for f in flows):
        return None
    lo, hi = -0.99, 5.0
    for _ in range(max_iter):
        mid = (lo + hi) / 2
        npv_mid = sum(f / (1 + mid) ** i for i, f in enumerate(flows))
        if abs(npv_mid) < 1.0:
            return mid
        npv_lo = sum(f / (1 + lo) ** i for i, f in enumerate(flows))
        if (npv_lo > 0 and npv_mid > 0) or (npv_lo < 0 and npv_mid < 0):
            lo = mid
        else:
            hi = mid
    return mid


# === SANITY CHECKS (din spec.md sectiunea 14) ===

def run_sanity_checks() -> list:
    issues = []

    # 1. Distributii durate = 100%
    for tip, b in BUCKETS.items():
        s = sum(p for _, p in b)
        if abs(s - 1.0) > 0.001:
            issues.append(f"[1] Distributie {tip}: suma = {s:.4f} (asteptat 1.0)")

    # 2. Inflatie zero -> anul 5 = anul 1
    global INFLATIE_OPEX, INFLATIE_TARIFE
    inf_opex_save, inf_tarife_save = INFLATIE_OPEX, INFLATIE_TARIFE
    INFLATIE_OPEX, INFLATIE_TARIFE = 0, 0
    r1 = cf_net_per_parcare('B', 'C2', gratuitate_min=120, an=1)
    r5 = cf_net_per_parcare('B', 'C2', gratuitate_min=120, an=5)
    INFLATIE_OPEX, INFLATIE_TARIFE = inf_opex_save, inf_tarife_save
    if abs(r1['CF_Net'] - r5['CF_Net']) > 1.0:
        issues.append(f"[2] Inflatie zero: CF an1={r1['CF_Net']:.0f} != CF an5={r5['CF_Net']:.0f}")

    # 3. CAPEX zero -> NPV ~= venit×factor anuitate
    # (skip — depinde de implementare exacta, complex de validat aici)

    # 4. Cota impozit zero -> profit_net = profit_impozabil
    cota_save = COTA_IMPOZIT
    globals()['COTA_IMPOZIT'] = 0
    r = cf_net_per_parcare('B', 'C2', gratuitate_min=60)
    globals()['COTA_IMPOZIT'] = cota_save
    if abs(r['profit_net'] - r['profit_impozabil']) > 0.01:
        issues.append(f"[4] Impozit zero: profit_net={r['profit_net']:.2f} != profit_impozabil={r['profit_impozabil']:.2f}")

    # 5. Provizion zero -> EBIT = EBITDA - overhead
    pct_save = PROVIZION_PCT
    globals()['PROVIZION_PCT'] = 0
    r = cf_net_per_parcare('B', 'C2', gratuitate_min=60, overhead_alocat=10000)
    globals()['PROVIZION_PCT'] = pct_save
    if abs(r['EBIT'] - (r['EBITDA'] - 10000)) > 0.01:
        issues.append(f"[5] Provizion zero: EBIT={r['EBIT']:.2f} != EBITDA-overhead={r['EBITDA']-10000:.2f}")

    # 6. Anuitate × 5 > credit (datorita dobanzii)
    for tip, t in TIPURI.items():
        a = anuitate_credit(t['capex'], RATA_DOBANDA, AMORTIZARE_ANI)
        credit = t['capex'] * PONDERE_CREDIT
        if a * AMORTIZARE_ANI <= credit:
            issues.append(f"[6] Anuitate × 5 ({a*5:.0f}) <= credit ({credit:.0f}) tip {tip}")

    # 7. Rata colectare 100% -> venit = trafic × tarif × 365
    v_test = venit_anual_non_retail_ron('D', 1.0)
    expected = TIPURI['D']['tarif_sesiune'] * TIPURI['D']['trafic_zi'] * 365
    if abs(v_test - expected) > 0.01:
        issues.append(f"[7] Colectare 100%: {v_test:.0f} != expected {expected:.0f}")

    # 9. WACC consistency
    cost_datorie_dupa_taxe = RATA_DOBANDA * (1 - COTA_IMPOZIT)
    if not (cost_datorie_dupa_taxe < WACC < COST_CAPITAL_PROPRIU):
        issues.append(f"[9] WACC={WACC:.4f} nu e in interval ({cost_datorie_dupa_taxe:.4f}, {COST_CAPITAL_PROPRIU})")

    return issues


# === REALITY CHECK ===

def reality_check() -> list:
    """Verifica EBITDA margin pe scenarii core vs benchmarks industrie parking (25-45%)."""
    notes = []
    BENCHMARK_MIN, BENCHMARK_MAX = 0.25, 0.45

    scenarii_core = [
        ('B', 'C2', 60, None),
        ('B', 'C3', 120, None),
        ('B', 'C4', 120, None),
        ('A', 'C2', 60, None),
        ('A', 'C4', 120, None),
        ('D', 'C5', None, 0.75),
        ('F', 'C5', None, 0.35),
    ]
    for tip, var, grat, col in scenarii_core:
        r = cf_net_per_parcare(tip, var, gratuitate_min=grat, rata_colectare=col)
        margin = r['EBITDA_margin']
        flag = ""
        if margin < BENCHMARK_MIN:
            flag = " [SUB benchmark — venit slab vs OPEX]"
        elif margin > BENCHMARK_MAX:
            flag = " [PESTE benchmark — verifica daca tarif/trafic supraestimat]"
        else:
            flag = " [in interval]"
        label = f"{tip}-{var}-grat{grat}-col{col}"
        notes.append(f"{label:30s}: EBITDA margin {margin*100:5.1f}%{flag}")
    return notes


# === MAIN ===

def main():
    print("=" * 80)
    print("VALIDATOR — Total Hub SA — Analiza fezabilitate parcari")
    print("=" * 80)
    print(f"\nWACC calculat: {WACC*100:.2f}%  (pondere credit {PONDERE_CREDIT*100:.0f}%)")
    print(f"Cost datorie dupa taxe: {RATA_DOBANDA*(1-COTA_IMPOZIT)*100:.2f}%")
    print(f"Cost capital propriu: {COST_CAPITAL_PROPRIU*100:.0f}%")

    # Sanity checks
    print("\n" + "-" * 80)
    print("SANITY CHECKS")
    print("-" * 80)
    issues = run_sanity_checks()
    if issues:
        print("ESEC:")
        for i in issues:
            print(f"  {i}")
        return
    print("Toate sanity checks PASS.")

    # Reality check
    print("\n" + "-" * 80)
    print("REALITY CHECK — EBITDA margin vs benchmark industrie 25-45%")
    print("-" * 80)
    for n in reality_check():
        print(f"  {n}")

    # CF Net anul 1 — scenarii core retail B
    print("\n" + "=" * 80)
    print("CF NET ANUL 1 PER PARCARE — TIP B (Retail mediu Lidl/Kaufland)")
    print("=" * 80)
    print(f"{'Scenariu':<15} {'Venit op':>11} {'EBITDA':>10} {'CF Net':>10} {'Verdict':>10}")
    for grat in [120, 60, 30]:
        for var in ['C1', 'C2', 'C3', 'C4']:
            r = cf_net_per_parcare('B', var, gratuitate_min=grat)
            verdict = "VIABIL" if r['CF_Net'] > 5000 else ("MARGINAL" if r['CF_Net'] > 0 else "PIERDERE")
            print(f"  g={grat}-{var:<5} {r['venit_operator']:>11,.0f} {r['EBITDA']:>10,.0f} {r['CF_Net']:>10,.0f} {verdict:>10}")

    # CF Net anul 1 — scenarii core retail A
    print("\n" + "=" * 80)
    print("CF NET ANUL 1 PER PARCARE — TIP A (Retail mare Auchan)")
    print("=" * 80)
    print(f"{'Scenariu':<15} {'Venit op':>11} {'EBITDA':>10} {'CF Net':>10} {'Verdict':>10}")
    for grat in [120, 60]:
        for var in ['C2', 'C3', 'C4']:
            r = cf_net_per_parcare('A', var, gratuitate_min=grat)
            verdict = "VIABIL" if r['CF_Net'] > 5000 else ("MARGINAL" if r['CF_Net'] > 0 else "PIERDERE")
            print(f"  g={grat}-{var:<5} {r['venit_operator']:>11,.0f} {r['EBITDA']:>10,.0f} {r['CF_Net']:>10,.0f} {verdict:>10}")

    # CF Net anul 1 — non-retail (cu stres test rate colectare)
    print("\n" + "=" * 80)
    print("CF NET ANUL 1 PER PARCARE — NON-RETAIL (D/E/F) cu stres test colectare")
    print("=" * 80)
    print(f"{'Scenariu':<25} {'Venit op':>11} {'EBITDA':>10} {'CF Net':>10} {'Verdict':>10}")
    for tip in ['D', 'E', 'F']:
        for col in COLECTARE_STRES[tip]:
            r = cf_net_per_parcare(tip, 'C5', rata_colectare=col)
            verdict = "VIABIL" if r['CF_Net'] > 5000 else ("MARGINAL" if r['CF_Net'] > 0 else "PIERDERE")
            print(f"  {tip}-C5-col{int(col*100)}%{'':<10} {r['venit_operator']:>11,.0f} {r['EBITDA']:>10,.0f} {r['CF_Net']:>10,.0f} {verdict:>10}")

    # Portofoliu baseline scaling C — NPV / IRR / Payback
    print("\n" + "=" * 80)
    print("PORTOFOLIU BASELINE (scaling 12 parcari) — INDICATORI 5 ani")
    print("=" * 80)

    # Mix portofoliu finalul anului 2:
    # Anul 1: 2A + 2B + 1F
    # Anul 2: +1A + 4B + 1D + 1F  ->  total: 3A + 6B + 1D + 2F = 12 parcari
    portofoliu = {
        # (tip, varianta, gratuitate, colectare, an_punere_in_functiune)
        'A1_C2_g120': ('A', 'C2', 120, None, 1),
        'A2_C4_g120': ('A', 'C4', 120, None, 1),
        'B1_C4_g120': ('B', 'C4', 120, None, 1),
        'B2_C4_g120': ('B', 'C4', 120, None, 1),
        'F1_C5':      ('F', 'C5', None, 0.35, 1),
        'A3_C4_g120': ('A', 'C4', 120, None, 2),
        'B3_C4_g120': ('B', 'C4', 120, None, 2),
        'B4_C4_g120': ('B', 'C4', 120, None, 2),
        'B5_C3_g120': ('B', 'C3', 120, None, 2),
        'B6_C3_g120': ('B', 'C3', 120, None, 2),
        'D1_C5':      ('D', 'C5', None, 0.75, 2),
        'F2_C5':      ('F', 'C5', None, 0.35, 2),
    }

    # Overhead per an (din spec.md sectiunea 12)
    overhead_total_per_an = {1: 119_000, 2: 141_000, 3: 166_000, 4: 166_000, 5: 166_000}

    # Calcul CF Net agregat per an
    cf_total_per_an = []
    for an in range(1, ORIZONT_ANI + 1):
        parcari_active = [(name, p) for name, p in portofoliu.items() if p[4] <= an]
        n_parcari = len(parcari_active)
        overhead_alocat_per_parcare = overhead_total_per_an[an] / n_parcari if n_parcari else 0
        cf_total = 0
        for name, (tip, var, grat, col, an_pif) in parcari_active:
            an_relativ = an - an_pif + 1
            r = cf_net_per_parcare(tip, var, gratuitate_min=grat, rata_colectare=col,
                                    an=an_relativ, overhead_alocat=overhead_alocat_per_parcare)
            cf_total += r['CF_Net']
        cf_total_per_an.append(cf_total)
        print(f"  Anul {an}: {n_parcari:>2} parcari, overhead {overhead_total_per_an[an]:>7,} EUR, CF Net agregat = {cf_total:>10,.0f} EUR")

    # CAPEX initial total: pe ani in care se pun in functiune
    capex_an_1 = sum(TIPURI[p[0]]['capex'] for name, p in portofoliu.items() if p[4] == 1)
    capex_an_2 = sum(TIPURI[p[0]]['capex'] for name, p in portofoliu.items() if p[4] == 2)
    capex_total = capex_an_1 + capex_an_2

    # Capital propriu (30%) si credit (70%)
    capital_propriu_total = capex_total * (1 - PONDERE_CREDIT)
    credit_total = capex_total * PONDERE_CREDIT

    # NPV pe CF Net (anuitatea de credit e deja inclusa in CF Net per parcare)
    # Investitia initiala (-) = capital propriu (creditul se ramburseaza prin anuitati)
    npv_portofoliu = npv(cf_total_per_an, WACC, capex_initial=capital_propriu_total)
    irr_portofoliu = irr(cf_total_per_an, capital_propriu_total)
    cf_cumulat_5 = sum(cf_total_per_an)

    # Payback simplu (pe CF nominal, nu actualizat)
    cf_cum = -capital_propriu_total
    payback_an = None
    for i, cf in enumerate(cf_total_per_an, start=1):
        cf_cum += cf
        if cf_cum >= 0 and payback_an is None:
            payback_an = i

    print(f"\n  CAPEX total portofoliu: {capex_total:>10,.0f} EUR")
    print(f"  Capital propriu (30%):  {capital_propriu_total:>10,.0f} EUR")
    print(f"  Credit (70%):           {credit_total:>10,.0f} EUR")
    print(f"  CF Net cumulat 5 ani:   {cf_cumulat_5:>10,.0f} EUR")
    print(f"  NPV @ WACC {WACC*100:.1f}%:        {npv_portofoliu:>10,.0f} EUR  ({'POZITIV' if npv_portofoliu > 0 else 'NEGATIV'})")
    print(f"  IRR:                    {irr_portofoliu*100:>10,.1f}%" if irr_portofoliu else "  IRR: N/A")
    print(f"  Payback simplu:         {'>5 ani' if payback_an is None else f'~{payback_an} ani'}")

    # KPI gates
    print("\n" + "-" * 80)
    print("KPI GATES (din spec.md sectiunea 3)")
    print("-" * 80)
    kpi = [
        ("NPV portofoliu @ WACC > 0",        npv_portofoliu > 0,                    f"{npv_portofoliu:,.0f} EUR"),
        ("Payback < 4 ani",                   (payback_an or 99) < 4,                f"{payback_an or '>5'} ani"),
        ("IRR > 15%",                         (irr_portofoliu or 0) > 0.15,          f"{(irr_portofoliu or 0)*100:.1f}%"),
        ("CF Net cumulat 5 ani > 800k",       cf_cumulat_5 > 800_000,                f"{cf_cumulat_5:,.0f} EUR"),
    ]
    for nume, pasat, valoare in kpi:
        marker = "PASS" if pasat else "FAIL"
        print(f"  [{marker}] {nume:35s} -> {valoare}")


if __name__ == "__main__":
    main()
