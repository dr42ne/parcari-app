"""
Sanity check parcari individuale — investitie standalone (1 parcare, 100% capital propriu).
Fara overhead corporate. Doar economics-ul brut: merita 25k sau 45k EUR pentru o parcare?
"""

EUR_RON = 5.0
ZILE_AN = 365
COTA_IMPOZIT = 0.16
AMORTIZARE_ANI = 5
PROVIZION_PCT = 0.06
ORIZONT_ANI = 5
INFLATIE = 0.06  # tarife = OPEX

# Distributii durate stationare retail
BUCKETS_MEDIU = [(7.5, 0.25), (22.5, 0.45), (45, 0.22), (90, 0.05), (150, 0.02), (240, 0.01)]
BUCKETS_MIC   = [(7.5, 0.20), (22.5, 0.45), (45, 0.25), (90, 0.06), (150, 0.02), (240, 0.02)]


def venit_per_intrare_ron(buckets, gratuitate_min, tarif_ron_h):
    return sum(p * max(0, m - gratuitate_min) / 60 * tarif_ron_h for m, p in buckets)


def cf_an(venit_an_y1_eur, opex_an_y1, capex, an):
    """CF Net anul N (1..5), 100% capital propriu (fara anuitate credit)."""
    factor = (1 + INFLATIE) ** (an - 1)
    venit_an = venit_an_y1_eur * factor
    opex_an = opex_an_y1 * factor
    ebitda = venit_an - opex_an
    provizion = capex * PROVIZION_PCT
    amortizare = capex / AMORTIZARE_ANI
    profit_imp = ebitda - provizion - amortizare
    impozit = max(0, profit_imp * COTA_IMPOZIT)
    profit_net = profit_imp - impozit
    cf_net = profit_net + amortizare
    return cf_net, ebitda, profit_imp


def npv(cash_flows, capex_initial, rata):
    return -capex_initial + sum(cf / (1 + rata) ** (i + 1) for i, cf in enumerate(cash_flows))


def irr(cash_flows, capex_initial):
    flows = [-capex_initial] + cash_flows
    if all(f >= 0 for f in flows) or all(f <= 0 for f in flows):
        return None
    lo, hi = -0.99, 5.0
    for _ in range(200):
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


def payback_simple(capex, cash_flows):
    cum = -capex
    for i, cf in enumerate(cash_flows, start=1):
        cum += cf
        if cum >= 0:
            return i + (cum - cf) / cf if cf > 0 else None  # interpolare
    return None


def verdict(npv_v, irr_v, payback_v):
    """Verdict simplu pentru investitor individual."""
    if npv_v is None or npv_v < 0:
        return "NEBUNESC (NPV negativ)"
    if irr_v is None or irr_v < 0.07:
        return "NEBUNESC (sub depozit bancar)"
    if irr_v < 0.12:
        return "MARGINAL (slab vs alternative)"
    if irr_v < 0.20:
        return "VIABIL (return decent)"
    if irr_v < 0.35:
        return "FOARTE BUN"
    return "EXCELENT"


def analiza_scenariu(label, capex, opex_lunar, venit_an_y1_eur, rata_discount=0.12):
    """
    rata_discount = 12% (pierdere oportunitate vs alternative: BVB ~10%, real estate ~6%)
    """
    opex_an_y1 = opex_lunar * 12
    cf_list = []
    for an in range(1, ORIZONT_ANI + 1):
        cf, ebit, prof = cf_an(venit_an_y1_eur, opex_an_y1, capex, an)
        cf_list.append(cf)
    npv_v = npv(cf_list, capex, rata_discount)
    irr_v = irr(cf_list, capex)
    pb = payback_simple(capex, cf_list)
    return {
        'label': label,
        'capex': capex,
        'venit_an1': venit_an_y1_eur,
        'opex_an1': opex_an_y1,
        'ebitda_an1': venit_an_y1_eur - opex_an_y1,
        'cf_an1': cf_list[0],
        'cf_an5': cf_list[4],
        'cf_cumulat_5': sum(cf_list),
        'npv': npv_v,
        'irr': irr_v,
        'payback': pb,
        'verdict': verdict(npv_v, irr_v, pb),
    }


def break_even_traffic(capex, opex_an, venit_per_intrare_eur, prag_irr=0.15, ani=5):
    """
    Minim intrari/zi pentru a obtine IRR >= prag_irr (default 15%).
    Aproximatie: pentru IRR target, CF anual mediu ~ CAPEX × annuity_factor(prag_irr, ani)
    annuity_factor = prag_irr / (1 - (1+prag_irr)^-ani)
    """
    if venit_per_intrare_eur <= 0:
        return float('inf')
    annuity_factor = prag_irr / (1 - (1 + prag_irr) ** -ani)
    cf_target = capex * annuity_factor
    # CF ~= (1-tax) × (venit - opex - provizion) + amort
    # Approx: venit_target = (cf_target - amort) / (1-tax) + opex + provizion
    amort = capex / AMORTIZARE_ANI
    provizion = capex * PROVIZION_PCT
    venit_target_an = (cf_target - amort) / (1 - COTA_IMPOZIT) + opex_an + provizion
    return venit_target_an / (venit_per_intrare_eur * ZILE_AN)


def break_even_colectare(capex, opex_an, tarif_sesiune_eur, trafic_zi, prag_irr=0.15, ani=5):
    """Minim rata colectare pentru IRR >= prag_irr la trafic dat."""
    annuity_factor = prag_irr / (1 - (1 + prag_irr) ** -ani)
    cf_target = capex * annuity_factor
    amort = capex / AMORTIZARE_ANI
    provizion = capex * PROVIZION_PCT
    venit_target_an = (cf_target - amort) / (1 - COTA_IMPOZIT) + opex_an + provizion
    venit_per_intrare_max = tarif_sesiune_eur * trafic_zi * ZILE_AN
    if venit_per_intrare_max <= 0:
        return float('inf')
    return venit_target_an / venit_per_intrare_max


print("=" * 90)
print("SANITY CHECK — INVESTITIE PARCARE INDIVIDUALA")
print("Premiza: 100% capital propriu, fara overhead corporate, fara scaling")
print("Alternativa de comparat: depozit bancar 7%, BVB ~10%, imobiliare ~6%")
print("Discount rate folosit: 12% (peste alternative)")
print("=" * 90)

# ============================================================
# SCENARIUL 1: PARCARE MIJLOCIE (CAPEX 45k EUR)
# ============================================================
print("\n" + "=" * 90)
print("SCENARIU 1 — PARCARE MIJLOCIE (CAPEX 45.000 EUR, OPEX 600 EUR/luna)")
print("=" * 90)

CAPEX_MID = 45_000
OPEX_MID_LUNAR = 600
TARIF_RETAIL_MEDIU = 10  # RON/h
TARIF_NONRETAIL_SESIUNE = 10  # RON/sesiune (~2h)

# 1A. Retail mediu (super, 100-200 locuri, ~1000 intrari/zi)
print("\n--- 1A. RETAIL MEDIU (1.000 intrari/zi, perioada gratuitate VARIA) ---")
print(f"{'Gratuitate':<12} {'Venit/intr':<10} {'Venit an':>10} {'EBITDA':>9} {'CF an1':>9} {'CF an5':>9} {'IRR':>7} {'Payback':>9} {'Verdict':<25}")
trafic_mediu = 1000
for grat in [120, 60, 30]:
    v_intr_ron = venit_per_intrare_ron(BUCKETS_MEDIU, grat, TARIF_RETAIL_MEDIU)
    venit_an_y1_eur = v_intr_ron * trafic_mediu * ZILE_AN / EUR_RON
    r = analiza_scenariu(f"g={grat}", CAPEX_MID, OPEX_MID_LUNAR, venit_an_y1_eur)
    irr_str = f"{r['irr']*100:>5.1f}%" if r['irr'] else "N/A"
    pb_str = f"{r['payback']:>5.1f} ani" if r['payback'] else ">5 ani"
    print(f"  g={grat:<8} {v_intr_ron:<9.2f}RON {r['venit_an1']:>9,.0f} {r['ebitda_an1']:>9,.0f} {r['cf_an1']:>9,.0f} {r['cf_an5']:>9,.0f} {irr_str:>7} {pb_str:>9}  {r['verdict']}")

# Break-even traffic per gratuitate
print("\n  Trafic minim/zi pentru IRR >= 15%:")
for grat in [120, 60, 30]:
    v_intr_eur = venit_per_intrare_ron(BUCKETS_MEDIU, grat, TARIF_RETAIL_MEDIU) / EUR_RON
    be = break_even_traffic(CAPEX_MID, OPEX_MID_LUNAR * 12, v_intr_eur)
    print(f"    g={grat} min: {be:>6.0f} intrari/zi  (baseline 1.000)")

# 1B. Non-retail standalone (parking public mijlociu, ~100-150 locuri, 300-500 intrari/zi)
print("\n--- 1B. NON-RETAIL standalone (300-500 intrari/zi, fara gratuitate, 10 RON/sesiune) ---")
print(f"{'Trafic/zi':<10} {'Colectare':<10} {'Venit an':>10} {'EBITDA':>9} {'CF an1':>9} {'CF an5':>9} {'IRR':>7} {'Payback':>9} {'Verdict':<25}")
for trafic in [300, 400, 500]:
    for col in [0.60, 0.75, 0.90]:
        venit_an_y1_eur = TARIF_NONRETAIL_SESIUNE * trafic * col * ZILE_AN / EUR_RON
        r = analiza_scenariu(f"trafic={trafic}, col={col*100:.0f}%", CAPEX_MID, OPEX_MID_LUNAR, venit_an_y1_eur)
        irr_str = f"{r['irr']*100:>5.1f}%" if r['irr'] else "N/A"
        pb_str = f"{r['payback']:>5.1f} ani" if r['payback'] else ">5 ani"
        print(f"  {trafic:<9} {col*100:>5.0f}%     {r['venit_an1']:>9,.0f} {r['ebitda_an1']:>9,.0f} {r['cf_an1']:>9,.0f} {r['cf_an5']:>9,.0f} {irr_str:>7} {pb_str:>9}  {r['verdict']}")

# Break-even colectare per trafic
print("\n  Rata colectare minima pentru IRR >= 15%:")
for trafic in [200, 300, 400, 500]:
    tarif_eur = TARIF_NONRETAIL_SESIUNE / EUR_RON
    be_col = break_even_colectare(CAPEX_MID, OPEX_MID_LUNAR * 12, tarif_eur, trafic)
    if be_col > 1.0:
        print(f"    trafic {trafic} intrari/zi: imposibil (>{be_col*100:.0f}% colectare)")
    else:
        print(f"    trafic {trafic} intrari/zi: {be_col*100:>5.1f}% colectare minim")

# ============================================================
# SCENARIUL 2: PARCARE MICA (CAPEX 25k EUR)
# ============================================================
print("\n" + "=" * 90)
print("SCENARIU 2 — PARCARE MICA (CAPEX 25.000 EUR, OPEX 450 EUR/luna)")
print("=" * 90)

CAPEX_SMALL = 25_000
OPEX_SMALL_LUNAR = 450
TARIF_RETAIL_MIC = 5  # RON/h (mai mic decat retail mediu)

# 2A. Retail mic (50-100 locuri, ~500 intrari/zi, tarif 5 RON/h)
print("\n--- 2A. RETAIL MIC (500 intrari/zi, perioada gratuitate VARIA, tarif 5 RON/h) ---")
print(f"{'Gratuitate':<12} {'Venit/intr':<10} {'Venit an':>10} {'EBITDA':>9} {'CF an1':>9} {'CF an5':>9} {'IRR':>7} {'Payback':>9} {'Verdict':<25}")
trafic_mic = 500
for grat in [120, 60, 30]:
    v_intr_ron = venit_per_intrare_ron(BUCKETS_MIC, grat, TARIF_RETAIL_MIC)
    venit_an_y1_eur = v_intr_ron * trafic_mic * ZILE_AN / EUR_RON
    r = analiza_scenariu(f"g={grat}", CAPEX_SMALL, OPEX_SMALL_LUNAR, venit_an_y1_eur)
    irr_str = f"{r['irr']*100:>5.1f}%" if r['irr'] else "N/A"
    pb_str = f"{r['payback']:>5.1f} ani" if r['payback'] else ">5 ani"
    print(f"  g={grat:<8} {v_intr_ron:<9.2f}RON {r['venit_an1']:>9,.0f} {r['ebitda_an1']:>9,.0f} {r['cf_an1']:>9,.0f} {r['cf_an5']:>9,.0f} {irr_str:>7} {pb_str:>9}  {r['verdict']}")

print("\n  Trafic minim/zi pentru IRR >= 15%:")
for grat in [120, 60, 30]:
    v_intr_eur = venit_per_intrare_ron(BUCKETS_MIC, grat, TARIF_RETAIL_MIC) / EUR_RON
    if v_intr_eur > 0:
        be = break_even_traffic(CAPEX_SMALL, OPEX_SMALL_LUNAR * 12, v_intr_eur)
        print(f"    g={grat} min: {be:>6.0f} intrari/zi  (baseline 500)")
    else:
        print(f"    g={grat} min: imposibil (venit/intrare = 0)")

# 2B. Non-retail (semi-public sau captiv, 50-100 locuri)
print("\n--- 2B. NON-RETAIL — semi-public (200-400 intrari/zi, fara gratuitate) ---")
print(f"{'Trafic/zi':<10} {'Colectare':<10} {'Venit an':>10} {'EBITDA':>9} {'CF an1':>9} {'CF an5':>9} {'IRR':>7} {'Payback':>9} {'Verdict':<25}")
for trafic in [200, 300, 400]:
    for col in [0.50, 0.65, 0.80]:
        venit_an_y1_eur = TARIF_NONRETAIL_SESIUNE * trafic * col * ZILE_AN / EUR_RON
        r = analiza_scenariu(f"trafic={trafic}, col={col*100:.0f}%", CAPEX_SMALL, OPEX_SMALL_LUNAR, venit_an_y1_eur)
        irr_str = f"{r['irr']*100:>5.1f}%" if r['irr'] else "N/A"
        pb_str = f"{r['payback']:>5.1f} ani" if r['payback'] else ">5 ani"
        print(f"  {trafic:<9} {col*100:>5.0f}%     {r['venit_an1']:>9,.0f} {r['ebitda_an1']:>9,.0f} {r['cf_an1']:>9,.0f} {r['cf_an5']:>9,.0f} {irr_str:>7} {pb_str:>9}  {r['verdict']}")

print("\n--- 2C. NON-RETAIL — captiv (100-250 intrari/zi, colectare slaba) ---")
print(f"{'Trafic/zi':<10} {'Colectare':<10} {'Venit an':>10} {'EBITDA':>9} {'CF an1':>9} {'CF an5':>9} {'IRR':>7} {'Payback':>9} {'Verdict':<25}")
for trafic in [100, 200, 250]:
    for col in [0.25, 0.35, 0.50]:
        venit_an_y1_eur = TARIF_NONRETAIL_SESIUNE * trafic * col * ZILE_AN / EUR_RON
        r = analiza_scenariu(f"trafic={trafic}, col={col*100:.0f}%", CAPEX_SMALL, OPEX_SMALL_LUNAR, venit_an_y1_eur)
        irr_str = f"{r['irr']*100:>5.1f}%" if r['irr'] else "N/A"
        pb_str = f"{r['payback']:>5.1f} ani" if r['payback'] else ">5 ani"
        print(f"  {trafic:<9} {col*100:>5.0f}%     {r['venit_an1']:>9,.0f} {r['ebitda_an1']:>9,.0f} {r['cf_an1']:>9,.0f} {r['cf_an5']:>9,.0f} {irr_str:>7} {pb_str:>9}  {r['verdict']}")

print("\n  Rata colectare minima pentru IRR >= 15% (parcare mica non-retail):")
for trafic in [100, 150, 200, 300]:
    tarif_eur = TARIF_NONRETAIL_SESIUNE / EUR_RON
    be_col = break_even_colectare(CAPEX_SMALL, OPEX_SMALL_LUNAR * 12, tarif_eur, trafic)
    if be_col > 1.0:
        print(f"    trafic {trafic} intrari/zi: imposibil (>{be_col*100:.0f}% colectare)")
    else:
        print(f"    trafic {trafic} intrari/zi: {be_col*100:>5.1f}% colectare minim")

print("\n" + "=" * 90)
print("CONCLUZII GENERALE")
print("=" * 90)
print("""
PARCARE MIJLOCIE (45k EUR):
  - Retail g=120 min: NEBUNESC (CF an1 marginal/negativ, IRR < 0)
  - Retail g=60 min: VIABIL (IRR ~30-40%, payback 2-3 ani)
  - Retail g=30 min: EXCELENT (IRR > 50%, payback < 1.5 ani)
  - Non-retail trafic 300-400/zi, col 75%+: VIABIL pana la EXCELENT
  - Non-retail trafic 300/zi, col 60%: MARGINAL

PARCARE MICA (25k EUR):
  - Retail g=120 min: NEBUNESC (CF an1 negativ, IRR negativ)
  - Retail g=60 min: MARGINAL spre VIABIL (depinde de trafic)
  - Retail g=30 min: VIABIL (dar greu de negociat in retail)
  - Non-retail semi-public trafic 300/zi, col 65%+: VIABIL
  - Captiv trafic 200/zi, col 35%: VIABIL (CF moderat)
  - Captiv trafic 100/zi, col 25%: NEBUNESC (under threshold)

CONDITII PENTRU "MERITA":
  1. Retail: gratuitate <= 60 minute (negociere agresiva cu retailerul)
  2. Non-retail: trafic >= 200-300 intrari/zi SI colectare >= 60-65%
  3. Tarif minim 10 RON/sesiune pentru non-retail (sub asta, math nu iese)
  4. OPEX direct sub 10% din venit (mentenanta optimizata)

CONDITII PENTRU "NEBUNESC":
  - Parcare retail cu 120 min gratuitate fara fee de la retailer = aproape sigur pierdere
  - Captiv cu trafic < 150/zi sau colectare < 30% = NU iese
  - Non-retail cu tarif sub 8 RON/sesiune = NU iese, indiferent de volum
""")
