"""
Tooltip-uri pentru fiecare parametru afisat in UI Streamlit.
Format: scurt, contextual, cu valori tipice si referinta industrie.
"""

TOOLTIPS = {
    # Investitie
    "tip_investitie": "Capital propriu = 100% bani din buzunar. Credit = 100% finantare bancara. Mix = combinatie (default 70% credit + 30% propriu).",
    "capex_eur": "Investitia totala in echipamente (ANPR, bariere, terminale, instalare). Tipic: 25k-50k pentru parcare mica/medie, 50k-130k pentru retail mare/mall.",
    "pondere_credit": "Procent din CAPEX finantat prin credit. 0% = totul cash; 70% = standard pentru o investitie cu garantii echipament.",
    "durata_credit_ani": "Durata creditului echipament. Standard in Romania: 5 ani pentru ANPR + bariere.",
    "rata_dobanda": "Rata dobanda anuala credit echipament. Tipic 8-10% in Romania 2026 (IMM cu garantie hardware).",

    # Operare
    "durata_contract_ani": "Durata contractului cu proprietarul / orizontul analizei NPV. Standard parking retail: 5 ani.",
    "overhead_anual_eur": "Costuri corporate alocate (contabilitate, juridic, management partial). Pentru o parcare standalone: 0-5k EUR daca operezi solo, 10-20k daca implici companie.",
    "numar_locuri": "Numar locuri fizice. Influenteaza CAPEX si trafic potential maxim.",
    "trafic_zilnic": "Numar mediu de intrari/zi. Validare critica! Verifica cu counter sau observatie 2 saptamani inainte de semnare.",
    "opex_lunar_eur": "Costuri operationale recurente: energie, internet, mentenanta, software. Tipic 100-700 EUR/luna in functie de marime.",

    # Tarife retail
    "gratuitate_min": "Minute de stationare gratuita. Standard piata romanesti retail: 120 min. Negocierea sub 120 min este parametrul cu cel mai mare impact pe NPV (vezi tornado).",
    "tarif_ora_1_ron": "Tarif primele 60 min taxabile (dupa gratuitate). Standard 5 RON/h pentru parcari mici, 10 RON/h pentru retail mare.",
    "tarif_ora_2_ron": "Tarif urmatoarele 60 min. Tipic mai mare ca ora 1 pentru a descuraja statie lunga (ex: 7-8 RON/h).",
    "tarif_ora_3plus_ron": "Tarif peste 2h taxabile. Cel mai mare pentru a descuraja parazitismul lung (ex: 10-15 RON/h).",

    # Tarife non-retail
    "tarif_sesiune_ron": "Tarif mediu/sesiune (calculat la durata medie a stationarii). Tipic 5-15 RON pentru parcari publice/captive.",
    "rata_colectare": "Procent intrari care platesc efectiv. Tipic in Romania: 60-90% pentru parcari private cu ANPR + bariera; 35-50% pentru parcari captive (hotel/birou) unde personal plateste mai rar.",

    # Fiscal
    "cota_impozit": "Cota impozit pe profit. SA standard: 16%. SRL micro 1-3% pentru CA < 250k EUR (nu se aplica la SA).",
    "inflatie_opex": "Inflatie anuala OPEX (energie, mentenanta, salarii). Tipic 5-7% in Romania 2026.",
    "inflatie_tarife": "Inflatie anuala tarife (in conditiile in care contractul permite ajustare anuala). Tipic 5-7%.",
    "discount_rate": "Rata de actualizare pentru NPV. Reflecta cost capital propriu sau alternative (bursa, imobiliare). Tipic 10-18%.",
    "eur_ron": "Curs EUR/RON pentru conversie. Aprilie 2026: ~5.0.",
    "provizion_pct": "Procent anual din CAPEX rezervat pentru reparatii / inlocuiri partiale. Hardware ANPR: 6-10%.",
    "amortizare_ani": "Durata amortizare CAPEX echipamente. Standard 5 ani pentru hardware ANPR/bariere.",

    # Suplimentare
    "rampup_an1": "% din traficul baseline atins in anul 1 (perioada de invatare client + reglare operationala). Tipic 60-80%.",
    "asigurare_anuala_eur": "Asigurare echipamente (vandalism, deteriorare, pierdere). Tipic 500-1500 EUR/an.",
    "marketing_initial_eur": "Cheltuiala initiala pentru semnalizare locatie, materiale comunicare client. Tipic 1500-3000 EUR.",
    "marketing_lunar_eur": "Marketing recurrent: campanii, semnalizare, suport client. Tipic 100-300 EUR/luna.",
}


def get(key: str) -> str:
    """Returneaza tooltip pentru un parametru, sau string gol daca nu exista."""
    return TOOLTIPS.get(key, "")


# ============================================================
# Tooltip-uri pentru celule Excel (scenarii_individuale.xlsx).
# Cheia = textul EXACT al label-ului asa cum apare in foaie.
# Atasate ca Comment openpyxl → apar pe hover in Excel/LibreOffice.
# ============================================================
XLSX_TOOLTIPS = {
    # --- Sheet "Parametri" — sectiunea A. economici globali ---
    "Curs EUR/RON": "Curs de schimb folosit pentru conversia tarifelor RON -> EUR. Aprilie 2026: ~5.0. Modifica daca cursul se schimba semnificativ.",
    "Cota impozit profit (SA = 16%)": "Total Hub este SA (Societate pe Actiuni) -> impozit profit 16% obligatoriu. Regimul micro 1-3% NU se aplica la SA.",
    "Amortizare CAPEX (ani)": "Durata in care echipamentele (ANPR, bariere, terminale) sunt amortizate contabil. Standard piata: 5 ani pentru hardware ANPR.",
    "Provizion CAPEX recurent (% anual)": "Procent anual din CAPEX rezervat pentru reparatii / inlocuiri partiale. Hardware ANPR uzura tipica: 6-10%/an. Aici e cost contabil, nu cash imediat.",
    "Inflatie OPEX si tarife (% anual)": "Inflatie aplicata anual pe OPEX si pe tarife (presupunand contract care permite indexare). Tipic 5-7% in Romania 2026.",
    "Orizont analiza (ani)": "Numarul de ani peste care se calculeaza NPV/IRR/Payback. Standard: 5 ani = durata tipica a unui contract de parcare retail.",
    "Rata discount (cost capital propriu)": "Rata de actualizare pentru NPV. Reflecta costul capitalului tau (alternative: bursa, imobiliare). Tipic 10-18%. Cu cat mai mare, cu atat NPV mai mic.",

    # --- Sheet "Parametri" — sectiunea B. RETAIL ---
    "CAPEX (EUR)": "Investitia totala initiala in echipamente: camere ANPR (recunoastere placa), bariere, terminal de plata, instalare. Tipic 25k-50k pentru parcari mici/medii, 50k-130k pentru retail mare/mall.",
    "OPEX lunar (EUR)": "Costuri operationale recurente: energie, internet, mentenanta hardware, licente software. Tipic 100-700 EUR/luna in functie de marime.",
    "Numar locuri parcare": "Numarul fizic de locuri de parcare. Influenteaza traficul potential maxim si CAPEX (cate bariere/senzori).",
    "Trafic intrari/zi": "Numarul mediu de intrari pe zi. Validare critica: verifica cu counter sau observatie 2 saptamani inainte de semnare contract. Asumptia gresita aici e cea mai frecventa cauza de proiect esuat.",
    "Tarif (RON/h)": "Tariful pe ora taxabila (dupa perioada de gratuitate). Standard piata: 5 RON/h pentru parcari mici, 10 RON/h pentru retail mare, 12-15 RON/h centre comerciale premium.",
    "Perioada gratuitate (min)": "Minute de stationare gratuita la inceput. Standard piata romaneasca retail: 120 min. PARAMETRUL CU CEL MAI MARE IMPACT pe NPV (vezi Tornado Retail). Sub 120 min = NPV pozitiv; peste = marginal/negativ.",

    # --- Sheet "Parametri" — distributie buckets ---
    "Distributie durate stationare retail": "Cum se distribuie clientii pe intervale de stationare. Suma ponderilor trebuie 100%. Default: 92% stau sub 60 min (gratuit la g=120), doar 8% genereaza venit.",
    "Bucket": "Interval de durata stationare (ex: 0-15 min, 15-30 min). Folosit pentru a calcula venitul mediu/intrare.",
    "Mid-point (min)": "Punctul mijlociu al intervalului in minute. Folosit la calculul venitului taxabil pe bucket = max(0, mid_point - gratuitate)/60 × tarif.",
    "Pondere (%)": "Procentul de clienti care stau in bucketul respectiv. Suma TOATE ponderile trebuie sa fie 100%. Aceasta distributie e calibrata pe date Lidl/Auchan retail Romania 2025.",
    "TOTAL pondere (trebuie 100%)": "Verificare: suma ponderilor pe buckets. Daca nu e 100%, calculele de venit sunt distorsionate proportional.",

    # --- Sheet "Parametri" — sectiunea C. NON-RETAIL ---
    "Tarif (RON/sesiune medie 2h)": "Tarif mediu pe sesiune (calculat la o durata tipica de stationare ~2h). Tipic 5-15 RON pentru parcari publice/captive. Daca tariful tau e RON/h, multiplica cu durata medie.",
    "Rata colectare (% platitori)": "Procent din intrari care platesc efectiv. 60-90% la parcari private cu ANPR + bariera fizica; 35-50% la parcari captive (hotel, birou) unde personalul plateste rar; sub 30% la parcari publice fara enforcement.",

    # --- Sheet "Calcul" — P&L ---
    "Venit operational": "Veniturile totale din parcari intr-un an, in EUR. Calcul: venit/intrare × trafic × 365 zile / curs EUR-RON. Inflatat anual cu rata de inflatie tarife.",
    "(-) OPEX direct": "Costurile operationale anuale (OPEX lunar × 12), inflatate anual. Energie + internet + mentenanta + software.",
    "EBITDA": "Earnings Before Interest, Taxes, Depreciation, Amortization. Profit operational brut = Venit - OPEX. Masoara cash-ul generat pur din operatiuni, fara efecte contabile sau finantare.",
    "(-) Provizion CAPEX (constant)": "Cost contabil anual pentru reparatii/inlocuiri viitoare echipamente. Reduce profitul impozabil, dar nu e cash imediat (banii raman in firma).",
    "(-) Amortizare (constant)": "Cost contabil care recupereaza CAPEX uniform pe ani (CAPEX / 5 ani). Reduce impozitul, dar NU este o iesire reala de cash. Doar contabilitate.",
    "Profit impozabil": "Baza de calcul pentru impozit pe profit. = EBITDA - Provizion - Amortizare. Daca e negativ, impozit = 0.",
    "(-) Impozit (16% pe pozitiv)": "Impozit pe profit 16% (SA Romania, regim standard). Se aplica DOAR pe profit pozitiv; pe pierderi contabile = 0.",
    "Profit net": "Profit dupa impozit. = Profit impozabil - Impozit. Cifra contabila finala.",
    "(+) Amortizare (non-cash)": "Readuga amortizarea pentru a calcula cash-ul real. Amortizarea redusese profit impozabil dar nu a iesit din cont -> o adaugi inapoi pentru cash flow real.",
    "CF NET (cash 100%)": "Cash flow real anual, presupunand finantare 100% din capital propriu (fara rate credit). = Profit net + Amortizare. Aceasta e cifra folosita pentru NPV/IRR.",
    "CF cumulat": "Suma cash-urilor de la an 0 (-CAPEX) pana la an N. Cand devine pozitiv pentru prima data = payback atins.",
    "INDICATORI": "Sectiunea de KPI-uri finale care se evalueaza pentru a decide GO/NO-GO pe investitie.",
    "NPV @ discount rate": "Net Present Value. Suma cash-urilor viitoare actualizate cu rata de discount, minus CAPEX. NPV > 0 = investitia bate alternativele tale. Tipic NPV > 50k EUR e bun pentru o parcare standalone.",
    "IRR": "Internal Rate of Return. Rata la care NPV ar fi exact 0 = randamentul real al investitiei. Reguli grosier: <7% = sub depozit (nebunesc), 12-20% = viabil, >20% = foarte bun.",
    "Payback simplu (ani)": "Numarul de ani pana cand CF cumulat devine pozitiv (recuperezi CAPEX). Sub 4 ani = bun; peste 5 ani = risc mare (ce se intampla daca contractul nu se prelungeste?).",
    "CF Net cumulat 5 ani": "Suma cash-urilor pe orizontul de analiza (5 ani), fara CAPEX scazut. Util pentru a vedea cat genereaza brut investitia.",
    "Verdict (bazat pe IRR)": "Clasificare automata pe baza IRR: NEBUNESC (<7% sub depozit), MARGINAL (7-12%), VIABIL (12-20%), FOARTE BUN (20-35%), EXCELENT (>35%).",
    "Helper IRR (an 0..5)": "Rand auxiliar cu fluxurile -CAPEX si CF1..CF5, folosit ca range de input pentru formula Excel =IRR(). Nu modifica.",

    # --- Sheets "Tornado Retail" / "Tornado Non-Retail" ---
    "Variabila": "Parametrul testat (CAPEX, OPEX, trafic, tarif, etc.). Se variaza unul cate unul, restul raman la baseline.",
    "Val low": "Valoarea pesimista a parametrului (~25% sub baseline pentru CAPEX/OPEX/trafic, sau scenariu defavorabil).",
    "Val baseline": "Valoarea de referinta (caz central). Toate celelalte randuri din matrice folosesc aceasta valoare.",
    "Val high": "Valoarea optimista a parametrului (~25% peste baseline, sau scenariu favorabil).",
    "NPV low": "NPV cand parametrul e setat la valoarea low (toate celelalte = baseline). EUR.",
    "NPV baseline": "NPV cu toti parametrii la baseline. Identic pe toate randurile (linie de referinta).",
    "NPV high": "NPV cand parametrul e setat la valoarea high (toate celelalte = baseline). EUR.",
    "Impact NPV": "|NPV high - NPV low|. Cu cat mai mare, cu atat parametrul muta NPV mai mult -> e mai sensibil. Sortat descrescator: variabila de sus = cea mai critica de validat in realitate.",

    # --- Sheet "Praguri" ---
    "Categorie / parametru": "Scenariu testat: la ce valoare minima a unui parametru investitia atinge IRR >= 15% (prag de viabilitate)?",
    "CAPEX 25k (mica)": "Pragul calculat asumand CAPEX 25.000 EUR (parcare mica) si OPEX 450 EUR/luna.",
    "CAPEX 45k (medie)": "Pragul calculat asumand CAPEX 45.000 EUR (parcare medie) si OPEX 600 EUR/luna.",
    "Note": "Asumptiile suplimentare folosite pentru calcul (tarif, distributie durate, etc.).",
}


def xlsx_tooltip(label: str) -> str:
    """Returneaza tooltip pentru o celula Excel pe baza label-ului ei. Empty daca nu exista."""
    return XLSX_TOOLTIPS.get(label.strip() if isinstance(label, str) else "", "")
