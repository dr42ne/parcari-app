"""
Aplicatie Streamlit interactiva pentru fezabilitate parcari individuale.
Run: streamlit run app.py
"""

import json
from datetime import datetime
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from parking_calc import ScenarioParams, simulate_scenario
from defaults import (
    RETAIL_250, NONRETAIL_250,
    auto_capex_from_locuri, auto_trafic_from_locuri, auto_opex_from_capex,
)
import tooltips


# ============================================================
# CONFIG
# ============================================================
st.set_page_config(
    page_title="Fezabilitate parcari — Total Hub",
    page_icon="P",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Editorial design — Total Hub SA
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

:root {
    --bg: #F4F0E6;
    --bg-card: #FFFFFF;
    --bg-soft: #FBF8F0;
    --text: #0F1419;
    --text-muted: #5A6171;
    --border: #E5DDD0;
    --border-strong: #C9BFAA;
    --accent: #7B2D26;
    --accent-amber: #A48B5E;
    --positive: #1F4D3F;
    --positive-bg: #E8EFEA;
    --neutral: #9A7B0A;
    --neutral-bg: #F4EDD9;
    --negative: #7B2D26;
    --negative-bg: #F1E5E2;
}

/* Background app */
.stApp { background-color: var(--bg) !important; }

/* Body text — DOAR pe elementele text, NU pe span (ar afecta icoane) */
.stApp {
    font-family: 'IBM Plex Sans', 'Aptos', 'Helvetica Neue', Arial, sans-serif;
    color: var(--text);
}
.stApp p, .stApp label, .stApp button,
.stApp .stMarkdown, .stApp [data-testid="stMarkdownContainer"] {
    font-family: 'IBM Plex Sans', 'Aptos', 'Helvetica Neue', Arial, sans-serif;
}

/* Pastreaza Material Symbols pentru icoane (chevroane, etc) */
[data-testid="stIconMaterial"],
.material-symbols-outlined,
.material-symbols-rounded,
.material-symbols-sharp,
.material-icons,
span[class*="material-symbols"],
span[class*="material-icons"] {
    font-family: 'Material Symbols Rounded', 'Material Symbols Outlined', 'Material Icons' !important;
}

/* Headings — IBM Plex Sans bold (sobru) */
h1, h2, h3, h4 {
    font-family: 'IBM Plex Sans', 'Aptos', 'Helvetica Neue', Arial, sans-serif !important;
    font-weight: 600 !important;
    letter-spacing: -0.01em !important;
    color: var(--text) !important;
}

/* Hero header custom */
.hero-eyebrow {
    font-family: 'IBM Plex Mono', 'SF Mono', Menlo, monospace;
    font-size: 0.7rem;
    letter-spacing: 0.22em;
    text-transform: uppercase;
    color: var(--accent);
    margin-bottom: 0.4rem;
    font-weight: 500;
}
.hero-title {
    font-family: 'IBM Plex Sans', 'Aptos', 'Helvetica Neue', Arial, sans-serif;
    font-size: 3rem;
    line-height: 1;
    font-weight: 600;
    letter-spacing: -0.025em;
    color: var(--text);
    margin: 0 0 0.5rem 0;
}
.hero-subtitle {
    font-family: 'IBM Plex Sans', 'Aptos', 'Helvetica Neue', Arial, sans-serif;
    font-size: 1rem;
    color: var(--text-muted);
    margin-bottom: 1.5rem;
    line-height: 1.5;
    max-width: 640px;
}
.hero-divider {
    height: 1px;
    background: var(--border-strong);
    margin: 1.5rem 0 2rem 0;
}

/* Verdict card — editorial */
.verdict-card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-left: 4px solid var(--text);
    padding: 2.5rem 2.5rem 2rem 2.5rem;
    margin-bottom: 2rem;
    position: relative;
    overflow: hidden;
}
.verdict-card.profitabil { border-left-color: var(--positive); background: linear-gradient(135deg, var(--bg-card) 0%, var(--positive-bg) 100%); }
.verdict-card.marginal { border-left-color: var(--neutral); background: linear-gradient(135deg, var(--bg-card) 0%, var(--neutral-bg) 100%); }
.verdict-card.neprofitabil { border-left-color: var(--negative); background: linear-gradient(135deg, var(--bg-card) 0%, var(--negative-bg) 100%); }

.verdict-eyebrow {
    font-family: 'IBM Plex Mono', 'SF Mono', Menlo, monospace;
    font-size: 0.7rem;
    letter-spacing: 0.22em;
    text-transform: uppercase;
    color: var(--text-muted);
    margin-bottom: 0.6rem;
    font-weight: 500;
}
.verdict-headline {
    font-family: 'IBM Plex Sans', 'Aptos', 'Helvetica Neue', Arial, sans-serif;
    font-size: 4rem;
    line-height: 0.95;
    font-weight: 600;
    letter-spacing: -0.03em;
    margin: 0 0 0.75rem 0;
}
.verdict-card.profitabil .verdict-headline { color: var(--positive); }
.verdict-card.marginal .verdict-headline { color: var(--neutral); }
.verdict-card.neprofitabil .verdict-headline { color: var(--negative); }

.verdict-subtitle {
    font-family: 'IBM Plex Sans', 'Aptos', 'Helvetica Neue', Arial, sans-serif;
    font-size: 1rem;
    color: var(--text);
    line-height: 1.55;
    max-width: 720px;
    font-weight: 400;
}
.verdict-subtitle em {
    font-style: italic;
    font-family: 'IBM Plex Sans', 'Aptos', 'Helvetica Neue', Arial, sans-serif;
    color: var(--text-muted);
}

/* KPI cards custom */
.kpi-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 1rem;
    margin-bottom: 2rem;
}
.kpi-card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    padding: 1.5rem 1.25rem;
    position: relative;
    transition: border-color 0.2s;
}
.kpi-card:hover { border-color: var(--border-strong); }
.kpi-label {
    font-family: 'IBM Plex Mono', 'SF Mono', Menlo, monospace;
    font-size: 0.65rem;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    color: var(--text-muted);
    font-weight: 500;
    margin-bottom: 0.75rem;
}
.kpi-value {
    font-family: 'IBM Plex Sans', 'Aptos', 'Helvetica Neue', Arial, sans-serif;
    font-size: 2.1rem;
    line-height: 1;
    font-weight: 600;
    letter-spacing: -0.025em;
    font-variant-numeric: tabular-nums;
    color: var(--text);
    margin-bottom: 0.4rem;
}
.kpi-value.positive { color: var(--positive); }
.kpi-value.negative { color: var(--negative); }
.kpi-value.neutral { color: var(--neutral); }
.kpi-unit {
    font-family: 'IBM Plex Sans', 'Aptos', 'Helvetica Neue', Arial, sans-serif;
    font-size: 0.75rem;
    color: var(--text-muted);
    margin-left: 0.3rem;
    font-weight: 400;
}
.kpi-context {
    font-family: 'IBM Plex Mono', 'SF Mono', Menlo, monospace;
    font-size: 0.65rem;
    color: var(--text-muted);
    letter-spacing: 0.05em;
    margin-top: 0.3rem;
}

/* Section titles */
.section-title {
    font-family: 'IBM Plex Sans', 'Aptos', 'Helvetica Neue', Arial, sans-serif;
    font-size: 1.5rem;
    font-weight: 600;
    color: var(--text);
    margin: 2rem 0 1rem 0;
    letter-spacing: -0.015em;
}
.section-eyebrow {
    font-family: 'IBM Plex Mono', 'SF Mono', Menlo, monospace;
    font-size: 0.65rem;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    color: var(--accent);
    margin-bottom: 0.3rem;
    font-weight: 500;
}

/* Sidebar styling */
[data-testid="stSidebar"] {
    background-color: var(--bg-soft) !important;
    border-right: 1px solid var(--border);
}
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2 {
    font-family: 'IBM Plex Sans', 'Aptos', 'Helvetica Neue', Arial, sans-serif !important;
    font-size: 1.25rem !important;
}

/* Streamlit native st.metric (fallback if used) */
[data-testid="stMetricValue"] {
    font-family: 'IBM Plex Sans', 'Aptos', 'Helvetica Neue', Arial, sans-serif !important;
    font-variant-numeric: tabular-nums;
}
[data-testid="stMetricLabel"] {
    font-family: 'IBM Plex Mono', 'SF Mono', Menlo, monospace !important;
    font-size: 0.7rem !important;
    text-transform: uppercase;
    letter-spacing: 0.15em !important;
}

/* Inputs styling */
.stNumberInput input, .stTextInput input {
    font-family: 'IBM Plex Mono', 'SF Mono', Menlo, monospace !important;
    font-variant-numeric: tabular-nums;
    border-radius: 2px !important;
}
.stSlider > div > div > div { border-radius: 2px; }

/* Buttons */
.stButton > button {
    font-family: 'IBM Plex Sans', 'Aptos', 'Helvetica Neue', Arial, sans-serif !important;
    font-weight: 500 !important;
    letter-spacing: 0.02em;
    border-radius: 2px !important;
    border: 1px solid var(--border-strong) !important;
    background: var(--bg-card) !important;
    color: var(--text) !important;
    transition: all 0.2s;
}
.stButton > button:hover {
    border-color: var(--accent) !important;
    color: var(--accent) !important;
}
.stDownloadButton > button {
    font-family: 'IBM Plex Sans', 'Aptos', 'Helvetica Neue', Arial, sans-serif !important;
    border-radius: 2px !important;
}

/* Dataframe styling */
[data-testid="stDataFrame"] {
    font-family: 'IBM Plex Sans', 'Aptos', 'Helvetica Neue', Arial, sans-serif;
}
[data-testid="stDataFrame"] table {
    font-variant-numeric: tabular-nums;
}

/* Expander */
[data-testid="stExpander"] {
    background: var(--bg-card);
    border: 1px solid var(--border) !important;
    border-radius: 2px !important;
}

/* Footer */
.app-footer {
    margin-top: 3rem;
    padding-top: 1.5rem;
    border-top: 1px solid var(--border-strong);
    font-family: 'IBM Plex Mono', 'SF Mono', Menlo, monospace;
    font-size: 0.7rem;
    letter-spacing: 0.1em;
    color: var(--text-muted);
    text-transform: uppercase;
}

/* Hide Streamlit chrome */
#MainMenu { visibility: hidden; }
header[data-testid="stHeader"] { background: transparent; }
.stDeployButton { display: none; }

/* Custom badge */
.custom-badge {
    display: inline-block;
    font-family: 'IBM Plex Mono', 'SF Mono', Menlo, monospace;
    font-size: 0.6rem;
    letter-spacing: 0.1em;
    color: var(--accent);
    padding: 0.1rem 0.3rem;
    border: 1px solid var(--accent);
    margin-left: 0.3rem;
    text-transform: uppercase;
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# SESSION STATE INITIALIZATION
# ============================================================
def init_session():
    """Initializeaza session_state cu defaults retail daca nu e setat."""
    if "_initialized" not in st.session_state:
        for k, v in RETAIL_250.items():
            st.session_state[k] = v
        st.session_state["_user_overrides"] = set()
        st.session_state["_initialized"] = True


init_session()


# ============================================================
# AUTO-ADJUSTMENT LOGIC
# Se executa o singura data, INAINTE de render widget-uri.
# Citeste session_state si calculeaza derivate, scrie doar daca cheia NU e in _user_overrides.
# ============================================================
def apply_auto_adjustments():
    overrides = st.session_state.get("_user_overrides", set())
    locuri = st.session_state.get("numar_locuri", 250)
    tip = st.session_state.get("tip_parcare", "RETAIL")

    # 1. trafic_zilnic = f(locuri, tip) — daca user nu a override
    if "trafic_zilnic" not in overrides:
        st.session_state["trafic_zilnic"] = float(auto_trafic_from_locuri(locuri, tip))

    # 2. capex_eur = f(locuri, tip) — daca user nu a override
    if "capex_eur" not in overrides:
        st.session_state["capex_eur"] = float(auto_capex_from_locuri(locuri, tip))

    # 3. opex_lunar = f(capex_eur) — daca user nu a override
    if "opex_lunar_eur" not in overrides:
        st.session_state["opex_lunar_eur"] = float(auto_opex_from_capex(st.session_state["capex_eur"]))


def mark_override(key: str):
    """Callback la modificare manuala — marcheaza cheia ca user override."""
    st.session_state["_user_overrides"].add(key)


def reset_override(key: str):
    """Sterge override pentru o cheie (revine la auto-suggested)."""
    st.session_state["_user_overrides"].discard(key)


def is_overridden(key: str) -> bool:
    return key in st.session_state.get("_user_overrides", set())


# ============================================================
# DUAL INPUT WIDGET (number_input + slider sincronizat via callbacks)
# ============================================================
def _sync_widgets_to_canonical(key: str, num_key: str, sld_key: str, factor: float = 1.0):
    """Sincronizeaza widget keys cu valoarea canonica (apelat inainte de render)."""
    canonical = st.session_state.get(key)
    if canonical is None:
        return
    target = canonical * factor
    st.session_state[num_key] = target
    st.session_state[sld_key] = target


def dual_input(label: str, key: str, min_v: float, max_v: float, step: float,
               help_text: str = "", fmt: str = "%.2f", show_reset: bool = True,
               is_int: bool = False):
    """
    Number_input + slider sincronizate. Modificarea oricaruia:
      - actualizeaza celalalt widget
      - actualizeaza cheia canonica
      - marcheaza key in _user_overrides
    """
    num_key = f"_num_{key}"
    sld_key = f"_sld_{key}"

    # Sincronizare widget state cu canonical (la fiecare rerun, ca sa prinda auto-adjustments)
    canonical = st.session_state.get(key, min_v)
    canonical = max(min_v, min(max_v, canonical))
    st.session_state[num_key] = canonical
    st.session_state[sld_key] = canonical

    def on_num_change():
        v = st.session_state[num_key]
        st.session_state[sld_key] = v
        st.session_state[key] = int(round(v)) if is_int else float(v)
        st.session_state["_user_overrides"].add(key)

    def on_sld_change():
        v = st.session_state[sld_key]
        st.session_state[num_key] = v
        st.session_state[key] = int(round(v)) if is_int else float(v)
        st.session_state["_user_overrides"].add(key)

    custom = is_overridden(key)
    label_full = f"{label} (custom)" if custom and show_reset else label

    c1, c2 = st.columns([1, 2])
    with c1:
        st.number_input(
            label_full,
            min_value=float(min_v), max_value=float(max_v),
            step=float(step), key=num_key, help=help_text, format=fmt,
            on_change=on_num_change,
        )
    with c2:
        st.slider(
            " ", min_value=float(min_v), max_value=float(max_v),
            step=float(step), key=sld_key, label_visibility="collapsed",
            on_change=on_sld_change,
        )

    if custom and show_reset:
        if st.button(f"Reset la sugestie", key=f"_rst_{key}"):
            reset_override(key)
            st.rerun()

    return st.session_state[key]


def dual_input_int(label, key, min_v, max_v, step, help_text="", show_reset=True):
    """Number_input + slider int (toate valorile int, fara warnings)."""
    num_key = f"_num_{key}"
    sld_key = f"_sld_{key}"

    canonical = int(round(st.session_state.get(key, min_v)))
    canonical = max(int(min_v), min(int(max_v), canonical))
    st.session_state[num_key] = canonical
    st.session_state[sld_key] = canonical

    def on_num_change():
        v = int(st.session_state[num_key])
        st.session_state[sld_key] = v
        st.session_state[key] = v
        st.session_state["_user_overrides"].add(key)

    def on_sld_change():
        v = int(st.session_state[sld_key])
        st.session_state[num_key] = v
        st.session_state[key] = v
        st.session_state["_user_overrides"].add(key)

    custom = is_overridden(key)
    label_full = f"{label} (custom)" if custom and show_reset else label

    c1, c2 = st.columns([1, 2])
    with c1:
        st.number_input(
            label_full,
            min_value=int(min_v), max_value=int(max_v),
            step=int(step), key=num_key, help=help_text, format="%d",
            on_change=on_num_change,
        )
    with c2:
        st.slider(
            " ", min_value=int(min_v), max_value=int(max_v),
            step=int(step), key=sld_key, label_visibility="collapsed",
            on_change=on_sld_change,
        )

    if custom and show_reset:
        if st.button(f"Reset la sugestie", key=f"_rst_{key}"):
            reset_override(key)
            st.rerun()

    return st.session_state[key]


def dual_input_pct(label, key, min_v, max_v, step, help_text="", show_reset=True):
    """Procent: afiseaza in % dar canonical e decimal."""
    num_key = f"_num_{key}"
    sld_key = f"_sld_{key}"

    canonical = st.session_state.get(key, min_v)
    canonical = max(min_v, min(max_v, canonical))
    canonical_pct = canonical * 100
    st.session_state[num_key] = canonical_pct
    st.session_state[sld_key] = canonical_pct

    def on_num_change():
        v_pct = st.session_state[num_key]
        st.session_state[sld_key] = v_pct
        st.session_state[key] = v_pct / 100.0
        st.session_state["_user_overrides"].add(key)

    def on_sld_change():
        v_pct = st.session_state[sld_key]
        st.session_state[num_key] = v_pct
        st.session_state[key] = v_pct / 100.0
        st.session_state["_user_overrides"].add(key)

    custom = is_overridden(key)
    label_full = f"{label} (custom)" if custom and show_reset else label

    c1, c2 = st.columns([1, 2])
    with c1:
        st.number_input(
            label_full,
            min_value=float(min_v * 100), max_value=float(max_v * 100),
            step=float(step * 100), key=num_key, help=help_text, format="%.1f",
            on_change=on_num_change,
        )
    with c2:
        st.slider(
            " ", min_value=float(min_v * 100), max_value=float(max_v * 100),
            step=float(step * 100), key=sld_key, label_visibility="collapsed",
            format="%.1f%%",
            on_change=on_sld_change,
        )

    if custom and show_reset:
        if st.button(f"Reset la sugestie", key=f"_rst_{key}"):
            reset_override(key)
            st.rerun()

    return st.session_state[key]


# ============================================================
# HERO HEADER — editorial
# ============================================================
st.markdown("""
<div class="hero-eyebrow">Total Hub SA · Studio de modelare financiara</div>
<h1 class="hero-title">Analiza fezabilitate parcare</h1>
<p class="hero-subtitle">Modelare interactiva pentru o singura parcare. Modifica parametrii in panoul lateral si urmareste cum se schimba verdictul, NPV, IRR si payback in timp real.</p>
<div class="hero-divider"></div>
""", unsafe_allow_html=True)

col_a, col_b, col_c, col_d = st.columns([2, 2, 2, 2])

with col_a:
    new_tip = st.radio(
        "Tip parcare",
        ["RETAIL", "NON-RETAIL"],
        index=0 if st.session_state.get("tip_parcare", "RETAIL") == "RETAIL" else 1,
        horizontal=True,
        key="_tip_radio",
        help="RETAIL: cu perioada de gratuitate (Lidl/Kaufland/Auchan). NON-RETAIL: fara gratuitate, plata sesiune.",
    )
    if new_tip != st.session_state.get("tip_parcare"):
        st.session_state["tip_parcare"] = new_tip
        # Reset overrides cand schimbam tipul (defaults trebuie sa se aplice)
        st.session_state["_user_overrides"] = set()
        st.rerun()

with col_b:
    if st.button("Reseteaza la defaults", help="Toate parametrii revin la valorile initiale (250 locuri)"):
        defaults = RETAIL_250 if st.session_state["tip_parcare"] == "RETAIL" else NONRETAIL_250
        for k, v in defaults.items():
            st.session_state[k] = v
        st.session_state["_user_overrides"] = set()
        st.rerun()

with col_c:
    # Save scenariu
    scenariu_name = st.text_input("Nume scenariu", value="Scenariu_1", key="_sc_name")
    state_to_save = {
        "schema_version": "1.0",
        "name": scenariu_name,
        "created_at": datetime.now().isoformat(),
        "tip_parcare": st.session_state["tip_parcare"],
        "params": {k: st.session_state.get(k) for k in RETAIL_250.keys()},
        "user_overrides": list(st.session_state.get("_user_overrides", [])),
    }
    st.download_button(
        label="Salveaza scenariu (JSON)",
        data=json.dumps(state_to_save, indent=2, ensure_ascii=False).encode("utf-8"),
        file_name=f"{scenariu_name}.json",
        mime="application/json",
    )

with col_d:
    uploaded = st.file_uploader("Incarca scenariu (JSON)", type=["json"], key="_upload")
    if uploaded is not None:
        try:
            data = json.load(uploaded)
            if data.get("schema_version") != "1.0":
                st.error("Versiune schema necunoscuta")
            else:
                for k, v in data.get("params", {}).items():
                    st.session_state[k] = v
                st.session_state["tip_parcare"] = data.get("tip_parcare", "RETAIL")
                st.session_state["_user_overrides"] = set(data.get("user_overrides", []))
                st.success(f"Incarcat: {data.get('name', 'fara nume')}")
                st.rerun()
        except Exception as e:
            st.error(f"Eroare la incarcare: {e}")


# Aplica auto-adjustments INAINTE de render widget-uri sidebar
apply_auto_adjustments()


# ============================================================
# SIDEBAR — parametri grupati
# ============================================================
with st.sidebar:
    st.header("Parametri scenariu")
    tip_parcare = st.session_state["tip_parcare"]

    with st.expander("INVESTITIE", expanded=True):
        tip_inv = st.radio(
            "Tip finantare",
            ["capital_propriu", "credit", "mix"],
            index=["capital_propriu", "credit", "mix"].index(st.session_state.get("tip_investitie", "mix")),
            format_func=lambda x: {"capital_propriu": "Capital propriu (100%)", "credit": "Credit (100%)", "mix": "Mix credit + capital"}[x],
            key="_tip_inv_radio",
            help=tooltips.get("tip_investitie"),
        )
        st.session_state["tip_investitie"] = tip_inv
        # Setam pondere_credit pe baza tipului
        if tip_inv == "capital_propriu":
            st.session_state["pondere_credit"] = 0.0
        elif tip_inv == "credit":
            st.session_state["pondere_credit"] = 1.0

        dual_input_int(
            "CAPEX (EUR)", "capex_eur",
            min_v=10000, max_v=200000, step=1000,
            help_text=tooltips.get("capex_eur"),
        )

        if tip_inv == "mix":
            dual_input_pct(
                "Pondere credit", "pondere_credit",
                min_v=0.0, max_v=1.0, step=0.05,
                help_text=tooltips.get("pondere_credit"),
                show_reset=False,
            )

        if tip_inv != "capital_propriu":
            dual_input_int(
                "Durata credit (ani)", "durata_credit_ani",
                min_v=1, max_v=10, step=1,
                help_text=tooltips.get("durata_credit_ani"),
                show_reset=False,
            )
            dual_input_pct(
                "Rata dobanda", "rata_dobanda",
                min_v=0.0, max_v=0.20, step=0.005,
                help_text=tooltips.get("rata_dobanda"),
                show_reset=False,
            )

    with st.expander("OPERARE", expanded=True):
        dual_input_int(
            "Numar locuri parcare", "numar_locuri",
            min_v=20, max_v=3000, step=10,
            help_text=tooltips.get("numar_locuri"),
            show_reset=False,
        )
        dual_input_int(
            "Trafic zilnic (intrari/zi)", "trafic_zilnic",
            min_v=20, max_v=20000, step=25,
            help_text=tooltips.get("trafic_zilnic"),
        )
        dual_input_int(
            "OPEX lunar (EUR)", "opex_lunar_eur",
            min_v=50, max_v=3000, step=10,
            help_text=tooltips.get("opex_lunar_eur"),
        )
        dual_input_int(
            "Durata contract (ani)", "durata_contract_ani",
            min_v=1, max_v=10, step=1,
            help_text=tooltips.get("durata_contract_ani"),
            show_reset=False,
        )
        dual_input_int(
            "Overhead anual (EUR)", "overhead_anual_eur",
            min_v=0, max_v=200000, step=500,
            help_text=tooltips.get("overhead_anual_eur"),
            show_reset=False,
        )

    with st.expander("TARIFE", expanded=True):
        if tip_parcare == "RETAIL":
            dual_input_int(
                "Perioada gratuitate (min)", "gratuitate_min",
                min_v=0, max_v=240, step=15,
                help_text=tooltips.get("gratuitate_min"),
                show_reset=False,
            )
            dual_input(
                "Tarif ora 1 (RON/h)", "tarif_ora_1_ron",
                min_v=0.0, max_v=30.0, step=0.5,
                help_text=tooltips.get("tarif_ora_1_ron"),
                fmt="%.1f",
                show_reset=False,
            )
            dual_input(
                "Tarif ora 2 (RON/h)", "tarif_ora_2_ron",
                min_v=0.0, max_v=30.0, step=0.5,
                help_text=tooltips.get("tarif_ora_2_ron"),
                fmt="%.1f",
                show_reset=False,
            )
            dual_input(
                "Tarif ora 3+ (RON/h)", "tarif_ora_3plus_ron",
                min_v=0.0, max_v=30.0, step=0.5,
                help_text=tooltips.get("tarif_ora_3plus_ron"),
                fmt="%.1f",
                show_reset=False,
            )
        else:  # NON-RETAIL
            dual_input(
                "Tarif sesiune (RON)", "tarif_sesiune_ron",
                min_v=0.0, max_v=50.0, step=0.5,
                help_text=tooltips.get("tarif_sesiune_ron"),
                fmt="%.1f",
                show_reset=False,
            )
            dual_input_pct(
                "Rata colectare", "rata_colectare",
                min_v=0.10, max_v=1.00, step=0.05,
                help_text=tooltips.get("rata_colectare"),
                show_reset=False,
            )

    with st.expander("FISCAL / MACRO", expanded=False):
        dual_input_pct(
            "Cota impozit profit", "cota_impozit",
            min_v=0.0, max_v=0.30, step=0.01,
            help_text=tooltips.get("cota_impozit"),
            show_reset=False,
        )
        dual_input_pct(
            "Inflatie OPEX", "inflatie_opex",
            min_v=0.0, max_v=0.20, step=0.005,
            help_text=tooltips.get("inflatie_opex"),
            show_reset=False,
        )
        dual_input_pct(
            "Inflatie tarife", "inflatie_tarife",
            min_v=0.0, max_v=0.20, step=0.005,
            help_text=tooltips.get("inflatie_tarife"),
            show_reset=False,
        )
        dual_input_pct(
            "Discount rate (cost capital propriu)", "discount_rate",
            min_v=0.05, max_v=0.30, step=0.01,
            help_text=tooltips.get("discount_rate"),
            show_reset=False,
        )
        dual_input(
            "Curs EUR/RON", "eur_ron",
            min_v=4.5, max_v=6.0, step=0.05,
            help_text=tooltips.get("eur_ron"),
            fmt="%.2f",
            show_reset=False,
        )
        dual_input_pct(
            "Provizion CAPEX (% anual)", "provizion_pct",
            min_v=0.0, max_v=0.15, step=0.005,
            help_text=tooltips.get("provizion_pct"),
            show_reset=False,
        )
        dual_input_int(
            "Amortizare CAPEX (ani)", "amortizare_ani",
            min_v=3, max_v=10, step=1,
            help_text=tooltips.get("amortizare_ani"),
            show_reset=False,
        )

    with st.expander("SUPLIMENTARE", expanded=False):
        dual_input_pct(
            "Ramp-up trafic anul 1", "rampup_an1",
            min_v=0.20, max_v=1.0, step=0.05,
            help_text=tooltips.get("rampup_an1"),
            show_reset=False,
        )
        dual_input_int(
            "Asigurare anuala (EUR)", "asigurare_anuala_eur",
            min_v=0, max_v=5000, step=100,
            help_text=tooltips.get("asigurare_anuala_eur"),
            show_reset=False,
        )
        dual_input_int(
            "Marketing initial (EUR)", "marketing_initial_eur",
            min_v=0, max_v=20000, step=500,
            help_text=tooltips.get("marketing_initial_eur"),
            show_reset=False,
        )
        dual_input_int(
            "Marketing lunar (EUR)", "marketing_lunar_eur",
            min_v=0, max_v=2000, step=50,
            help_text=tooltips.get("marketing_lunar_eur"),
            show_reset=False,
        )


# ============================================================
# SIMULARE
# ============================================================
params = ScenarioParams(
    tip_parcare=st.session_state["tip_parcare"],
    tip_investitie=st.session_state["tip_investitie"],
    capex_eur=float(st.session_state["capex_eur"]),
    pondere_credit=float(st.session_state["pondere_credit"]),
    durata_credit_ani=int(st.session_state["durata_credit_ani"]),
    rata_dobanda=float(st.session_state["rata_dobanda"]),
    durata_contract_ani=int(st.session_state["durata_contract_ani"]),
    overhead_anual_eur=float(st.session_state["overhead_anual_eur"]),
    numar_locuri=int(st.session_state["numar_locuri"]),
    trafic_zilnic=float(st.session_state["trafic_zilnic"]),
    opex_lunar_eur=float(st.session_state["opex_lunar_eur"]),
    gratuitate_min=float(st.session_state.get("gratuitate_min", 120)),
    tarif_ora_1_ron=float(st.session_state.get("tarif_ora_1_ron", 5)),
    tarif_ora_2_ron=float(st.session_state.get("tarif_ora_2_ron", 7)),
    tarif_ora_3plus_ron=float(st.session_state.get("tarif_ora_3plus_ron", 10)),
    tarif_sesiune_ron=float(st.session_state.get("tarif_sesiune_ron", 10)),
    rata_colectare=float(st.session_state.get("rata_colectare", 0.75)),
    cota_impozit=float(st.session_state["cota_impozit"]),
    inflatie_opex=float(st.session_state["inflatie_opex"]),
    inflatie_tarife=float(st.session_state["inflatie_tarife"]),
    discount_rate=float(st.session_state["discount_rate"]),
    eur_ron=float(st.session_state["eur_ron"]),
    provizion_pct=float(st.session_state["provizion_pct"]),
    amortizare_ani=int(st.session_state["amortizare_ani"]),
    rampup_an1=float(st.session_state["rampup_an1"]),
    asigurare_anuala_eur=float(st.session_state["asigurare_anuala_eur"]),
    marketing_initial_eur=float(st.session_state["marketing_initial_eur"]),
    marketing_lunar_eur=float(st.session_state["marketing_lunar_eur"]),
)
result = simulate_scenario(params)


# ============================================================
# MAIN AREA — verdict editorial, KPI, chart, tabel P&L
# ============================================================
verdict = result["verdict"]
verdict_class = verdict.lower()

# Subtitle text per verdict — editorial, contextual
verdict_subtitle = {
    "PROFITABIL": (
        "Proiectul <em>creeaza valoare in plus</em> fata de costul oportunitatii. "
        "Fluxurile actualizate depasesc capitalul investit, iar randamentul intrec pragurile minimale "
        "de viabilitate. Investitia se justifica financiar."
    ),
    "MARGINAL": (
        "Proiectul <em>abia atinge pragul de viabilitate</em>. Robustetea este fragila: "
        "schimbari mici in trafic, gratuitate sau tarife pot rasturna verdictul. "
        "Recomandare: validare suplimentara a ipotezelor inainte de decizie."
    ),
    "NEPROFITABIL": (
        "Proiectul <em>nu acopera costul oportunitatii capitalului</em>. "
        "Banii pusi aici aduc mai putin decat alternativele de investitie. "
        "Modifica parametrii (gratuitate, trafic, tarife) pentru a explora conditiile in care ar deveni viabil."
    ),
}[verdict]

st.markdown(f"""
<div class="verdict-card {verdict_class}">
    <div class="verdict-eyebrow">Verdict financiar</div>
    <div class="verdict-headline">{verdict}</div>
    <div class="verdict-subtitle">{verdict_subtitle}</div>
</div>
""", unsafe_allow_html=True)


# ============================================================
# KPI cards — custom HTML
# ============================================================
def fmt_eur(v):
    return f"{v:,.0f}".replace(",", ".")

npv_v = result["npv"]
npv_class = "positive" if npv_v > 0 else ("negative" if npv_v < 0 else "neutral")

irr_v = result["irr"]
if irr_v is None:
    irr_str = "N/A"
    irr_class = "neutral"
else:
    irr_str = f"{irr_v*100:.1f}%"
    irr_class = "positive" if irr_v > 0.15 else ("negative" if irr_v < 0.07 else "neutral")

pb = result["payback"]
if pb is None:
    pb_str = "> orizont"
    pb_class = "negative"
else:
    pb_str = f"{pb:.1f} <span class='kpi-unit'>ani</span>"
    pb_class = "positive" if pb < 4 else ("negative" if pb > 6 else "neutral")

cf_cum = result["cf_cumulat_5"]
cf_class = "positive" if cf_cum > 0 else "negative"

st.markdown(f"""
<div class="kpi-grid">
    <div class="kpi-card">
        <div class="kpi-label">NPV @ Discount</div>
        <div class="kpi-value {npv_class}">{fmt_eur(npv_v)}<span class="kpi-unit">EUR</span></div>
        <div class="kpi-context">Valoare actualizata neta</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">IRR</div>
        <div class="kpi-value {irr_class}">{irr_str}</div>
        <div class="kpi-context">Rata interna de rentabilitate</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">Payback</div>
        <div class="kpi-value {pb_class}">{pb_str}</div>
        <div class="kpi-context">Recuperare capital propriu</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-label">CF Cumulat {params.durata_contract_ani} Ani</div>
        <div class="kpi-value {cf_class}">{fmt_eur(cf_cum)}<span class="kpi-unit">EUR</span></div>
        <div class="kpi-context">Numerar net total</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ============================================================
# Chart Plotly — editorial palette
# ============================================================
st.markdown(f"""
<div class="section-eyebrow">Fluxuri de numerar</div>
<h2 class="section-title">Cash flow net · {params.durata_contract_ani} ani</h2>
""", unsafe_allow_html=True)

ani_labels = [f"An {i}" for i in range(1, params.durata_contract_ani + 1)]
fig = make_subplots(specs=[[{"secondary_y": True}]])
fig.add_trace(
    go.Bar(
        x=ani_labels, y=result["cash_flows"],
        name="CF anual",
        marker_color=["#1F4D3F" if cf >= 0 else "#7B2D26" for cf in result["cash_flows"]],
        marker_line_width=0,
    ),
    secondary_y=False,
)
fig.add_trace(
    go.Scatter(
        x=ani_labels, y=result["cf_cumulat"],
        name="CF cumulat",
        mode="lines+markers",
        line=dict(color="#0F1419", width=2.5),
        marker=dict(size=8, color="#A48B5E", line=dict(color="#0F1419", width=1.5)),
    ),
    secondary_y=True,
)
fig.update_xaxes(
    title_text="",
    tickfont=dict(family="IBM Plex Mono, SF Mono, monospace", size=11, color="#5A6171"),
    showgrid=False,
    showline=True,
    linecolor="#C9BFAA",
    linewidth=1,
)
fig.update_yaxes(
    title_text="CF anual (EUR)",
    title_font=dict(family="IBM Plex Sans, Aptos, Helvetica Neue, Arial, sans-serif", size=11, color="#5A6171"),
    tickfont=dict(family="IBM Plex Mono, SF Mono, monospace", size=10, color="#5A6171"),
    gridcolor="#EBE5D5",
    zerolinecolor="#C9BFAA",
    zerolinewidth=1,
    secondary_y=False,
)
fig.update_yaxes(
    title_text="CF cumulat (EUR)",
    title_font=dict(family="IBM Plex Sans, Aptos, Helvetica Neue, Arial, sans-serif", size=11, color="#5A6171"),
    tickfont=dict(family="IBM Plex Mono, SF Mono, monospace", size=10, color="#5A6171"),
    showgrid=False,
    secondary_y=True,
)
fig.update_layout(
    height=380,
    hovermode="x unified",
    plot_bgcolor="#FFFFFF",
    paper_bgcolor="#F4F0E6",
    font=dict(family="IBM Plex Sans, Aptos, Helvetica Neue, Arial, sans-serif", color="#0F1419"),
    legend=dict(
        orientation="h", y=-0.18, x=0,
        font=dict(family="IBM Plex Mono, SF Mono, monospace", size=10),
        bgcolor="rgba(0,0,0,0)",
    ),
    margin=dict(l=60, r=60, t=20, b=60),
)
st.plotly_chart(fig, use_container_width=True)

# Tabel P&L
st.markdown(f"""
<div class="section-eyebrow">Profit & Loss</div>
<h2 class="section-title">P&L detaliat · {params.durata_contract_ani} ani</h2>
""", unsafe_allow_html=True)
df = pd.DataFrame(result["pnl"])
df_display = df[["an", "venit", "opex", "ebitda", "provizion", "overhead",
                  "ebit", "amortizare", "profit_impozabil", "impozit",
                  "profit_net", "anuitate", "cf_net"]].copy()
df_display.columns = ["An", "Venit", "OPEX", "EBITDA", "Provizion", "Overhead",
                       "EBIT", "Amortizare", "Profit imp.", "Impozit",
                       "Profit net", "Anuitate credit", "CF Net"]
st.dataframe(
    df_display.style.format({
        c: "{:,.0f} EUR" for c in df_display.columns if c != "An"
    }),
    use_container_width=True,
    hide_index=True,
)

# Sectiune ipoteze
with st.expander("Ipoteze si formule (referinta)", expanded=False):
    st.markdown(f"""
**Capital propriu investit:** {result['capex_propriu_total']:,.0f} EUR
(CAPEX × (1 - pondere_credit) + marketing initial)

**Anuitate anuala credit:** {result['anuitate']:,.0f} EUR
(plata anuala pentru creditul echipament pe `durata_credit_ani` ani la `rata_dobanda`)

**Venit baseline anul 1 (la trafic full, fara ramp-up):** {result['venit_y1_full']:,.0f} EUR

**Verdict thresholds:**
- PROFITABIL: NPV > 0 SI IRR > 15% SI Payback < 4 ani
- MARGINAL: NPV > 0 SAU IRR > 10%
- NEPROFITABIL: restul

**Reguli auto-ajustare** (cand modifici un parametru, se actualizeaza automat — daca nu ai override manual):
- Numar locuri → trafic baseline (rotatie 5/zi retail, 2.5/zi non-retail)
- Numar locuri + tip → CAPEX (15.000 + 100 × locuri retail; 12.000 + 80 × locuri non-retail)
- CAPEX → OPEX lunar (100 + CAPEX × 1,2%/12)

Marcheaza un parametru ca "(custom)" daca l-ai modificat manual — auto-update se pauzeaza pentru el. Foloseste butonul "Reset ... la sugestie" pentru a reactiva.

**Tarif progresiv retail:** dupa perioada de gratuitate, primele 60 min taxabile la tarif ora 1, urmatoarele 60 min la tarif ora 2, restul la tarif ora 3+. Calcul prorata pe minute.

**Sursa formulelor:** `parking_calc.py` (verificat numeric vs `outputs/validator.py`, 16 teste de regresie PASS).
""")

st.caption("Total Hub SA — analiza fezabilitate parcari individuale | datele sunt estimari, validare reala obligatorie inainte de semnare contract.")
