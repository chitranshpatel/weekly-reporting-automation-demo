
from pathlib import Path
import base64
import html
import os

import pandas as pd
import streamlit as st

from reporting import (
    calculate_kpis,
    generate_management_summary,
    generate_powerpoint,
    load_management_prompt,
    load_history,
)

st.set_page_config(
    page_title="Weekly Reporting Automation Demo",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Wesfarmers Health-inspired executive dashboard styling.
st.markdown(
    """
    <style>
      :root {
        --wh-deep: #004838;
        --wh-green: #007E48;
        --wh-mint: #E7F2EA;
        --wh-pale: #F5F9F6;
        --wh-ink: #17352D;
        --wh-muted: #111111;
        --wh-border: #D8E5DD;
        --wh-white: #FFFFFF;
      }
      html, body, [class*="css"], .stApp {
        font-family: "Figtree", "Aptos", "Segoe UI", sans-serif;
      }
      .stApp {background: linear-gradient(180deg, #F4F8F5 0%, #FFFFFF 38%);}
      .block-container {padding-top: 1.25rem; padding-bottom: 4rem; max-width: 1240px;}
      [data-testid="stToolbar"], #MainMenu, footer {display:none !important;}
      h1, h2, h3 {color:var(--wh-deep) !important; letter-spacing:-0.02em;}
      h2 {font-size:1.55rem !important; margin-top:1.1rem !important;}
      h3 {font-size:1.18rem !important; margin-top:1.1rem !important;}
      p {color:var(--wh-ink);}
      [data-testid="stCaptionContainer"] p {color:#111111 !important; font-weight:500;}
      .wh-hero {
        display:flex; align-items:center; justify-content:space-between; gap:34px;
        background:linear-gradient(125deg, var(--wh-deep) 0%, #006846 68%, var(--wh-green) 100%);
        padding:32px 38px; border-radius:22px; color:white; overflow:hidden;
        box-shadow:0 18px 46px rgba(0,72,56,.20); position:relative;
      }
      .wh-hero:after {
        content:""; position:absolute; width:270px; height:270px; right:-70px; top:-130px;
        border-radius:50%; border:45px solid rgba(231,242,234,.13);
      }
      .wh-hero-copy {position:relative; z-index:2; max-width:900px;}
      .wh-hero-logo {
        position:relative; z-index:2; flex:0 0 auto; width:190px; min-height:130px;
        display:flex; align-items:center; justify-content:center;
        background:white; border-radius:18px; padding:18px 20px;
        box-shadow:0 12px 28px rgba(0,40,30,.18);
      }
      .wh-hero-logo img {display:block; width:100%; height:auto;}
      .wh-eyebrow {font-size:.78rem; font-weight:800; letter-spacing:.13em; text-transform:uppercase; opacity:.82;}
      .wh-hero h1 {color:white !important; font-size:2.45rem !important; line-height:1.08; margin:.45rem 0 .65rem !important;}
      .wh-hero p {color:#EDF8F1; font-size:1.04rem; line-height:1.55; margin:0; max-width:720px;}
      .wh-badges {display:flex; flex-wrap:wrap; gap:9px; margin-top:18px;}
      .wh-badge {background:rgba(255,255,255,.14); border:1px solid rgba(255,255,255,.24); color:white; padding:7px 11px; border-radius:999px; font-size:.82rem; font-weight:650;}
      .wh-section-label {color:var(--wh-green); font-size:.76rem; font-weight:800; letter-spacing:.12em; text-transform:uppercase; margin-top:1.7rem; margin-bottom:.25rem;}
      .small-note {color:var(--wh-muted); font-size:0.92rem;}
      div[data-testid="stMetric"] {
        background:var(--wh-white); border:1px solid var(--wh-border);
        border-top:4px solid var(--wh-green); padding:17px 18px 15px;
        border-radius:15px; box-shadow:0 7px 20px rgba(0,72,56,.075);
        min-height:140px;
      }
      div[data-testid="stMetricLabel"] {font-weight:700; color:var(--wh-muted);}
      div[data-testid="stMetricValue"] {color:var(--wh-deep); font-weight:750;}
      div[data-testid="stMetricDelta"] {background:var(--wh-mint); padding:4px 8px; border-radius:999px; width:max-content;}
      div[data-testid="stAlert"] {border-radius:13px; border-color:var(--wh-border);}
      div[data-testid="stAlert"] p {font-weight:650;}
      div[data-baseweb="select"] > div {border-radius:12px !important; border-color:var(--wh-border) !important; background:white !important; min-height:48px;}
      .stButton > button, .stDownloadButton > button {
        min-height:49px; border-radius:12px; font-weight:750; border:1px solid var(--wh-green);
        color:var(--wh-deep); background:white; transition:.2s ease;
      }
      .stButton > button:hover, .stDownloadButton > button:hover {
        border-color:var(--wh-deep); color:var(--wh-deep); box-shadow:0 8px 18px rgba(0,72,56,.12); transform:translateY(-1px);
      }
      .stButton > button[kind="primary"] {background:var(--wh-green); color:white; border-color:var(--wh-green);}
      .stButton > button[kind="primary"]:hover {background:var(--wh-deep); color:white;}
      [data-testid="stExpander"] {background:white; border:1px solid var(--wh-border); border-radius:13px; overflow:hidden;}
      [data-testid="stTabs"] [role="tablist"] {gap:8px; border-bottom:1px solid var(--wh-border);}
      [data-testid="stTabs"] button[role="tab"] {color:var(--wh-deep); font-weight:750; padding:10px 16px;}
      [data-testid="stTabs"] button[role="tab"][aria-selected="true"] {background:var(--wh-mint); border-radius:10px 10px 0 0;}
      [data-testid="stTabs"] [role="tabpanel"] {color:#111111 !important;}
      [data-testid="stTabs"] [role="tabpanel"] h1,
      [data-testid="stTabs"] [role="tabpanel"] h2,
      [data-testid="stTabs"] [role="tabpanel"] h3,
      [data-testid="stTabs"] [role="tabpanel"] p,
      [data-testid="stTabs"] [role="tabpanel"] li {color:#111111 !important; opacity:1 !important;}
      .summary-card {background:#FFFFFF; border:1px solid var(--wh-border); border-left:5px solid var(--wh-green); border-radius:14px; padding:16px 20px; margin:12px 0 10px; box-shadow:0 6px 18px rgba(0,72,56,.06);}
      .summary-card ul {margin:0; padding-left:1.25rem;}
      .summary-card li {color:#111111 !important; font-size:.94rem; line-height:1.5; margin:.42rem 0; font-weight:500;}
      .flow-box {
        border:1px solid var(--wh-border); border-radius:16px; padding:18px 22px;
        background:var(--wh-mint); margin-top:15px; box-shadow:0 5px 14px rgba(0,72,56,.05);
      }
      .flow-line {font-size:.98rem; line-height:1.75; color:var(--wh-deep); text-align:center;}
      .source-ribbon {display:grid; grid-template-columns:1fr auto 1fr auto 1.15fr; gap:12px; align-items:center; margin:18px 0 8px;}
      .source-card {background:white; border:1px solid var(--wh-border); border-radius:15px; padding:17px; min-height:108px; box-shadow:0 6px 18px rgba(0,72,56,.06);}
      .source-card.green {background:var(--wh-deep); border-color:var(--wh-deep);}
      .source-card.green strong, .source-card.green span {color:white;}
      .source-icon {width:34px; height:34px; border-radius:10px; display:grid; place-items:center; background:var(--wh-mint); color:var(--wh-deep); font-size:1.15rem; margin-bottom:10px;}
      .source-card strong {display:block; color:var(--wh-deep); font-size:.94rem; margin-bottom:4px; font-weight:850; text-transform:uppercase; letter-spacing:.035em;}
      .source-card span {display:block; color:var(--wh-muted); font-size:.82rem; line-height:1.35;}
      .source-arrow {color:var(--wh-green); font-size:1.5rem; font-weight:800;}
      .readiness-intro {color:#111111; font-size:.94rem; line-height:1.5; max-width:950px; margin:0 0 15px;}
      .readiness-grid {display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin:10px 0 22px;}
      .readiness-card {background:white; border:1px solid var(--wh-border); border-top:4px solid var(--wh-green); border-radius:14px; padding:16px; box-shadow:0 6px 18px rgba(0,72,56,.06);}
      .readiness-card .check-no {display:inline-grid; place-items:center; width:28px; height:28px; border-radius:50%; background:var(--wh-mint); color:var(--wh-deep); font-size:.78rem; font-weight:850; margin-bottom:10px;}
      .readiness-card strong {display:block; color:var(--wh-deep); font-size:.9rem; line-height:1.3; margin-bottom:7px;}
      .readiness-card span {display:block; color:#111111; font-size:.78rem; line-height:1.42;}
      .architecture {background:linear-gradient(145deg, #F4F9F5, var(--wh-mint)); border:1px solid var(--wh-border); border-radius:22px; padding:26px; margin-top:14px;}
      .architecture-purpose {display:flex; justify-content:space-between; align-items:end; gap:18px; margin-bottom:18px;}
      .architecture-purpose strong {color:var(--wh-deep); font-size:1.08rem;}
      .architecture-purpose span {color:#111111; font-size:.82rem; line-height:1.4; max-width:650px; text-align:right;}
      .implementation-intro {display:flex; align-items:center; justify-content:space-between; gap:24px; background:var(--wh-deep); color:white; border-radius:18px; padding:22px 25px; margin:12px 0 18px;}
      .implementation-intro strong {display:block; font-size:1.12rem; margin-bottom:5px;}
      .implementation-intro span {display:block; color:#DCEFE3; font-size:.9rem; line-height:1.45; max-width:760px;}
      .implementation-intro .impact {background:#BFE3CA; color:var(--wh-deep); border-radius:12px; padding:11px 15px; font-weight:800; white-space:nowrap;}
      .production-flow-wrap {background:white; border:2px solid var(--wh-green); border-radius:20px; padding:21px; margin:18px 0; box-shadow:0 9px 24px rgba(0,72,56,.08);}
      .production-flow-head {display:flex; justify-content:space-between; align-items:end; gap:18px; margin-bottom:15px;}
      .production-flow-head strong {color:var(--wh-deep); font-size:1.08rem;}
      .production-flow-head span {color:#111111; font-size:.8rem;}
      .production-flow {display:grid; grid-template-columns:repeat(5,1fr); gap:10px;}
      .production-step {position:relative; background:#F4F8F5; border:1px solid var(--wh-border); border-radius:13px; padding:14px 12px; min-height:132px;}
      .production-step:not(:last-child):after {content:"→"; position:absolute; right:-9px; top:43%; color:var(--wh-green); font-weight:900; z-index:3;}
      .production-step .num {display:inline-grid; place-items:center; width:25px; height:25px; border-radius:50%; background:var(--wh-mint); color:var(--wh-deep); font-size:.72rem; font-weight:850;}
      .production-step strong {display:block; color:var(--wh-deep); font-size:.85rem; margin:9px 0 4px;}
      .production-step span {display:block; color:#111111; font-size:.72rem; line-height:1.4;}
      .production-tools {margin-top:13px; background:var(--wh-mint); color:var(--wh-deep); border-radius:10px; padding:10px 12px; text-align:center; font-size:.78rem; font-weight:750;}
      .focus-grid {display:grid; grid-template-columns:1fr 1fr; gap:14px; margin:18px 0;}
      .focus-card {background:white; border:1px solid var(--wh-border); border-radius:17px; padding:19px 20px; box-shadow:0 7px 20px rgba(0,72,56,.06);}
      .focus-card.ai {background:linear-gradient(135deg,#062F27,#006B49); border-color:#006B49;}
      .focus-card .kicker {color:var(--wh-green); font-size:.72rem; font-weight:850; letter-spacing:.1em; text-transform:uppercase;}
      .focus-card h4 {color:var(--wh-deep); font-size:1.05rem; margin:7px 0 9px;}
      .focus-card p {color:#111111; font-size:.82rem; line-height:1.45; margin:0 0 9px;}
      .focus-card ul {margin:0; padding-left:1.1rem;}
      .focus-card li {color:#111111; font-size:.78rem; line-height:1.42; margin:.28rem 0;}
      .focus-card.ai .kicker {color:#A9D9B8;}
      .focus-card.ai h4,.focus-card.ai p,.focus-card.ai li {color:white !important;}
      .integration-core {background:white; border:2px solid var(--wh-green); border-radius:20px; padding:22px; margin:18px 0; box-shadow:0 10px 28px rgba(0,72,56,.10);}
      .integration-heading {display:flex; align-items:end; justify-content:space-between; gap:16px; margin-bottom:17px;}
      .integration-heading strong {color:var(--wh-deep); font-size:1.12rem;}
      .integration-heading span {color:#111111; font-size:.82rem;}
      .integration-flow {display:grid; grid-template-columns:1fr auto 1fr auto 1.05fr auto 1.05fr; gap:10px; align-items:stretch;}
      .integration-node {border-radius:14px; padding:17px 15px; background:#F4F8F5; border:1px solid var(--wh-border);}
      .integration-node.excel {background:#F7F4EE; border-color:#E5DCCD;}
      .integration-node.gate {background:#FFF8E6; border-color:#E9D79B;}
      .integration-node.combined {background:var(--wh-deep); border-color:var(--wh-deep);}
      .integration-node .node-icon {font-size:1.3rem; margin-bottom:8px;}
      .integration-node strong {display:block; color:var(--wh-deep); font-size:.9rem; margin-bottom:5px;}
      .integration-node span {display:block; color:#111111; font-size:.75rem; line-height:1.4;}
      .integration-node.combined strong,.integration-node.combined span {color:white;}
      .integration-plus {display:grid; place-items:center; color:var(--wh-green); font-size:1.45rem; font-weight:900;}
      .join-key {margin-top:14px; background:var(--wh-mint); color:var(--wh-deep); border-radius:10px; padding:10px 13px; font-size:.8rem; text-align:center; font-weight:750;}
      .change-grid {display:grid; grid-template-columns:repeat(3,1fr); gap:14px; margin-bottom:16px;}
      .change-card {background:white; border:1px solid var(--wh-border); border-radius:16px; padding:20px; box-shadow:0 7px 20px rgba(0,72,56,.07);}
      .change-card .tag {display:inline-block; color:var(--wh-green); background:var(--wh-mint); border-radius:999px; padding:5px 9px; font-size:.72rem; font-weight:850; letter-spacing:.05em; text-transform:uppercase;}
      .change-card .icon {font-size:1.4rem; margin:15px 0 7px;}
      .change-card h4 {color:var(--wh-deep); margin:0 0 7px; font-size:1.02rem;}
      .change-card p {color:#111111; margin:0; font-size:.88rem; line-height:1.48;}
      .section-purpose {display:flex; justify-content:space-between; align-items:end; gap:20px; margin:22px 2px 12px;}
      .section-purpose strong {color:var(--wh-deep); font-size:1.05rem;}
      .section-purpose span {color:#111111; font-size:.82rem; line-height:1.4; max-width:670px; text-align:right;}
      .section-purpose.compact {margin-top:18px;}
      .component-chips {display:flex; flex-wrap:wrap; gap:6px; margin-top:14px;}
      .component-chip {background:#F1F6F3; border:1px solid var(--wh-border); color:var(--wh-deep); border-radius:999px; padding:5px 8px; font-size:.72rem; font-weight:700;}
      .controls-strip {display:grid; grid-template-columns:repeat(4,1fr); gap:10px; margin:17px 0;}
      .control-item {border-left:4px solid var(--wh-green); background:#F4F8F5; border-radius:10px; padding:12px 13px;}
      .control-item strong {display:block; color:var(--wh-deep); font-size:.82rem; margin-bottom:3px;}
      .control-item span {display:block; color:#111111; font-size:.75rem; line-height:1.35;}
      .ppt-production {background:#F4F8F5; border:1px solid var(--wh-border); border-radius:20px; padding:23px; margin:18px 0;}
      .ppt-production-head {display:flex; justify-content:space-between; gap:18px; align-items:end; margin-bottom:17px;}
      .ppt-production-head strong {color:var(--wh-deep); font-size:1.08rem;}
      .ppt-production-head span {color:#111111; font-size:.8rem;}
      .ppt-sources {display:grid; grid-template-columns:repeat(3,1fr); gap:10px; margin-bottom:15px;}
      .ppt-source {background:white; border:1px solid var(--wh-border); border-radius:12px; padding:13px;}
      .ppt-source strong {display:block; color:var(--wh-deep); font-size:.82rem; margin-bottom:4px;}
      .ppt-source span {display:block; color:#111111; font-size:.74rem; line-height:1.38;}
      .ppt-process {display:grid; grid-template-columns:repeat(4,1fr); gap:10px;}
      .ppt-step {position:relative; background:white; border:1px solid var(--wh-border); border-radius:13px; padding:15px 13px; min-height:118px;}
      .ppt-step:not(:last-child):after {content:"→"; position:absolute; right:-9px; top:42%; color:var(--wh-green); font-weight:900; z-index:2;}
      .ppt-step .step-no {display:inline-grid; place-items:center; width:24px; height:24px; border-radius:50%; background:var(--wh-mint); color:var(--wh-deep); font-size:.72rem; font-weight:850;}
      .ppt-step strong {display:block; color:var(--wh-deep); font-size:.82rem; margin:8px 0 4px;}
      .ppt-step span {display:block; color:#111111; font-size:.73rem; line-height:1.38;}
      .ppt-note {margin-top:14px; border-left:4px solid var(--wh-green); background:white; border-radius:9px; padding:11px 13px; color:#111111; font-size:.78rem; line-height:1.4;}
      .ai-panel {display:grid; grid-template-columns:1.05fr 1.4fr; gap:16px; background:linear-gradient(135deg,#062F27,#006B49); border-radius:18px; padding:22px; margin:18px 0; color:white;}
      .ai-title .ai-kicker {color:#9EE0B1; font-size:.72rem; font-weight:850; letter-spacing:.1em; text-transform:uppercase;}
      .ai-title strong {display:block; color:white; font-size:1.08rem; margin:8px 0;}
      .ai-title p {color:#DFEEE4; margin:0; font-size:.84rem; line-height:1.5;}
      .ai-rules {display:grid; grid-template-columns:repeat(3,1fr); gap:9px;}
      .ai-rule {background:rgba(255,255,255,.09); border:1px solid rgba(255,255,255,.15); border-radius:12px; padding:13px;}
      .ai-rule strong {display:block; color:white; font-size:.8rem; margin-bottom:5px;}
      .ai-rule span {display:block; color:#DCEBE2; font-size:.72rem; line-height:1.4;}
      .before-after {display:grid; grid-template-columns:1fr auto 1fr; gap:14px; align-items:stretch; margin:18px 0 6px;}
      .state-card {border-radius:16px; padding:19px 21px;}
      .state-card.before {background:#F7F3F0; border:1px solid #E7DCD4;}
      .state-card.after {background:var(--wh-mint); border:1px solid #C9E0D1;}
      .state-card .state-label {font-size:.72rem; font-weight:850; letter-spacing:.1em; text-transform:uppercase; color:var(--wh-green);}
      .state-card strong {display:block; color:var(--wh-deep); font-size:1.02rem; margin:6px 0 3px;}
      .state-card span {color:#111111; font-size:.86rem;}
      .transform-arrow {display:grid; place-items:center; color:var(--wh-green); font-size:1.65rem; font-weight:900;}
      .arch-inputs {display:grid; grid-template-columns:1fr 1fr; gap:16px;}
      .arch-card {background:white; border:1px solid var(--wh-border); border-radius:15px; padding:18px; box-shadow:0 7px 18px rgba(0,72,56,.06);}
      .arch-card h4 {color:var(--wh-deep); margin:7px 0 4px; font-size:1rem;}
      .arch-card p {color:var(--wh-muted); margin:0; font-size:.86rem; line-height:1.45;}
      .arch-icon {width:40px; height:40px; border-radius:12px; display:grid; place-items:center; background:var(--wh-mint); font-size:1.2rem;}
      .arch-connector {text-align:center; color:var(--wh-green); font-size:1.55rem; line-height:1; padding:10px 0;}
      .arch-steps {display:grid; grid-template-columns:repeat(4,1fr); gap:12px;}
      .arch-step {background:var(--wh-deep); color:white; border-radius:15px; padding:18px 15px; text-align:center; position:relative;}
      .arch-step:not(:last-child):after {content:"→"; position:absolute; right:-18px; top:38%; color:var(--wh-green); font-size:1.25rem; font-weight:900; z-index:3;}
      .arch-step .num {display:inline-grid; place-items:center; width:25px; height:25px; border-radius:50%; background:#BFE3CA; color:var(--wh-deep); font-weight:850; font-size:.75rem; margin-bottom:8px;}
      .arch-step strong {display:block; font-size:.9rem;}
      .arch-step span {display:block; color:#D9EEE0; font-size:.78rem; line-height:1.35; margin-top:4px;}
      .outcome-banner {margin-top:16px; background:var(--wh-green); color:white; border-radius:14px; padding:15px 18px; text-align:center; font-weight:750;}
      @media (max-width:850px) {
        .wh-hero {padding:25px; flex-direction:column; align-items:flex-start;}.wh-hero h1{font-size:2rem!important}
        .wh-hero-logo {width:165px; min-height:105px; padding:14px 17px;}
        .source-ribbon{grid-template-columns:1fr}.source-arrow{transform:rotate(90deg);text-align:center}
        .arch-inputs,.arch-steps,.change-grid,.controls-strip,.integration-flow,.ai-panel,.ai-rules,.ppt-sources,.ppt-process,.readiness-grid,.production-flow,.focus-grid{grid-template-columns:1fr}.arch-step:not(:last-child):after{content:"↓";right:49%;top:auto;bottom:-19px}
        .production-step:not(:last-child):after{content:"↓";right:49%;top:auto;bottom:-16px}
        .ppt-step:not(:last-child):after{content:"↓";right:49%;top:auto;bottom:-15px}
        .integration-plus{transform:rotate(90deg);min-height:20px}
        .implementation-intro{align-items:flex-start;flex-direction:column}.implementation-intro .impact{white-space:normal}
        .before-after{grid-template-columns:1fr}.transform-arrow{transform:rotate(90deg)}
        .section-purpose,.architecture-purpose{display:block}.section-purpose span,.architecture-purpose span{display:block;text-align:left;margin-top:5px}
      }
    </style>
    """,
    unsafe_allow_html=True,
)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

@st.cache_data(show_spinner=False)
def get_history():
    return load_history(DATA_DIR)

history, file_status = get_history()
weeks = sorted(pd.to_datetime(history["Week Ending"]).dt.normalize().unique())
weeks = [pd.Timestamp(w).normalize() for w in weeks]

# Session state
if "summary" not in st.session_state:
    st.session_state.summary = None
if "summary_source" not in st.session_state:
    st.session_state.summary_source = None
if "summary_week" not in st.session_state:
    st.session_state.summary_week = None
if "ppt_bytes" not in st.session_state:
    st.session_state.ppt_bytes = None
if "ppt_week" not in st.session_state:
    st.session_state.ppt_week = None


def get_secret(name, default=None):
    # Works locally, on Streamlit Cloud, or with environment variables.
    try:
        if name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass
    return os.getenv(name, default)


def money(value):
    sign = "-" if value < 0 else ""
    value = abs(value)
    if value >= 1_000_000:
        return f"{sign}${value / 1_000_000:.2f}M"
    if value >= 1_000:
        return f"{sign}${value / 1_000:.0f}K"
    return f"{sign}${value:,.0f}"


def delta_number(value):
    if value is None:
        return None
    return f"{value:+d} vs previous week"


def delta_revenue_pct(value):
    if value is None:
        return None
    return f"{value * 100:+.1f}% vs previous week"


def summary_as_html(summary):
    """Render model or fallback output as a high-contrast HTML bullet list."""
    lines = []
    for raw_line in summary.splitlines():
        line = raw_line.strip().lstrip("•-* ").strip()
        if line:
            lines.append(html.escape(line))
    items = "".join(f"<li>{line}</li>" for line in lines)
    return f'<div class="summary-card"><ul>{items}</ul></div>'


logo_path = Path(__file__).resolve().parent / "assets" / "wesfarmers_health_logo.png"
logo_data = base64.b64encode(logo_path.read_bytes()).decode("ascii")

st.markdown(
    f"""
    <section class="wh-hero">
      <div class="wh-hero-copy">
        <div class="wh-eyebrow">Insight &amp; Automation Proof of Concept</div>
        <h1>Weekly Reporting, ready before Monday</h1>
        <p>Automatically combine refreshed Power BI data with additional weekly Excel activity data, create management commentary and prepare the presentation pack.</p>
        <div class="wh-badges">
          <span class="wh-badge">Validated weekly data</span>
          <span class="wh-badge">Automated PowerPoint</span>
        </div>
      </div>
      <div class="wh-hero-logo">
        <img src="data:image/png;base64,{logo_data}" alt="Wesfarmers Health">
      </div>
    </section>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="source-ribbon">
      <div class="source-card">
        <div class="source-icon">◉</div>
        <strong>Existing database data</strong>
        <span>Already connected and refreshed automatically through Power BI</span>
      </div>
      <div class="source-arrow">+</div>
      <div class="source-card">
        <div class="source-icon">▦</div>
        <strong>Additional weekly Excel</strong>
        <span>Education sessions and brand activations previously added manually</span>
      </div>
      <div class="source-arrow">→</div>
      <div class="source-card green">
        <div class="source-icon">✓</div>
        <strong>One reporting pack</strong>
        <span>Combined KPIs, commentary and presentation ready for review</span>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.write("")

# Select from files already stored in the app. No upload control.
display_options = {
    w: f"Week ending {w.strftime('%d %B %Y')}"
    for w in weeks
}
selected_week = st.selectbox(
    "Choose reporting week",
    options=weeks,
    index=len(weeks) - 1,
    format_func=lambda w: display_options[w],
    help="The weekly Excel files are already stored in the demo. In production, they would sit in SharePoint.",
)

# Reset generated outputs if the selected week changes.
if st.session_state.summary_week != selected_week:
    st.session_state.summary = None
    st.session_state.summary_source = None
    st.session_state.ppt_bytes = None
    st.session_state.ppt_week = None

kpis = calculate_kpis(history, selected_week)
selected_file = kpis["current_df"]["Source File"].iloc[0]

selected_week_text = selected_week.strftime("%d %B %Y")
if kpis["previous_week"] is not None:
    comparison_text = f"Comparison is ready against {kpis['previous_week'].strftime('%d %B %Y')}."
else:
    comparison_text = "This is the first available week, so no prior-week comparison is shown."

st.markdown('<div class="wh-section-label">01 · Checks completed before reporting</div>', unsafe_allow_html=True)
st.subheader("Why this reporting week is ready")
st.markdown(
    "<div class='readiness-intro'>Before the app calculates KPIs, writes commentary or creates the PowerPoint, it confirms that both data sources are available, accurate and aligned to the same reporting week.</div>",
    unsafe_allow_html=True,
)
st.markdown(
    f"""
    <div class="readiness-grid">
      <div class="readiness-card">
        <div class="check-no">1</div>
        <strong>Power BI refresh confirmed</strong>
        <span>Revenue and the existing database KPIs are available for the week ending {selected_week_text}.</span>
      </div>
      <div class="readiness-card">
        <div class="check-no">2</div>
        <strong>Weekly Excel file received</strong>
        <span>{selected_file} contains Team Education Sessions and In-store Brand Activations for {len(kpis['current_df'])} stores.</span>
      </div>
      <div class="readiness-card">
        <div class="check-no">3</div>
        <strong>Excel validation passed</strong>
        <span>The correct week, required columns, valid values, complete store rows and duplicate records have been checked.</span>
      </div>
      <div class="readiness-card">
        <div class="check-no">4</div>
        <strong>Combined dataset ready</strong>
        <span>Power BI and Excel records are aligned using Week Ending and Store. {comparison_text}</span>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="wh-section-label">02 · Weekly performance</div>', unsafe_allow_html=True)
st.subheader("Existing Power BI KPIs")
st.caption(
    "These represent database measures already connected to Power BI and refreshed automatically."
)
c1, c2, c3, c4 = st.columns(4)
with c1:
    st.metric(
        "Revenue",
        money(kpis["current"]["revenue"]),
        delta_revenue_pct(kpis["changes"]["revenue_pct"]),
    )
with c2:
    st.metric(
        "Revenue change",
        money(kpis["changes"]["revenue"] or 0),
        "vs previous week" if kpis["previous_week"] is not None else None,
    )
with c3:
    st.metric(
        "Revenue per active store",
        money(kpis["current"]["revenue_per_store"]),
        delta_revenue_pct(kpis["changes"]["revenue_per_store_pct"]),
    )
with c4:
    st.metric(
        "Top revenue region",
        kpis["top_region"],
        money(kpis["top_region_revenue"]),
    )

st.subheader("Additional Excel KPIs")
st.caption(
    "These measures come from the separate weekly Excel file that was previously added to the presentation manually."
)
e1, e2, e3, e4 = st.columns(4)
with e1:
    st.metric(
        "Team Education Sessions",
        f"{kpis['current']['workshops']}",
        delta_number(kpis["changes"]["workshops"]),
    )
with e2:
    st.metric(
        "In-store Brand Activations",
        f"{kpis['current']['brand_presentations']}",
        delta_number(kpis["changes"]["brand_presentations"]),
    )
with e3:
    st.metric(
        "Total additional activities",
        f"{kpis['current']['total_activities']}",
        delta_number(kpis["changes"]["total_activities"]),
    )
with e4:
    st.metric(
        "Active stores",
        f"{kpis['current']['active_stores']}",
        delta_number(kpis["changes"]["active_stores"]),
    )

st.caption(
    f"Highest activity store this week: {kpis['top_store']} "
    f"({kpis['top_store_activity']} education sessions and brand activations combined)."
)

st.write("")

# Generate commentary automatically as soon as the selected week's data is ready.
if st.session_state.summary is None or st.session_state.summary_week != selected_week:
    api_key = get_secret("OPENROUTER_API_KEY")
    model = get_secret("OPENROUTER_MODEL", "openai/gpt-4o-mini")
    with st.spinner("Automatically creating the management summary..."):
        summary, source = generate_management_summary(
            kpis,
            api_key=api_key,
            model=model,
        )
    st.session_state.summary = summary
    st.session_state.summary_source = source
    st.session_state.summary_week = selected_week
    st.session_state.ppt_bytes = None
    st.session_state.ppt_week = None

st.markdown('<div class="wh-section-label">03 · Automated reporting outputs</div>', unsafe_allow_html=True)
summary_tab, prompt_tab = st.tabs(["Management summary", "AI Summary Prompt"])
with summary_tab:
    st.subheader("AI-generated management summary")
    st.caption(
        "Created automatically from the validated Power BI and additional Excel KPIs for the selected week."
    )
    st.markdown(summary_as_html(st.session_state.summary), unsafe_allow_html=True)
    if st.session_state.summary_source == "OpenRouter":
        st.caption(
            "AI-generated wording based only on validated KPIs. All calculations were completed before the model received the data."
        )
    else:
        st.caption(
            "Summary generated automatically using the built-in fallback. Add an OpenRouter key to enable model-generated wording."
        )

with prompt_tab:
    st.subheader("Governed management-summary prompt")
    st.caption(
        "This version-controlled instruction file governs the AI response. The selected week's validated figures are appended at runtime."
    )
    st.code(load_management_prompt(), language="text", wrap_lines=True)

generate_ppt_clicked = st.button(
    "Generate PowerPoint",
    type="primary",
    use_container_width=True,
)

if generate_ppt_clicked:
    with st.spinner("Creating presentation..."):
        st.session_state.ppt_bytes = generate_powerpoint(
            kpis,
            st.session_state.summary,
        )
    st.session_state.ppt_week = selected_week

if st.session_state.ppt_bytes and st.session_state.ppt_week == selected_week:
    file_name = f"Weekly_Performance_{selected_week.strftime('%Y-%m-%d')}.pptx"
    st.success("PowerPoint created and ready to download.")
    st.download_button(
        "Download PowerPoint",
        data=st.session_state.ppt_bytes,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
    )

with st.expander("View source data"):
    display_cols = [
        "Week Ending",
        "Store",
        "Region",
        "Workshops",
        "Brand Presentations",
        "Revenue",
    ]
    source = kpis["current_df"][display_cols].copy()
    source["Week Ending"] = source["Week Ending"].dt.strftime("%d/%m/%Y")
    source = source.rename(
        columns={
            "Workshops": "Team Education Sessions",
            "Brand Presentations": "In-store Brand Activations",
        }
    )
    st.dataframe(source, use_container_width=True, hide_index=True)

st.divider()

st.markdown('<div class="wh-section-label">04 · Microsoft implementation</div>', unsafe_allow_html=True)
st.header("How I would implement this at Wesfarmers Health")
st.markdown(
    """
    <div class="implementation-intro">
      <div>
        <strong>One controlled path from two data sources to a finished weekly presentation</strong>
        <span>Keep the existing Power BI process, add the weekly Excel activity data, validate and join both sources, then automate the commentary, PowerPoint and distribution.</span>
      </div>
      <div class="impact">Power BI + Excel + AI</div>
    </div>

    <div class="production-flow-wrap">
      <div class="production-flow-head">
        <strong>Production workflow</strong>
        <span>From refreshed data to a review-ready pack</span>
      </div>
      <div class="production-flow">
        <div class="production-step"><div class="num">1</div><strong>Combine data</strong><span>Existing Power BI measures plus weekly store-level Excel activity data from SharePoint.</span></div>
        <div class="production-step"><div class="num">2</div><strong>Validate and join</strong><span>Check week, schema, values and duplicates; join using Week Ending + Store.</span></div>
        <div class="production-step"><div class="num">3</div><strong>Calculate and visualise</strong><span>Power BI calculates trusted KPIs, comparisons and presentation-ready charts.</span></div>
        <div class="production-step"><div class="num">4</div><strong>Generate the pack</strong><span>AI drafts the summary; Power Automate exports approved Power BI pages to PPTX.</span></div>
        <div class="production-step"><div class="num">5</div><strong>Distribute and review</strong><span>Save the dated deck in SharePoint and notify the reporting team in Teams.</span></div>
      </div>
      <div class="production-tools">SharePoint + Power Query → Power BI → Power Automate → Azure OpenAI or Copilot Studio → PowerPoint + Teams</div>
    </div>

    <div class="focus-grid">
      <div class="focus-card">
        <div class="kicker">Reporting controls</div>
        <h4>Only trusted data reaches the presentation</h4>
        <ul>
          <li>Failed validation stops the workflow and sends a Teams alert.</li>
          <li>Refresh time, source file, reporting week and output version are recorded.</li>
          <li>KPIs remain deterministic and a person approves the final pack.</li>
        </ul>
      </div>
      <div class="focus-card ai">
        <div class="kicker">AI-assisted reporting</div>
        <h4>AI explains validated numbers—it does not calculate them</h4>
        <p>The model receives a controlled KPI payload and produces concise management wording for review.</p>
        <ul>
          <li>Grounded inputs and a version-controlled prompt</li>
          <li>Restricted output with no invented figures or causes</li>
          <li>Output checks, fallback summary and human approval</li>
        </ul>
      </div>
    </div>

    <div class="before-after">
      <div class="state-card before">
        <div class="state-label">Before</div>
        <strong>Build the pack on Monday morning</strong>
        <span>Copy data, update commentary and assemble slides manually.</span>
      </div>
      <div class="transform-arrow">→</div>
      <div class="state-card after">
        <div class="state-label">After</div>
        <strong>Review a prepared and controlled pack</strong>
        <span>Focus on exceptions, insights and decisions instead of presentation production.</span>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.expander("How the AI is governed — technical detail"):
    st.markdown(
        """
        - **Calculations stay outside the model.** Power BI/DAX calculates every KPI and week-on-week movement before AI is called.
        - **Grounded payload.** The model receives only validated revenue, activity, comparison and leading-store results—not an unrestricted spreadsheet.
        - **Governed prompt.** The prompt fixes the five-bullet structure, KPI priority, tone, formatting rules and prohibited content.
        - **Consistent generation.** The prototype uses a low temperature of `0.2` and a short token limit to reduce variation and unnecessary text.
        - **Output controls.** The model is instructed not to invent figures, causes, recommendations or explanations that are absent from the data.
        - **Resilience.** A deterministic fallback summary keeps the weekly process running if the AI service is unavailable.
        - **Audit and approval.** In production, the prompt version, model, reporting week and output would be logged, with analyst approval before distribution.
        - **Enterprise deployment.** This demo uses OpenRouter; the Microsoft production pattern would use Azure OpenAI or Copilot Studio through managed Power Automate connections.
        """
    )

st.caption(
    "Demo data is synthetic and created only to demonstrate the reporting workflow."
)
