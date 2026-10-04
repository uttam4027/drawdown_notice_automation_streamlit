"""Drawdown Notice Automation - Streamlit UI.

app.py has no business rules. Calculations, template mapping, notices, PDFs, dummy emails and QA all come from
the modules in src/. This file only handles navigation and display.

Run:  streamlit run app.py
"""
import html
import io
import sys
import tempfile
import time
import traceback
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

import altair as alt                     # noqa: E402  (ships with Streamlit)
import pandas as pd                      # noqa: E402
import pymupdf                           # noqa: E402
import streamlit as st                   # noqa: E402

import config                            # noqa: E402
import main as pipeline                  # noqa: E402
import pdf_generator                     # noqa: E402
import reporting                         # noqa: E402
import utils                             # noqa: E402
import validator                         # noqa: E402

APP_NAME = "Drawdown Notice Automation"
SUBTITLE = "Excel to PDF and Email Preparation"
PAGES = ["Overview", "Process", "Investor Notices", "Email Automation",
         "QA Dashboard", "Analytics", "Documentation"]

st.set_page_config(page_title=APP_NAME, page_icon="📑", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
:root {
  --navy:#14284B; --blue:#2F5D9E; --teal:#0E9F8E; --ink:#1F2937; --muted:#6B7280;
  --line:#E3E8F0; --bg:#F5F7FB; --card:#FFFFFF; --ok:#16794C; --bad:#B3261E; --warn:#9A6200;
}
/* ---- chrome ---- */
#MainMenu, footer, [data-testid="stAppDeployButton"], [data-testid="stMainMenu"] {display:none !important;}
[data-testid="stHeader"] {background:transparent;}
.stApp {background:var(--bg);}
.block-container {padding-top:1.6rem; padding-bottom:2.5rem; padding-left:2.2rem; padding-right:2.2rem; max-width:1180px;}
h1,h2,h3 {color:var(--navy);}

/* ---- sidebar ---- */
[data-testid="stSidebar"] {background:linear-gradient(180deg,#14284B 0%,#1D3A66 100%);}
[data-testid="stSidebar"][aria-expanded="true"] {min-width:240px !important; max-width:240px !important; width:240px !important;}
[data-testid="stSidebar"] * {color:#E8EEF8;}
[data-testid="stSidebar"] [role="radiogroup"] {gap:2px;}
[data-testid="stSidebar"] [data-testid="stRadioOption"] {padding:9px 14px; border-radius:10px; width:100%; cursor:pointer;}
[data-testid="stSidebar"] [data-testid="stRadioOption"]:hover {background:rgba(255,255,255,.08);}
[data-testid="stSidebar"] [data-testid="stRadioOption"][data-selected="true"] {background:rgba(255,255,255,.17);}
[data-testid="stSidebar"] [data-testid="stRadioOption"][data-selected="true"] p {font-weight:700;}
[data-testid="stSidebar"] [data-testid="stRadioOption"] > div > div:first-child {display:none;}
.brand {display:flex; align-items:center; gap:12px; margin:6px 0 22px 4px;}
.brand-logo {width:38px; height:38px; border-radius:10px; background:linear-gradient(135deg,#3C7DD9,#0E9F8E);
             display:flex; align-items:center; justify-content:center; font-weight:800; font-size:.95rem; color:#fff;}
.brand-name {font-weight:700; font-size:.95rem; line-height:1.15; color:#fff;}
.brand-sub {font-size:.76rem; color:#A9B8D3; line-height:1.2;}
.side-pill {display:inline-block; margin:26px 0 4px 4px; padding:4px 11px; border-radius:999px; font-size:.74rem;
            font-weight:700; letter-spacing:.04em; background:rgba(22,185,129,.16); color:#7EE2B8 !important;}
.side-note {font-size:.76rem; color:#A9B8D3 !important; margin-left:4px;}

/* ---- hero ---- */
.hero {display:flex; justify-content:space-between; align-items:flex-start; gap:20px;
       background:linear-gradient(120deg,#14284B 0%,#2A5596 100%); color:#fff; border-radius:18px;
       padding:28px 34px; margin-bottom:18px; box-shadow:0 8px 24px rgba(20,40,75,.18);}
.eyebrow {font-size:.74rem; letter-spacing:.14em; font-weight:700; color:#9FC2F2; margin-bottom:6px;}
.hero-title {font-size:2.15rem; font-weight:800; line-height:1.1; color:#fff;}
.hero-sub {font-size:1.02rem; color:#D6E3F7; margin-top:8px;}
.hero > div:first-child {flex:1 1 auto; min-width:0;}
.hero-pill {padding:7px 15px; border-radius:999px; background:rgba(255,255,255,.14); border:1px solid rgba(255,255,255,.28);
            font-size:.8rem; font-weight:600; color:#fff; white-space:nowrap; flex:0 0 auto;}
.hero-skills {margin-top:16px;}
.hero-skills-l {font-size:.68rem; letter-spacing:.12em; font-weight:700; color:#9FC2F2; margin-bottom:7px;}
.hchip {display:inline-block; padding:4px 11px; border-radius:999px; background:rgba(255,255,255,.12);
        border:1px solid rgba(255,255,255,.24); font-size:.78rem; font-weight:600; color:#EAF1FC; margin:0 6px 6px 0;}
@media (max-width: 760px) {.hero {flex-direction:column;}}
.hero-pill b {color:#7EE2B8;}

/* ---- page header ---- */
.page-title {font-size:1.7rem; font-weight:800; color:var(--navy); margin:0; line-height:1.2;}
.page-sub {font-size:.95rem; color:var(--muted); margin:2px 0 18px 0;}

/* ---- cards ---- */
.grid {display:grid; grid-template-columns:repeat(auto-fit,minmax(175px,1fr)); gap:14px; margin-bottom:4px;}
.grid3 {display:grid; grid-template-columns:repeat(auto-fit,minmax(210px,1fr)); gap:14px;}
.kpi {position:relative; overflow:hidden; background:var(--card); border:1px solid var(--line); border-radius:14px;
      padding:17px 18px 13px; box-shadow:0 1px 3px rgba(16,24,40,.05); height:100%;}
.kpi::before {content:""; position:absolute; left:0; right:0; top:0; height:3px; background:var(--accent,#2F5D9E);}
.kpi-label {font-size:.72rem; font-weight:700; letter-spacing:.07em; text-transform:uppercase; color:var(--muted);}
.kpi-value {font-size:2rem; font-weight:800; color:var(--navy); line-height:1.2; margin-top:2px;}
.kpi-sub {font-size:.76rem; color:var(--muted); min-height:1.45em;}
.story {background:var(--card); border:1px solid var(--line); border-radius:14px; padding:16px 18px; height:100%;
        box-shadow:0 1px 3px rgba(16,24,40,.05);}
.story-h {display:flex; align-items:center; gap:10px; font-weight:700; color:var(--navy); margin-bottom:6px;}
.story-n {width:24px; height:24px; border-radius:50%; background:var(--navy); color:#fff; font-size:.78rem;
          display:flex; align-items:center; justify-content:center; font-weight:700;}
.story-t {font-size:.92rem; color:#374151; line-height:1.45;}
.card {background:var(--card); border:1px solid var(--line); border-radius:14px; padding:18px 20px;
       box-shadow:0 1px 3px rgba(16,24,40,.05);}
.section {font-size:.74rem; font-weight:700; letter-spacing:.1em; color:var(--muted); text-transform:uppercase;
          margin:24px 0 9px 2px;}

/* ---- pipeline strip ---- */
.pipeline {display:flex; align-items:center; flex-wrap:wrap; gap:8px; background:var(--card); border:1px solid var(--line);
           border-radius:14px; padding:14px 16px; box-shadow:0 1px 3px rgba(16,24,40,.05); margin-bottom:16px;}
.pstep {display:flex; align-items:center; gap:8px; padding:7px 15px; border-radius:999px; font-weight:600; font-size:.9rem;
        border:1px solid transparent;}
.pdot {display:inline-flex; width:20px; height:20px; border-radius:50%; align-items:center; justify-content:center;
       font-size:.7rem; font-weight:800; color:#fff;}
.pstep.done {background:#E7F5EE; color:var(--ok); border-color:#BFE3CF;} .pstep.done .pdot {background:var(--ok);}
.pstep.failed {background:#FDECEA; color:var(--bad); border-color:#F3C2BE;} .pstep.failed .pdot {background:var(--bad);}
.pstep.running {background:#FFF4DE; color:var(--warn); border-color:#F2D9A0;} .pstep.running .pdot {background:#E0A100;}
.pstep.pending {background:#F0F2F6; color:#6B7280; border-color:#E1E5EC;} .pstep.pending .pdot {background:#B6BDC9;}
.pstep.info {background:#EAF0FA; color:var(--navy); border-color:#D3DFF3;} .pstep.info .pdot {background:var(--blue);}
.parrow {color:#9AA5B8; font-size:1.2rem; font-weight:700;}

/* ---- chips / pills ---- */
.chip {display:inline-block; padding:6px 12px; border-radius:999px; background:#EAF0FA; color:var(--navy);
       font-size:.83rem; font-weight:600; margin:0 5px 7px 0; border:1px solid #D3DFF3;}
.result {display:flex; align-items:center; gap:10px; padding:12px 18px; border-radius:12px; font-weight:600; font-size:.97rem;}
.result.ok {background:#E7F5EE; color:var(--ok); border:1px solid #BFE3CF;}
.result.bad {background:#FDECEA; color:var(--bad); border:1px solid #F3C2BE;}
.result.idle {background:#F0F2F6; color:#4B5563; border:1px solid #E1E5EC;}
.result small {font-weight:500; opacity:.85;}
.pill {display:inline-block; padding:3px 11px; border-radius:999px; font-size:.78rem; font-weight:700;}
.pill.ok {background:#E7F5EE; color:var(--ok);} .pill.bad {background:#FDECEA; color:var(--bad);}
.pill.info {background:#EAF0FA; color:var(--navy);}

/* ---- notice / email cards ---- */
.amount-label {font-size:.72rem; font-weight:700; letter-spacing:.07em; text-transform:uppercase; color:var(--muted);}
.amount {font-size:1.9rem; font-weight:800; color:var(--navy); line-height:1.2; margin-bottom:10px;}
.rows {border-top:1px solid var(--line);}
.row {display:flex; justify-content:space-between; gap:12px; padding:9px 0; border-bottom:1px solid #EEF1F6; font-size:.92rem;}
.row:last-child {border-bottom:none;}
.row span:first-child {color:var(--muted);} .row span:last-child {font-weight:600; color:var(--ink); text-align:right;}
.mail-row {padding:11px 0; border-bottom:1px solid #EEF1F6;}
.mail-row:last-child {border-bottom:none;}
.mail-k {font-size:.7rem; font-weight:700; letter-spacing:.08em; text-transform:uppercase; color:var(--muted);}
.mail-v {font-size:1rem; font-weight:600; color:var(--ink);}
.skills {width:100%; border-collapse:collapse; font-size:.93rem;}
.skills td {padding:10px 12px; border-bottom:1px solid #EEF1F6; vertical-align:top;}
.skills td:first-child {font-weight:700; color:var(--navy); width:30%; white-space:nowrap;}
.skills tr:last-child td {border-bottom:none;}

/* ---- streamlit widgets ---- */
[data-testid="stBaseButton-primary"] {background:linear-gradient(135deg,#1F3A5F,#2F5D9E); border:none; border-radius:10px;
    font-weight:700; padding:.55rem 1.4rem; box-shadow:0 3px 10px rgba(31,58,95,.25);}
[data-testid="stBaseButton-primary"]:hover {filter:brightness(1.12);}
[data-testid="stBaseButton-secondary"] {border-radius:10px; font-weight:600; border:1px solid #CBD3E1; background:#fff;}
[data-testid="stBaseButton-primary"], [data-testid="stBaseButton-secondary"] {white-space:nowrap;}
[data-testid="stSidebar"] [data-testid="stBaseButton-secondary"] {background:transparent; border:1px solid rgba(255,255,255,.3);
    color:#E8EEF8; margin:12px 0 0 4px; font-size:.82rem; padding:.25rem .9rem;}
[data-testid="stSidebar"] [data-testid="stBaseButton-secondary"]:hover {background:rgba(255,255,255,.1);}
[data-testid="stExpander"] {border:1px solid var(--line); border-radius:12px; background:var(--card);}
[data-testid="stAlert"] {border-radius:12px;}
[data-testid="stDataFrame"] {border:1px solid var(--line); border-radius:12px; overflow:hidden;}
[data-testid="stVerticalBlockBorderWrapper"] {border-radius:14px;}
[data-testid="stTabs"] button[aria-selected="true"] {color:var(--navy); font-weight:700;}
</style>
""", unsafe_allow_html=True)


# ============================================================================ look & feel helpers (display only)
GROUPS = [("Excel", ["Excel Input", "Data Validation"]),
          ("Notice", ["Template Mapping", "Financial Calculation", "Notice Generation"]),
          ("PDF", ["PDF Generation"]),
          ("Email", ["Dummy Email"]),
          ("QA", ["PDF QA", "Audit Report"])]
ICONS = {"done": "✓", "failed": "✕", "running": "●", "pending": "○", "info": "›"}


def esc(x):
    return html.escape(str(x))


def group_states(stages):
    """Collapse the 9 pipeline stages reported by main.py into the 5 steps shown to the manager."""
    out = []
    for label, names in GROUPS:
        vals = [stages.get(n) for n in names]
        if "failed" in vals:
            s = "failed"
        elif all(v == "done" for v in vals):
            s = "done"
        elif any(v in ("running", "done") for v in vals):
            s = "running"
        else:
            s = "pending"
        out.append((label, s))
    return out


def strip_html(states, numbered=False):
    parts = []
    for i, (label, s) in enumerate(states):
        icon = str(i + 1) if numbered else ICONS[s]
        parts.append(f'<div class="pstep {s}"><span class="pdot">{icon}</span>{esc(label)}</div>')
        if i < len(states) - 1:
            parts.append('<div class="parrow">›</div>')
    return f'<div class="pipeline">{"".join(parts)}</div>'


def kpi_html(label, value, sub="", accent="#2F5D9E"):
    return (f'<div class="kpi" style="--accent:{accent}">'
            f'<div class="kpi-label">{esc(label)}</div>'
            f'<div class="kpi-value">{esc(value)}</div><div class="kpi-sub">{esc(sub)}</div></div>')


def kpi_row(items):
    st.markdown('<div class="grid">' + "".join(kpi_html(*item) for item in items) + '</div>', unsafe_allow_html=True)


def page_header(title, sub=""):
    st.markdown(f'<div class="page-title">{esc(title)}</div><div class="page-sub">{esc(sub)}</div>',
                unsafe_allow_html=True)


def section(text):
    st.markdown(f'<div class="section">{esc(text)}</div>', unsafe_allow_html=True)


def crore(x):
    return f"{x / 1e7:,.2f} Cr"          # display only: INR crore



# ============================================================================ helpers (display only)
DISPLAY_NAMES = {"2nd close NR inves": "2nd close NR investors", "3rd close NR inves": "3rd close NR investors",
                 "3rd close R invest": "3rd close R investors", "Institutional Inv": "Institutional investors",
                 "R Inv (Empl & Sponsor)": "Resident - Employee & Sponsor"}


def short(fmt):
    """Display name of a notice format (label tidy-up only; the stored value is never changed)."""
    name = str(fmt).replace("Format - ", "")
    return DISPLAY_NAMES.get(name, name)


def active_workbook():
    return st.session_state.get("workbook")           # None = bundled dataset


def stamp():
    f = reporting.SUMMARY_FILE()
    return f.stat().st_mtime if f.exists() else 0


@st.cache_data(show_spinner=False)
def investor_table(wb, wb_mtime):
    return reporting.build_investor_table(wb)[0]


def investors():
    wb = active_workbook()
    p = Path(wb) if wb else config.INPUT_WORKBOOK
    try:
        return investor_table(wb, p.stat().st_mtime)
    except Exception as exc:                          # noqa: BLE001
        st.error(f"Could not read the input workbook: {exc}")
        return pd.DataFrame()


@st.cache_data(show_spinner="Checking PDFs and emails...")
def qa_detail(_stamp, wb):
    return reporting.build_qa_detail(wb)


@st.cache_data(show_spinner=False)
def summary_cached(_stamp):
    return reporting.load_summary()


def results():
    """(summary dict, qa DataFrame) for the last run on disk; ({}, empty) if there is none."""
    s = summary_cached(stamp())
    return s, (reporting.load_qa_report() if s else pd.DataFrame())


@st.cache_data(show_spinner=False)
def build_zip(_stamp):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for folder, sub in ((config.PDF_DIR, "pdf"), (config.EMAIL_DIR, "emails")):
            for f in sorted(folder.glob("*.*")):
                z.write(f, f"{sub}/{f.name}")
    return buf.getvalue()


@st.cache_data(show_spinner="Zipping output folder...")
def build_full_zip(_stamp):
    """Zip the whole output folder (excel, pdf, emails, reports) keeping the sub-folder structure."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(config.OUT_DIR.rglob("*")):
            if f.is_file():
                z.write(f, Path("output") / f.relative_to(config.OUT_DIR))
    return buf.getvalue()


# ============================================================================ running the existing automation
def run_automation(strip_ph=None):
    """Call the existing pipeline (main.main) with a progress bar. Returns True when it ran.
    strip_ph: optional st.empty() that shows the live pipeline strip (display only)."""
    stages = {}
    bar = st.progress(0.0)
    status = st.empty()
    bar_value = [0.0]
    last_sig = [None]

    def on_progress(stage, state, done, total, folio):
        stages[stage] = state
        if strip_ph is not None:                      # refresh the strip only when a step changes
            sig = tuple(group_states(stages))
            if sig != last_sig[0]:
                last_sig[0] = sig
                strip_ph.markdown(strip_html(list(sig)), unsafe_allow_html=True)
        if stage == "Excel Input" and state == "done":
            p = 0.03
        elif stage == "Data Validation":
            p = 0.07 if state == "done" else 0.03
        elif stage == "Notice Generation" and total:
            p = 0.07 + 0.38 * done / total
            status.info(f"Generating notices: {done} / {total}  ({folio})")
        elif stage == "PDF Generation":
            p = 0.70 if state != "running" else 0.5
            if state == "running":
                status.info(f"Converting {total} notices to PDF ({pdf_generator.active_engine().split(' (')[0]})...")
        elif stage == "PDF QA" and total and state == "running":
            p = 0.70 + 0.27 * done / total
            status.info(f"QA and dummy emails: {done} / {total}  ({folio})")
        elif stage == "Audit Report":
            p = 1.0 if state != "running" else 0.98
        else:
            p = bar_value[0]
        bar_value[0] = max(bar_value[0], p)
        bar.progress(min(bar_value[0], 1.0))

    wb = active_workbook()
    t0 = time.time()
    try:
        if wb:
            with utils.use_workbook(wb):
                code = pipeline.main([], progress=on_progress)
        else:
            code = pipeline.main([], progress=on_progress)
    except Exception as exc:                          # noqa: BLE001
        status.error(f"Automation stopped: {exc}")
        with st.expander("Error details"):
            st.code(traceback.format_exc())
        st.session_state.stages = dict(stages)
        return False
    finally:
        st.cache_data.clear()
    st.session_state.stages = dict(stages)
    st.session_state.last_run = {"code": code, "seconds": round(time.time() - t0, 1),
                                 "at": time.strftime("%d %b %Y %H:%M:%S")}
    bar.progress(1.0)
    summary = reporting.load_summary()
    if code == 0 and not summary.get("Failed"):
        status.success("Automation completed.")
    else:
        status.error(f"Automation finished with issues (exit code {code}). See QA Dashboard.")
    return True


def reset_outputs():
    reporting.reset_outputs()
    for k in ("stages", "last_run", "workbook", "workbook_name", "preflight"):
        st.session_state.pop(k, None)
    st.cache_data.clear()
    st.session_state.toast = "Generated files cleared."


# ============================================================================ sidebar
st.session_state.setdefault("page", PAGES[0])
if "toast" in st.session_state:
    st.toast(st.session_state.pop("toast"))

with st.sidebar:
    st.markdown('<div class="brand"><div class="brand-logo">DN</div><div><div class="brand-name">Drawdown Notice</div>'
                '<div class="brand-sub">Automation</div></div></div>', unsafe_allow_html=True)
    page = st.radio("Navigation", PAGES, key="page", label_visibility="collapsed")
    st.markdown('<div class="side-pill">● DEMO MODE</div><div class="side-note">No real emails are sent</div>',
                unsafe_allow_html=True)
    st.button("Reset demo", on_click=reset_outputs, key="reset_side")

inv_df = investors()
summary, qa_df = results()
have_results = bool(summary)


def need_results():
    if not have_results:
        st.info("No results yet. Click Run Automation on the Overview or Process page.")
        return True
    return False


def current_stages():
    """Stage -> state for the pipeline strip: live from this session's run, else derived from the last report."""
    stages = st.session_state.get("stages")
    if stages:
        return stages
    if have_results:
        derived = {n: "done" for _, names in GROUPS for n in names}
        if summary.get("Failed") or summary.get("Batch verification") == "FAIL":
            derived["PDF QA"] = "failed"
        return derived
    return {}


def result_html():
    if not have_results:
        return '<div class="result idle">Not run yet — click <b>Run Automation</b> to generate everything.</div>'
    total, failed, secs = summary.get("Total investors"), summary.get("Failed", 0), summary.get("Run time (s)")
    if failed or summary.get("Batch verification") == "FAIL":
        return (f'<div class="result bad">✕ {failed} of {total} notices need attention '
                f'<small>· see QA Dashboard</small></div>')
    return (f'<div class="result ok">✓ All {total} notices passed QA — 0 failures '
            f'<small>· {esc(secs)} s</small></div>')


# ============================================================================ 1. OVERVIEW
def page_overview():
    ok, total = summary.get("Successful"), summary.get("Total investors")
    n = total if have_results else len(inv_df)
    formats = inv_df["Notice Format"].nunique() if not inv_df.empty else 10
    skills = ["Python & pandas", "Excel automation", "Data validation", "Financial calculations",
              "PDF & email automation", "QA & audit trail", "Dashboard (Streamlit)"]
    st.markdown(
        f'<div class="hero"><div><div class="eyebrow">DATA ANALYST ASSIGNMENT</div>'
        f'<div class="hero-title">Drawdown Notice Automation</div>'
        f'<div class="hero-sub">{esc(n)} investor notices — from one Excel file to PDF, email and QA, automatically.</div>'
        f'<div class="hero-skills"><div class="hero-skills-l">SKILLS DEMONSTRATED</div>'
        + "".join(f'<span class="hchip">{c}</span>' for c in skills) +
        '</div></div><div class="hero-pill"><b>●</b> DEMO MODE · no real emails sent</div></div>',
        unsafe_allow_html=True)

    if have_results:
        result_text = (f"{ok} of {total} notices passed QA in {esc(summary.get('Run time (s)', '?'))} s — "
                       f"{summary.get('PDFs')} PDFs and {summary.get('Dummy emails')} dummy emails ready.")
    else:
        result_text = "Run the automation to see the result."
    cards = [("1", "Problem", f"Drawdown notices for {n} investors in {formats} formats were prepared by hand — "
                              "slow and easy to get wrong."),
             ("2", "Approach", "Python reads the Excel, builds each notice, PDF and email, then checks every file."),
             ("3", "Result", result_text)]
    st.markdown('<div class="grid3">' + "".join(
        f'<div class="story"><div class="story-h"><span class="story-n">{num}</span>{title}</div>'
        f'<div class="story-t">{text}</div></div>' for num, title, text in cards) + '</div>', unsafe_allow_html=True)

    section("Automation pipeline")
    strip_ph = st.empty()
    strip_ph.markdown(strip_html(group_states(current_stages())), unsafe_allow_html=True)
    if st.button("▶  Run Automation", type="primary") and run_automation(strip_ph):
        st.rerun()                                    # refresh KPIs and last-run status
    st.markdown(result_html(), unsafe_allow_html=True)

    section("Live numbers")
    kpi_row([("Investors", n, "in the Excel file", "#2F5D9E"),
             ("PDFs generated", summary.get("PDFs", "—"), "one per investor", "#0E9F8E"),
             ("Emails prepared", summary.get("Dummy emails", "—"), "dummy .eml files", "#0E9F8E"),
             ("QA passed", f"{ok} / {total}" if have_results else "—", "every PDF and email checked", "#16794C")])


# ============================================================================ 2. PROCESS
def step_head(num, text):
    st.markdown(f'<div class="story-h"><span class="story-n">{num}</span>{esc(text)}</div>', unsafe_allow_html=True)


def page_process():
    page_header("Process", "Select the Excel, validate it, run.")

    with st.container(border=True):
        step_head(1, "Select Excel")
        choice = st.radio("Dataset", ["Default workbook", "Upload workbook"], horizontal=True,
                          index=0 if active_workbook() is None else 1, label_visibility="collapsed")
        if choice == "Default workbook":
            if active_workbook() is not None:
                st.session_state.pop("workbook", None)
                st.session_state.pop("preflight", None)
                st.rerun()
            st.caption(f"{config.INPUT_WORKBOOK.name}  |  {len(inv_df)} investors")
        else:
            up = st.file_uploader("Workbook (same layout: Working sheet + format sheets)", type=["xlsx"])
            if up is not None:
                tmp = Path(tempfile.gettempdir()) / "drawdown_uploaded.xlsx"
                tmp.write_bytes(up.getvalue())
                if st.session_state.get("workbook") != tmp or st.session_state.get("workbook_name") != up.name:
                    st.session_state.workbook, st.session_state.workbook_name = tmp, up.name
                    st.session_state.pop("preflight", None)
                    st.cache_data.clear()
                    st.rerun()

    with st.container(border=True):
        step_head(2, "Validate data")
        if st.button("Validate"):
            with st.spinner("Checking columns, template mapping and calculations..."):
                st.session_state.preflight = reporting.preflight_check(active_workbook())
        pf = st.session_state.get("preflight")
        if pf:
            if pf["issues"]:
                st.error(f"{len(pf['issues'])} issue(s) found in {pf['investors']} investors")
                st.dataframe(pd.DataFrame({"Issue": pf["issues"]}), hide_index=True, width="stretch")
            else:
                st.success(f"{pf['investors']} investors validated. No issues found.")
            with st.expander("Investors per notice format"):
                st.dataframe(pd.DataFrame({"Notice Format": [short(k) for k in pf["formats"]],
                                           "Investors": list(pf["formats"].values())}),
                             hide_index=True, width="stretch")

    with st.container(border=True):
        step_head(3, "Run automation")
        strip_ph = st.empty()
        strip_ph.markdown(strip_html(group_states(current_stages())), unsafe_allow_html=True)
        if st.button("▶  Run Automation", type="primary", key="process_run"):
            run_automation(strip_ph)

    section("Processing summary")
    s, _ = results()
    if not s:
        st.caption("No run yet.")
        return
    kpi_row([("Investors", s.get("Total investors"), "", "#2F5D9E"), ("Successful", s.get("Successful"), "", "#16794C"),
             ("Failed", s.get("Failed"), "", "#B3261E" if s.get("Failed") else "#16794C"),
             ("PDFs generated", s.get("PDFs"), "", "#0E9F8E"), ("Emails prepared", s.get("Dummy emails"), "", "#0E9F8E")])
    err = reporting.load_error_report()
    if not err.empty:
        st.error("Errors recorded")
        st.dataframe(err, hide_index=True, width="stretch")
    qa_file = reporting.QA_REPORT()
    st.write("")
    d1, d2, d3 = st.columns([1.2, 1.3, 1.6])
    d1.download_button("QA report (.xlsx)", qa_file.read_bytes(), qa_file.name)
    d2.download_button("PDFs + emails (.zip)", build_zip(stamp()), "drawdown_outputs.zip")
    d3.download_button("Full output folder (.zip)", build_full_zip(stamp()),
                       f"drawdown_output_folder_{time.strftime('%Y%m%d_%H%M%S')}.zip",
                       "application/zip", type="primary")


# ============================================================================ 3. INVESTOR NOTICES
@st.cache_data(show_spinner=False)
def render_pdf(path, mtime):
    doc = pymupdf.open(path)
    pages = [p.get_pixmap(dpi=110).tobytes("png") for p in doc]
    doc.close()
    return pages


def page_notices():
    page_header("Investor Notices", "Pick a folio to see its figures and the generated PDF.")
    if inv_df.empty:
        return
    left, right = st.columns([5, 7], gap="large")
    with left:
        folio = st.selectbox("Folio", inv_df["Folio"].tolist())
        row = inv_df[inv_df["Folio"] == folio].iloc[0]
        rows = [("Folio", row["Folio"]), ("Notice format", short(row["Notice Format"])),
                ("Drawdown date", utils.format_date(row["Drawdown Date"])),
                ("Capital commitment", utils.fmt_amount(row["Capital Commitment"])),
                ("Previous contribution", utils.fmt_amount(row["Previous Contribution"])),
                ("Drawdown amount", utils.fmt_amount(row["Drawdown Amount"]))]
        st.markdown(
            '<div class="card"><div class="amount-label">Amount payable (INR)</div>'
            f'<div class="amount">{esc(utils.fmt_amount(row["Amount Payable"]))}</div><div class="rows">'
            + "".join(f'<div class="row"><span>{esc(k)}</span><span>{esc(v)}</span></div>' for k, v in rows)
            + '</div></div>', unsafe_allow_html=True)

        pdf = config.PDF_DIR / f"{utils.notice_basename(folio, row['Template'])}.pdf"
        st.write("")
        if pdf.exists():
            st.download_button("Download PDF", pdf.read_bytes(), pdf.name, "application/pdf", type="primary")
    with right:
        if not pdf.exists():
            st.info("PDF not generated yet. Run the automation first.")
            return
        pages = render_pdf(str(pdf), pdf.stat().st_mtime)
        tabs = st.tabs([f"Page {i}" for i in range(1, len(pages) + 1)])
        for tab, img in zip(tabs, pages):
            with tab:
                st.image(img, width="stretch")


# ============================================================================ 4. EMAIL AUTOMATION
def page_email():
    page_header("Email Automation", "Each investor gets a dummy email with their PDF attached.")
    st.markdown('<span class="pill ok">● DEMO MODE</span> &nbsp;<span class="pill info">'
                'Dummy .eml files only — nothing is sent</span>', unsafe_allow_html=True)
    st.write("")
    if inv_df.empty or need_results():
        return
    folio = st.selectbox("Investor / Folio", inv_df["Folio"].tolist())
    template = inv_df.loc[inv_df["Folio"] == folio, "Template"].iloc[0]
    eml = config.EMAIL_DIR / f"{utils.email_basename(folio, template)}.eml"
    if not eml.exists():
        st.warning("No email file for this investor in the last run.")
        return
    m = reporting.read_eml(eml)
    st.markdown(
        '<div class="card">'
        f'<div class="mail-row"><div class="mail-k">To</div><div class="mail-v">{esc(m["to"])}</div></div>'
        f'<div class="mail-row"><div class="mail-k">Subject</div><div class="mail-v">{esc(m["subject"])}</div></div>'
        '<div class="mail-row"><div class="mail-k">Attachment</div>'
        f'<div class="mail-v"><span class="chip">{esc(m["attachment"])} · {esc(m["attachment_kb"])} KB</span></div></div>'
        '<div class="mail-row"><div class="mail-k">Status</div>'
        '<div class="mail-v"><span class="pill ok">Dummy email prepared</span></div></div></div>',
        unsafe_allow_html=True)
    st.write("")
    with st.expander("Email body"):
        st.text(m["body"])

    with st.expander("All emails"):
        rows = []
        for _, r in inv_df.iterrows():
            f = config.EMAIL_DIR / f"{utils.email_basename(r['Folio'], r['Template'])}.eml"
            if f.exists():
                x = reporting.read_eml(f)
                rows.append({"Folio": r["Folio"], "To": x["to"], "Subject": x["subject"],
                             "Attachment": x["attachment"]})
        st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")


# ============================================================================ 5. QA DASHBOARD
def style_status(df, cols):
    def color(v):
        return {"PASS": "color:#1B7F4B;font-weight:600", "FAIL": "color:#B3261E;font-weight:600"}.get(v, "")
    return df.style.map(color, subset=cols)


def check_result(checks, labels):
    """Combine rows of the backend check table into one PASS/FAIL."""
    sub = checks[checks["Check"].isin(labels)]
    if sub.empty or (sub["Result"] == "n/a").all():
        return "n/a"
    return "FAIL" if (sub["Result"] == "FAIL").any() else "PASS"


def page_qa():
    page_header("QA Dashboard", "Every PDF and email is checked automatically.")
    if need_results():
        return
    detail, checks = qa_detail(stamp(), active_workbook())
    total = len(detail)
    passed = int((detail["QA Status"] == "PASS").sum())
    pdf_ok = int((detail["PDF Status"] == "PASS").sum())
    mail_ok = int((detail["Email Status"] == "PASS").sum())
    calc_ok = int((detail["Calculation Status"] == "PASS").sum())

    good, bad = "#16794C", "#B3261E"
    kpi_row([("Investors passed", f"{passed} / {total}", f"{total - passed} failed", good if passed == total else bad),
             ("PDF QA", f"{pdf_ok} / {total}", "content, references, layout", good if pdf_ok == total else bad),
             ("Email QA", f"{mail_ok} / {total}", "recipient, subject, attachment", good if mail_ok == total else bad),
             ("Calculation QA", f"{calc_ok} / {total}", "figures re-checked", good if calc_ok == total else bad)])

    section("Validation checks")
    results_by_check = [
        ("Broken references", check_result(checks, ["#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "Cross-sheet references"])),
        ("Financial values", "PASS" if calc_ok == total else "FAIL"),
        ("PDF generated", check_result(checks, ["Missing PDF"])),
        ("Email generated", "PASS" if mail_ok == total else "FAIL"),
        ("Attachment match", check_result(checks, ["Wrong attachment"]))]
    cls = {"PASS": "done", "FAIL": "failed", "n/a": "pending"}
    st.markdown('<div class="pipeline">' + "".join(
        f'<div class="pstep {cls[r]}"><span class="pdot">{ICONS[cls[r]]}</span>{esc(name)}</div>'
        for name, r in results_by_check) + '</div>', unsafe_allow_html=True)

    recs = reporting.load_records(active_workbook())
    fails = validator.run_a1_001_calc_test(recs) + validator.check_a1_001_outputs()
    st.write("")
    if fails:
        st.error("A1-001 reference test failed: " + "; ".join(fails))
    else:
        st.markdown('<span class="pill ok">✓ A1-001 reference test passed</span> &nbsp;'
                    '<span class="pill info">calculation and PDF content</span>', unsafe_allow_html=True)

    section("Investor results")
    flt = st.radio("Filter", ["All", "PASS", "FAIL"], horizontal=True, label_visibility="collapsed")
    view = detail if flt == "All" else detail[detail["QA Status"] == flt]
    cols_show = ["Folio", "Template", "PDF Status", "Email Status", "Calculation Status", "QA Status"]
    view = view.assign(Template=view["Template"].map(short))
    st.dataframe(style_status(view[cols_show], cols_show[2:]), hide_index=True, width="stretch")

    with st.expander("Detailed checks"):
        st.dataframe(style_status(checks, ["Result"]), hide_index=True, width="stretch")
    with st.expander("Known data gaps"):
        for w in config.KNOWN_WARNINGS:
            st.markdown(f"- {w}")
    err = reporting.load_error_report()
    if not err.empty:
        with st.expander(f"Error report ({len(err)})", expanded=True):
            st.dataframe(err, hide_index=True, width="stretch")


# ============================================================================ 6. ANALYTICS
def bar_chart(df, label_col, value_col, fmt):
    base = alt.Chart(df).encode(
        y=alt.Y(f"{label_col}:N", sort="-x", title=None, axis=alt.Axis(labelLimit=240, ticks=False, domain=False)),
        x=alt.X(f"{value_col}:Q", title=None, axis=alt.Axis(grid=True, ticks=False, domain=False)))
    bars = base.mark_bar(color="#2F5D9E", cornerRadiusEnd=4, size=18)
    text = base.mark_text(align="left", dx=5, color="#1F2937", fontWeight=600).encode(
        text=alt.Text(f"{value_col}:Q", format=fmt))
    return (bars + text).properties(height=max(230, 32 * len(df))).configure_view(strokeWidth=0)


def page_analytics():
    page_header("Analytics", "The investor base at a glance (INR crore).")
    if inv_df.empty:
        return
    d = inv_df.assign(Format=inv_df["Notice Format"].map(short))
    kpi_row([("Capital commitment", crore(d["Capital Commitment"].sum()), "all investors", "#2F5D9E"),
             ("Previous contribution", crore(d["Previous Contribution"].sum()), "1st to 14th drawdown", "#2F5D9E"),
             ("Drawdown amount", crore(d["Drawdown Amount"].sum()), "this drawdown", "#0E9F8E"),
             ("Amount payable", crore(d["Amount Payable"].sum()), "after adjustments", "#0E9F8E")])

    by_count = d.groupby("Format").size().rename("Investors").reset_index()
    by_pay = d.groupby("Format")["Amount Payable"].sum().div(1e7).round(2).rename("Amount payable (Cr)").reset_index()
    a, b = st.columns(2, gap="large")
    with a:
        section("Investors by notice format")
        st.altair_chart(bar_chart(by_count, "Format", "Investors", "d"), width="stretch")
    with b:
        section("Amount payable by notice format (Cr)")
        st.altair_chart(bar_chart(by_pay, "Format", "Amount payable (Cr)", ",.2f"), width="stretch")


# ============================================================================ 7. DOCUMENTATION
def page_docs():
    page_header("Documentation", "The problem, the approach and the skills behind it.")

    section("Business problem")
    st.markdown('<div class="card">Each drawdown needs a notice for every investor, in one of 10 formats, then a PDF '
                'and an email. By hand this is slow, repetitive and one wrong figure goes straight to an investor.</div>',
                unsafe_allow_html=True)

    section("Approach")
    st.markdown(strip_html([(label, "info") for label in
                            ("Read Excel", "Validate", "Calculate", "Notice", "PDF", "Email", "QA")], numbered=True),
                unsafe_allow_html=True)

    section("Data analyst skills shown")
    st.markdown('<div class="card"><table class="skills">'
                '<tr><td>Python &amp; pandas</td><td>Reads and validates every investor record; recalculates the figures.</td></tr>'
                '<tr><td>Excel automation</td><td>Fills 10 notice templates and repairs broken <code>#REF!</code> links.</td></tr>'
                '<tr><td>Data validation</td><td>Pre-flight checks on columns, template mapping and calculations.</td></tr>'
                '<tr><td>Process automation</td><td>One click: notices, PDFs, dummy emails and audit reports.</td></tr>'
                '<tr><td>QA &amp; reporting</td><td>Every PDF and email checked; QA report and error log kept.</td></tr>'
                '<tr><td>Dashboarding</td><td>This Streamlit app: KPIs, QA results, analytics.</td></tr>'
                '</table></div>', unsafe_allow_html=True)

    section("Technology")
    st.markdown("".join(f'<span class="chip">{t}</span>' for t in
                        ["Python", "pandas", "OpenPyXL", "PyMuPDF", "ReportLab / LibreOffice", "Streamlit"]),
                unsafe_allow_html=True)

    with st.expander("Modules"):
        st.dataframe(pd.DataFrame([
            ("excel_reader.py", "Reads Excel data"),
            ("calculation_engine.py", "Financial calculations"),
            ("template_mapper.py", "Selects notice format"),
            ("notice_generator.py", "Creates notice"),
            ("pdf_generator.py", "Creates PDF"),
            ("validator.py", "QA checks"),
            ("email_service.py", "Creates dummy emails"),
            ("reporting.py", "Generates reports"),
        ], columns=["Module", "Purpose"]), hide_index=True, width="stretch")


# ============================================================================ router
try:
    {"Overview": page_overview, "Process": page_process, "Investor Notices": page_notices,
     "Email Automation": page_email, "QA Dashboard": page_qa, "Analytics": page_analytics,
     "Documentation": page_docs}[page]()
except Exception as exc:                              # noqa: BLE001
    st.error(f"Something went wrong on this page: {exc}")
    with st.expander("Error details"):
        st.code(traceback.format_exc())