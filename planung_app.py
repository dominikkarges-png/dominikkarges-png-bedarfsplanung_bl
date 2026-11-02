"""
BEDARFSPLANUNG  |  W&P  |  S&OP Planungs-App auf Marktbereichsebene
==============================================================================

BETRIEB LOKAL:  Doppelklick auf "Planung_starten.bat". Laeuft nur auf diesem Rechner
                (localhost), ohne Anmeldung, ohne Netzverbindung.
BETRIEB CLOUD:  Nutzer mit Passwort-Hash in den Secrets hinterlegen (siehe
                secrets_vorlage.toml, Hash mit passwort_hash_erzeugen.py). Ohne
                hinterlegte Nutzer verweigert die App ausserhalb von localhost den Start.

SCHUTZ
  * Anmeldung je Person mit Name und Passwort; Rolle (MB-Verantwortung oder
    S&OP-Koordination) kommt aus den Secrets und ist nicht waehlbar.
  * Passwoerter nur als PBKDF2-SHA256-Hash mit Salt (600.000 Runden).
  * Wartezeit nach 5 Fehlversuchen je Sitzung, globale Sperre nach 15 Fehlversuchen
    in 15 Minuten, automatische Abmeldung nach Inaktivitaet (Standard 60 Minuten).
  * Kundenname erscheint erst nach der Anmeldung; der Code enthaelt keine Kundendaten.
  * Hochgeladene Dateien liegen nur im Arbeitsspeicher der Sitzung; Abmelden loescht sie.
  * Gesichert wird ein Planstand nur, wenn er aktiv als Excel heruntergeladen wird.

ABLAUF EINER PLANUNGSRUNDE (monatlich)
  1. Prognoseskript laufen lassen (ab System 0.6.0)
  2. Ergebnis-Excel hier hochladen, optional den letzten Planstand dazu
  3. MB-Verantwortliche passen den Plan an und reichen ein
  4. S&OP-Koordination gibt frei, Planstand als Excel sichern
"""

from __future__ import annotations

import base64
import gzip
import hashlib
import io
import json
import warnings
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.colors import sample_colorscale
import streamlit as st

APP_VERSION = "1.0.0"
warnings.filterwarnings("ignore", message="The 'generic' unit for NumPy timedelta", category=DeprecationWarning)

# ============================================================================
# DESIGN-TOKENS (W&P Folienmaster)
# ============================================================================
NAVY, MITTEL, HELL, LILA, GRAU, HELLGRAU, ZUSATZ = (
    "#000055", "#6688BB", "#BBCCEE", "#E4E4F6", "#999999", "#D2D2D2", "#EBEBEB")
CANVAS = "#F5F6FA"
ROT, GRUEN, ORANGE = "#B3261E", "#2E7D5B", "#C98A12"
FONT = "Calibri, Carlito, 'Segoe UI', 'Helvetica Neue', Arial, sans-serif"

STATUS = ["Offen", "Eingereicht", "Freigegeben", "Abgelehnt"]
STATUS_FARBE = {"Offen": HELLGRAU, "Eingereicht": MITTEL, "Freigegeben": NAVY, "Abgelehnt": ROT}
STATUS_TEXT = {"Offen": NAVY, "Eingereicht": "#FFFFFF", "Freigegeben": "#FFFFFF", "Abgelehnt": "#FFFFFF"}
AMPEL = {"hoch": GRUEN, "mittel": ORANGE, "gering": ROT, "nicht steuerungsrelevant": GRAU}

ROLLE_KOORD = "S&OP-Koordination"
ROLLE_MB = "MB-Verantwortung"
OHNE = "nicht zugeordnet"
SCHWELLE = 0.10        # Abweichung Plan zu Prognose je MB, ab der eine Begruendung Pflicht ist
T = 1000.0             # kg je t
MONATE = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"]
MONATE_LANG = ["Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August", "September",
               "Oktober", "November", "Dezember"]
APP_BLAETTER = ["APP_META", "APP_MB", "APP_KW", "APP_MATERIAL", "APP_DETAIL"]

st.set_page_config(page_title="Absatzplanung", layout="wide",
                   initial_sidebar_state="expanded")


# ============================================================================
# FORMATIERUNG
# ============================================================================
def de(x, d=0) -> str:
    if x is None or (isinstance(x, float) and not np.isfinite(x)):
        return "n/a"
    return f"{x:,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pct(x, vz=True, d=1) -> str:
    if x is None or not np.isfinite(x):
        return "n/a"
    return (f"{x * 100:+.{d}f}" if vz else f"{x * 100:.{d}f}").replace(".", ",") + " %"


def kw_label(w: pd.Timestamp) -> str:
    iso = w.isocalendar()
    return f"KW {iso[1]:02d}/{str(iso[0])[2:]}"


def monat_key(w: pd.Timestamp) -> str:
    return (w + pd.Timedelta(days=3)).strftime("%Y-%m")


def monat_label(key: str, n_kw: int | None = None) -> str:
    j, m = key.split("-")
    s = f"{MONATE[int(m) - 1]} {j[2:]}"
    return s if n_kw is None or n_kw >= 4 else f"{s} ({n_kw} KW)"


def h(text) -> str:
    return (str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


# ============================================================================
# STYLING
# ============================================================================
CSS = f"""
<style>
html, body, .stApp, [data-testid="stAppViewContainer"], button, input, textarea, select,
[data-testid="stMarkdownContainer"], [data-testid="stMetric"] {{ font-family: {FONT} !important; }}
.stApp {{ background: {CANVAS}; color: {NAVY}; }}
[data-testid="stHeader"] {{ background: transparent; }}
[data-testid="stToolbar"], [data-testid="stDecoration"], #MainMenu, footer {{ display: none !important; }}
.block-container {{ padding-top: 1.6rem; padding-bottom: 3rem; max-width: 1480px; }}
h1, h2, h3, h4, p, li, label, span {{ color: {NAVY}; }}
h3 {{ font-size: 1.15rem !important; font-weight: 700 !important; margin-bottom: .2rem !important; }}

/* Sidebar */
[data-testid="stSidebar"] {{ background: {NAVY}; border-right: none; }}
[data-testid="stSidebar"] * {{ color: #FFFFFF; }}
[data-testid="stSidebar"] [data-testid="stSidebarNav"] a {{ border-radius: 6px; padding: .35rem .6rem; }}
[data-testid="stSidebar"] [data-testid="stSidebarNav"] a:hover {{ background: rgba(255,255,255,.08); }}
[data-testid="stSidebar"] [data-testid="stSidebarNav"] a[aria-current="page"] {{ background: rgba(255,255,255,.14); }}
[data-testid="stSidebar"] [data-testid="stSidebarNav"] a span {{ color: #FFFFFF !important; font-size: .98rem; }}
[data-testid="stSidebar"] [data-baseweb="select"] > div,
[data-testid="stSidebar"] input {{ background: #FFFFFF !important; }}
[data-testid="stSidebar"] [data-baseweb="select"] *, [data-testid="stSidebar"] input {{ color: {NAVY} !important; }}
[data-testid="stFileUploaderDropzoneInstructions"] small {{ display:none; }}
[data-testid="stFileUploaderDropzone"] {{ background: rgba(255,255,255,.07);
   border: 1px dashed rgba(255,255,255,.35); }}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button {{ background: #FFFFFF; border: none; }}
[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button * {{ color: {NAVY} !important; }}
[data-testid="stSidebar"] hr {{ border-color: rgba(255,255,255,.18); }}
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] * {{ color: rgba(255,255,255,.72) !important; }}
[data-testid="stSidebar"] [data-baseweb="button-group"] button {{ background: transparent;
   border-color: rgba(255,255,255,.35); }}
[data-testid="stSidebar"] [data-baseweb="button-group"] button[aria-checked="true"],
[data-testid="stSidebar"] [data-baseweb="button-group"] button[kind="segmented_controlActive"] {{
   background: #FFFFFF; }}
[data-testid="stSidebar"] [data-baseweb="button-group"] button[kind="segmented_controlActive"] * {{ color: {NAVY} !important; }}
[data-testid="stSidebarNav"]::before {{ content:"W&P   Bedarfsplanung"; display:block;
   white-space:pre; color:#FFFFFF; font-size:1.02rem; font-weight:700; padding:.2rem .6rem 1rem .6rem; }}
.wp-marke {{ display:flex; align-items:baseline; gap:.5rem; margin:.2rem 0 1.1rem 0; }}
.wp-marke b {{ font-size:1.35rem; letter-spacing:.01em; }}
.wp-marke span {{ font-size:.9rem; opacity:.75; }}
.wp-schutz {{ font-size:.82rem; line-height:1.35; background:rgba(255,255,255,.08); border-radius:6px;
   padding:.6rem .7rem; }}

/* Kopf */
.wp-kopf {{ display:flex; justify-content:space-between; align-items:flex-end; gap:1rem;
   border-bottom:2px solid {NAVY}; padding-bottom:.55rem; margin-bottom:1.1rem; }}
.wp-kopf h1 {{ font-size:1.75rem; font-weight:700; margin:0; line-height:1.15; padding:0; }}
.wp-kopf .kern {{ font-size:1.05rem; color:{MITTEL}; margin-top:.15rem; }}
.wp-kopf .meta {{ text-align:right; font-size:.88rem; color:{MITTEL}; line-height:1.45; white-space:nowrap; }}

/* Freigabeband */
.wp-band {{ display:flex; height:46px; border-radius:6px; overflow:hidden; background:{ZUSATZ}; }}
.wp-band div {{ display:flex; align-items:center; justify-content:center; font-size:.85rem; font-weight:700;
   border-right:2px solid {CANVAS}; min-width:0; overflow:hidden; white-space:nowrap; }}
.wp-legende {{ display:flex; flex-wrap:wrap; gap:1.1rem; font-size:.85rem; margin:.45rem 0 0 0; color:{NAVY}; }}
.wp-legende i {{ display:inline-block; width:.8rem; height:.8rem; border-radius:2px; margin-right:.35rem;
   vertical-align:-1px; }}

/* Karten und Kennzahlen */
[data-testid="stVerticalBlockBorderWrapper"] {{ background:#FFFFFF; border-color:{ZUSATZ} !important;
   border-radius:8px; }}
[data-testid="stMetric"] {{ background:#FFFFFF; border-color:{ZUSATZ} !important; border-radius:8px;
   padding:.85rem 1rem; }}
[data-testid="stMetricLabel"] p {{ font-size:.92rem !important; color:{MITTEL} !important; }}
[data-testid="stMetricValue"] {{ font-size:1.9rem !important; font-weight:700; color:{NAVY};
   font-variant-numeric: tabular-nums; }}
.wp-satz {{ font-size:1.02rem; line-height:1.5; margin:.1rem 0 .2rem 0; }}
.wp-satz b {{ font-weight:700; }}
.wp-hinweis {{ font-size:.86rem; color:{GRAU}; }}
.wp-pill {{ display:inline-block; padding:.12rem .55rem; border-radius:999px; font-size:.8rem; font-weight:700; }}

/* Mehrfachauswahl: Chips hell mit dunkler Schrift */
[data-testid="stMultiSelectTagsContainer"] > span > span {{ background:{LILA} !important;
   border:1px solid {HELL}; color:{NAVY} !important; }}
[data-testid="stMultiSelectTagsContainer"] > span * {{ color:{NAVY} !important; fill:{NAVY} !important; }}
/* Zeitstrahl */
[data-testid="stSlider"] [role="slider"] {{ background:{NAVY} !important; box-shadow:none !important; }}
[data-testid="stSliderThumbValue"] {{ color:{NAVY} !important; font-weight:600; font-size:.85rem; }}
[data-testid="stSliderTickBarMin"], [data-testid="stSliderTickBarMax"] {{ color:{GRAU} !important; font-size:.8rem; }}
/* Ansicht umschalten (Seitenleiste): Optionen als Schaltflaechen, aktive weiss */
[data-testid="stSidebar"] [role="radiogroup"] {{ gap:.45rem; }}
[data-testid="stSidebar"] [role="radiogroup"] label {{ border:1px solid rgba(255,255,255,.45); border-radius:6px;
   padding:.32rem .9rem; margin:0; cursor:pointer; }}
[data-testid="stSidebar"] [role="radiogroup"] label > div:first-child {{ display:none; }}
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) {{ background:#FFFFFF; border-color:#FFFFFF; }}
[data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) * {{ color:{NAVY} !important; font-weight:700; }}
/* Kernaussagen, Hinweise, Filter */
.wp-kern {{ background:#FFFFFF; border:1px solid {ZUSATZ}; border-left:4px solid {NAVY}; border-radius:8px;
   padding:.85rem 1.2rem .6rem 1.2rem; margin:0 0 1rem 0; }}
.wp-kern-titel {{ font-weight:700; font-size:1.05rem; margin-bottom:.35rem; color:{NAVY}; }}
.wp-kern-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:.1rem 2.2rem; min-height:6.4rem;
   align-content:start; }}
.wp-kern-grid p {{ margin:.15rem 0 .35rem 0; }}
.wp-banner {{ background:{LILA}; border:1px solid {HELL}; border-radius:8px; padding:.65rem 1rem; color:{NAVY}; }}
.wp-filter {{ font-size:.95rem; background:#FFFFFF; border:1px dashed {HELL}; border-radius:6px;
   padding:.5rem .85rem; color:{NAVY}; }}
[data-testid="stSidebar"] .wp-badge, [data-testid="stSidebar"] .wp-badge * {{ color:{NAVY} !important; }}
.wp-badge {{ background:#FFFFFF; font-weight:700; border-radius:6px; padding:.4rem .65rem; font-size:.86rem;
   margin-top:.3rem; }}
.wp-zitat {{ border-left:3px solid {HELL}; padding:.25rem .75rem; margin:.45rem 0; font-size:.95rem; color:{NAVY}; }}
.wp-gross {{ font-size:1.35rem; font-weight:700; color:{NAVY}; }}
.wp-leer {{ padding:1.6rem .4rem; }}
[data-testid="stMain"] [data-testid="stExpander"] details {{ background:#FFFFFF; border-color:{ZUSATZ} !important;
   border-radius:8px; }}
[data-testid="stMain"] [data-testid="stExpander"] summary p {{ font-weight:600; }}
@media (max-width: 900px) {{ .wp-kern-grid {{ grid-template-columns:1fr; }} }}
/* Buttons */
[data-testid="stFormSubmitButton"] button {{ background:{NAVY} !important; border-color:{NAVY} !important; }}
[data-testid="stFormSubmitButton"] button * {{ color:#FFFFFF !important; }}
[data-testid="stSidebar"] .stButton > button {{ background:transparent; border:1px solid rgba(255,255,255,.45); }}
[data-testid="stSidebar"] .stButton > button:hover {{ background:rgba(255,255,255,.1); border-color:#FFFFFF; }}
[data-testid="stSidebar"] .stButton > button * {{ color:#FFFFFF !important; }}
[data-testid="stHeadingWithActionElements"] a, [data-testid="stHeaderActionElements"] {{ display:none !important; }}
.stButton > button, .stDownloadButton > button {{ border-radius:6px; font-weight:600; }}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {{
   background:{NAVY}; border-color:{NAVY}; color:#FFFFFF; }}
.stButton > button[kind="primary"] *, .stDownloadButton > button[kind="primary"] * {{ color:#FFFFFF; }}
[data-baseweb="tab-list"] {{ gap:1.4rem; }}
[data-baseweb="tab"] p {{ font-size:1rem; }}

/* Startseite */
.wp-start h1 {{ font-size:2.6rem; line-height:1.1; margin:.5rem 0 .6rem 0; font-weight:700; }}
.wp-start p.lead {{ font-size:1.15rem; color:{MITTEL}; max-width:44rem; line-height:1.5; }}
.wp-schritte {{ display:grid; grid-template-columns:repeat(3,1fr); gap:1rem; margin:1.6rem 0 1.4rem 0; }}
.wp-schritte div {{ border-top:3px solid {NAVY}; padding-top:.6rem; }}
.wp-schritte b {{ display:block; font-size:1.05rem; margin-bottom:.2rem; }}
.wp-schritte span {{ font-size:.95rem; color:{NAVY}; opacity:.8; line-height:1.45; }}
@media (max-width: 900px) {{ .wp-schritte {{ grid-template-columns:1fr; }} .wp-kopf {{ flex-direction:column;
   align-items:flex-start; }} .wp-kopf .meta {{ text-align:left; }} }}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


# ============================================================================
# DATEN EINLESEN (nur im Arbeitsspeicher)
# ============================================================================
def _engine() -> str:
    try:
        import python_calamine  # noqa: F401
        return "calamine"
    except ImportError:
        return "openpyxl"


@st.cache_data(show_spinner="Prognose wird eingelesen ...", max_entries=2)
def lese_prognose(inhalt: bytes) -> dict:
    with pd.ExcelFile(io.BytesIO(inhalt), engine=_engine()) as xl:
        fehlend = [s for s in APP_BLAETTER if s not in xl.sheet_names]
        if fehlend:
            raise ValueError("Die Datei enthält nicht die Datenblätter für die Planungs-App "
                             f"({', '.join(fehlend)}). Bitte die Ergebnis-Excel aus dem Prognoseskript ab "
                             "System 0.6.0 hochladen.")
        d = {s: xl.parse(s) for s in APP_BLAETTER}
        if "APP_ZEIT" in xl.sheet_names:
            d["APP_ZEIT"] = xl.parse("APP_ZEIT")
    return aufbereiten(d)


def aufbereiten(d: dict) -> dict:
    meta = dict(zip(d["APP_META"]["Schluessel"].astype(str), d["APP_META"]["Wert"].astype(str)))
    kw = d["APP_KW"].copy()
    kw["MB"] = kw["MB"].astype(str)
    kw["KW_Start"] = pd.to_datetime(kw["KW_Start"]).dt.normalize()
    mbinfo = d["APP_MB"].copy()
    mbinfo["MB"] = mbinfo["MB"].astype(str)
    mat = d["APP_MATERIAL"].copy()
    det = d["APP_DETAIL"].copy()
    for df_ in (mat, det):
        for c in ("MB", "Material", "Land", "Landname", "Auftraggeber"):
            if c in df_:
                df_[c] = df_[c].astype(str).str.replace(r"\.0$", "", regex=True)

    p = kw[(kw["Typ"] == "Prognose") & (kw["MB"] != "GESAMT")]
    prog = p.pivot(index="MB", columns="KW_Start", values="Menge_kg").sort_index(axis=1).fillna(0.0)
    bu = p.pivot(index="MB", columns="KW_Start", values="Band_unten_kg").reindex_like(prog)
    bo = p.pivot(index="MB", columns="KW_Start", values="Band_oben_kg").reindex_like(prog)
    ist = (kw[kw["Typ"] == "Ist"].pivot(index="MB", columns="KW_Start", values="Menge_kg")
           .sort_index(axis=1).fillna(0.0).reindex(prog.index, fill_value=0.0))
    g = kw[(kw["Typ"] == "Prognose") & (kw["MB"] == "GESAMT")].set_index("KW_Start").reindex(prog.columns)
    if g["Menge_kg"].isna().all():
        g = pd.DataFrame({"Menge_kg": prog.sum(), "Band_unten_kg": bu.sum(), "Band_oben_kg": bo.sum()})
    reihenfolge = prog.sum(axis=1).sort_values(ascending=False).index
    W = list(prog.columns)
    vj = pd.DataFrame({w: ist.get(w - pd.Timedelta(weeks=52), pd.Series(np.nan, index=ist.index)) for w in W})
    monat = {w: monat_key(w) for w in W}
    monate = list(dict.fromkeys(monat.values()))
    n_kw = {m: sum(1 for v in monat.values() if v == m) for m in monate}
    zeit = d.get("APP_ZEIT")
    if zeit is not None and len(zeit):
        zeit = zeit.copy()
        for c in ("MB", "Dimension", "Schluessel", "Monat"):
            zeit[c] = zeit[c].astype(str).str.replace(r"\.0$", "", regex=True)
        zeit["Menge_kg"] = pd.to_numeric(zeit["Menge_kg"], errors="coerce").fillna(0.0)
    else:
        zeit = None
    return dict(meta=meta, mbinfo=mbinfo.set_index("MB").reindex(reihenfolge), prog=prog.loc[reihenfolge],
                bu=bu.loc[reihenfolge], bo=bo.loc[reihenfolge], ist=ist.loc[reihenfolge],
                vj=vj.loc[reihenfolge].fillna(0.0), ges=g, W=W, monat=monat, monate=monate, n_kw=n_kw,
                mat=mat, det=det, zeit=zeit, mbs=list(reihenfolge))


@st.cache_data(show_spinner=False, max_entries=2)
def lese_planstand(inhalt: bytes) -> dict:
    with pd.ExcelFile(io.BytesIO(inhalt), engine=_engine()) as xl:
        if not {"APP_PLAN", "APP_STATUS", "APP_PLANMETA"} <= set(xl.sheet_names):
            raise ValueError("Die Datei ist kein Planstand dieser App. Bitte eine mit 'Planstand sichern' "
                             "erzeugte Excel hochladen.")
        plan = xl.parse("APP_PLAN")
        status = xl.parse("APP_STATUS")
        meta = xl.parse("APP_PLANMETA")
        log = xl.parse("APP_LOG") if "APP_LOG" in xl.sheet_names else pd.DataFrame()
    plan["MB"] = plan["MB"].astype(str)
    plan["KW_Start"] = pd.to_datetime(plan["KW_Start"]).dt.normalize()
    status["MB"] = status["MB"].astype(str)
    return dict(plan=plan, status=status, meta=dict(zip(meta["Schluessel"].astype(str), meta["Wert"].astype(str))),
                log=log)


# ============================================================================
# DEMODATEN (synthetisch, ohne Kundenbezug)
# ============================================================================
@st.cache_data(show_spinner=False)
def demo_daten() -> dict:
    rng = np.random.default_rng(42)
    stand = pd.Timestamp("2026-10-26")
    ist_w = pd.date_range(stand - pd.Timedelta(weeks=155), stand, freq="7D")
    prog_w = pd.date_range(stand + pd.Timedelta(days=7), periods=26, freq="7D")
    mbs = {"110": 2400, "120": 1700, "130": 1150, "210": 820, "220": 600, "310": 380, "320": 170, "410": 60}
    namen = ["Ensemble Statistik", "Monatsmodell Theta", "Wochenmodell Theta", "Monatsmodell ETS (ab 2023)",
             "Niveau 13 Wochen x Saison", "Ensemble Monat", "Intermittenz TSB", "Kalendermittel 52 Wochen"]
    kw, mbz, mat, det = [], [], [], []
    laender = [("DE", "Deutschland"), ("FR", "Frankreich"), ("IT", "Italien"), ("PL", "Polen"),
               ("ES", "Spanien"), ("NL", "Niederlande"), ("TR", "Türkei"), ("US", "USA"), ("AT", "Österreich"),
               ("CZ", "Tschechien"), ("BE", "Belgien"), ("HU", "Ungarn"), ("GB", "Großbritannien"),
               ("SE", "Schweden"), ("RO", "Rumänien"), ("CN", "China")]
    lw = np.array([.30, .12, .10, .08, .07, .06, .05, .04, .045, .035, .03, .025, .02, .015, .01, .015])
    lw = lw / lw.sum()

    def form(w):
        f = 1 + 0.18 * np.sin(2 * np.pi * (w.dayofyear - 70) / 365)
        return f * (0.25 if (w.month == 12 and w.day > 20) or (w.month == 1 and w.day < 4) else 1.0)
    ges = np.zeros(26)
    for k, (mb, basis) in enumerate(mbs.items()):
        trend = 1 + rng.normal(0.04, 0.05)
        for i, w in enumerate(ist_w):
            v = basis * T * form(w) * (1 + 0.04 * i / 52) * rng.lognormal(0, 0.16)
            kw.append((mb, w, "Ist", v, None, None))
        f = np.array([basis * T * form(w) * trend * 1.11 for w in prog_w])
        ges += f
        breite = np.linspace(0.10, 0.22, 26) * (1 + k * 0.08)
        for j, w in enumerate(prog_w):
            kw.append((mb, w, "Prognose", f[j], f[j] * (1 - breite[j]), f[j] * (1 + breite[j] * 1.2)))
        f4 = [0.09, 0.12, 0.15, 0.18, 0.24, 0.31, 0.46, 0.61][k]
        mbz.append(dict(MB=mb, Modell=namen[k], XYZ="XXYXYYZZ"[k], Regel="Wettbewerb" if k < 7 else "Kleinstmenge",
                        Fehler_4W=f4, Fehler_KW=f4 * 1.6, Bias=rng.normal(0, 0.04) if k != 4 else 0.12,
                        Fehler_4W_Vorjahresmethode=f4 * 1.45, Drift=k == 4, Drift_t=2.9 if k == 4 else rng.normal(0, 1),
                        Verlaesslichkeit=("hoch" if f4 <= .2 else "mittel" if f4 <= .35 else "gering") if k < 7
                        else "nicht steuerungsrelevant"))
        n_mat = int(rng.integers(8, 24))
        ant = rng.dirichlet(np.full(n_mat, 0.6))
        for j in range(n_mat):
            m = f"{mb}{j + 1:05d}"
            ist52 = basis * T * 52 * ant[j] * rng.lognormal(0, 0.1)
            vj52 = ist52 * rng.lognormal(-0.03, 0.18)
            mat.append(dict(MB=mb, Material=m, Anteil=ant[j], Prognose_26W_kg=f.sum() * ant[j], Ist_52W_kg=ist52,
                            Ist_VJ_52W_kg=vj52))
            for _ in range(int(rng.integers(2, 7))):
                code, name = laender[int(rng.choice(len(laender), p=lw))]
                kunde = f"{100000 + int(rng.integers(1, 60))}"
                a = rng.uniform(.05, .5)
                det.append(dict(MB=mb, Material=m, Land=code, Landname=name, Auftraggeber=kunde,
                                Ist_52W_kg=ist52 * a, Ist_VJ_52W_kg=vj52 * a * rng.lognormal(0, .25),
                                Pos_52W=int(rng.integers(3, 60))))
    for j, w in enumerate(prog_w):
        kw.append(("GESAMT", w, "Prognose", ges[j], ges[j] * (0.93 - j * 0.002), ges[j] * (1.07 + j * 0.003)))
    monate24 = pd.period_range("2024-10", "2026-09", freq="M")
    saison = np.array([1 + 0.18 * np.sin(2 * np.pi * (m.month * 30.4 - 85) / 365) * (0.6 if m.month == 12 else 1)
                       for m in monate24])
    zeilen = []
    dd = pd.DataFrame(det)
    for dim in ("Material", "Land", "Auftraggeber"):
        g = dd.groupby(["MB", dim])[["Ist_52W_kg", "Ist_VJ_52W_kg"]].sum().reset_index()
        for _, r in g.iterrows():
            tr = rng.normal(0, 0.012)
            for k_, m in enumerate(monate24):
                basis = r["Ist_VJ_52W_kg"] if k_ < 12 else r["Ist_52W_kg"]
                v = basis / 12 * saison[k_] * (1 + tr * (k_ - 12)) * rng.lognormal(0, 0.12)
                if v > 0:
                    zeilen.append((r["MB"], dim, str(r[dim]), str(m), v))
    zeit = pd.DataFrame(zeilen, columns=["MB", "Dimension", "Schluessel", "Monat", "Menge_kg"])
    meta = pd.DataFrame([("App_Version", "1"), ("System_Version", "Demo"), ("Datenstand", "2026-10-29"),
                         ("Letzte_Ist_KW", f"{stand:%Y-%m-%d}"), ("Erste_Prognose_KW", f"{prog_w[0]:%Y-%m-%d}"),
                         ("Horizont_Wochen", "26"), ("Quelldatei", "Demodaten (synthetisch)"),
                         ("Daten_SHA256", "demo"), ("Erstellt", "2026-10-30 08:00"), ("Zeitachse", "Fakturadatum"),
                         ("Einheit", "KG"), ("Fehler_4W_Gesamt", "0.084"), ("Fehler_4W_Vorjahresmethode", "0.142"),
                         ("Bias_Gesamt", "0.011")], columns=["Schluessel", "Wert"])
    return aufbereiten({"APP_META": meta, "APP_MB": pd.DataFrame(mbz),
                        "APP_KW": pd.DataFrame(kw, columns=["MB", "KW_Start", "Typ", "Menge_kg", "Band_unten_kg",
                                                           "Band_oben_kg"]),
                        "APP_MATERIAL": pd.DataFrame(mat), "APP_DETAIL": pd.DataFrame(det), "APP_ZEIT": zeit})


# ============================================================================
# SITZUNGSZUSTAND: Plan, Status, Protokoll, Ansicht
# ============================================================================
ss = st.session_state
ANSICHT_PLAN, ANSICHT_FREIGABE = "Planung", "Freigabe"
STATUS_SPALTEN = ["Status", "Begruendung", "Kommentar_Freigabe", "Geaendert_von", "Geaendert_am",
                  "Eingereicht_am", "Freigegeben_von", "Freigegeben_am"]
STATUS_ALT = {"In Bearbeitung": "Offen", "Zurückgewiesen": "Abgelehnt"}


def jetzt() -> str:
    return datetime.now().strftime("%d.%m.%Y %H:%M")


def lauf_id(D) -> str:
    return f"{D['meta'].get('Daten_SHA256', '')[:16]}|{D['meta'].get('Erste_Prognose_KW', '')}"


def nutzer_name() -> str:
    return ss.get("nutzer") or "Lokal"


def ansicht() -> str:
    return ss.get("ansicht", ANSICHT_PLAN)


def runde_starten(D: dict, vorstand: dict | None = None, startwert: str = "Modellprognose"):
    """Neue Planungsrunde oder Fortsetzung eines gesicherten Planstands."""
    plan = D["prog"].copy()
    status = pd.DataFrame({"MB": D["mbs"], **{c: "" for c in STATUS_SPALTEN}}).set_index("MB")
    status["Status"] = "Offen"
    log = []
    ss["vorplan"] = None
    if vorstand is not None:
        vs = vorstand["status"].set_index("MB").rename(columns={"Kommentar_Koordination": "Kommentar_Freigabe"})
        vp = vorstand["plan"].pivot(index="MB", columns="KW_Start", values="Plan_kg")
        if vorstand["meta"].get("Lauf_ID") == lauf_id(D):
            plan.update(vp.reindex(index=plan.index, columns=plan.columns))
            idx = status.index.intersection(vs.index)
            for c in STATUS_SPALTEN:
                if c in vs:
                    status.loc[idx, c] = vs.loc[idx, c].fillna("").astype(str)
            status["Status"] = status["Status"].replace(STATUS_ALT).replace("", "Offen")
            if len(vorstand["log"]):
                log = vorstand["log"].fillna("").to_dict("records")
            log.append(eintrag("Planstand fortgesetzt", "", "", None, None, vorstand["meta"].get("Exportiert", "")))
        else:
            ss["vorplan"] = vp
            if startwert == "Vormonatsplan":
                gemeinsam = plan.columns.intersection(vp.columns)
                plan.update(vp.reindex(index=plan.index, columns=gemeinsam))
            log.append(eintrag("Neue Planungsrunde", "", "", None, None,
                               f"Startwert {startwert}; Vormonatsplan vom {vorstand['meta'].get('Exportiert', '')}"))
    else:
        log.append(eintrag("Neue Planungsrunde", "", "", None, None, "Startwert Modellprognose"))
    ss["plan"], ss["status"], ss["log"], ss["lauf"] = plan, status, log, lauf_id(D)
    ss["auswahl"] = []
    ss["editor_v"] = ss.get("editor_v", 0) + 1


def eintrag(aktion, mb, objekt, vorher, nachher, kommentar="") -> dict:
    return {"Zeit": jetzt(), "Nutzer": nutzer_name(), "Ansicht": ansicht(), "Marktbereich": mb, "Aktion": aktion,
            "Objekt": objekt, "Vorher_t": None if vorher is None else round(vorher / T, 1),
            "Nachher_t": None if nachher is None else round(nachher / T, 1), "Kommentar": kommentar}


def protokolliere(*a, **k):
    ss["log"].append(eintrag(*a, **k))


def plan_monat(D) -> pd.DataFrame:
    return ss["plan"].T.groupby(pd.Series(D["monat"])).sum().T.reindex(columns=D["monate"])


def wert_monat(D, df) -> pd.DataFrame:
    return df.T.groupby(pd.Series(D["monat"])).sum().T.reindex(columns=D["monate"])


def bearbeitet(mb):
    ss["status"].loc[mb, ["Geaendert_von", "Geaendert_am"]] = [nutzer_name(), jetzt()]


def setze_monat(D, mb, mkey, neu_kg):
    wochen = [w for w in D["W"] if D["monat"][w] == mkey]
    alt = ss["plan"].loc[mb, wochen].astype(float)
    basis = alt if alt.sum() > 0 else D["prog"].loc[mb, wochen]
    if basis.sum() <= 0:
        basis = pd.Series(1.0, index=wochen)
    ss["plan"].loc[mb, wochen] = (basis / basis.sum() * max(neu_kg, 0.0)).to_numpy()
    protokolliere("Monat angepasst", mb, monat_label(mkey), alt.sum(), neu_kg)
    bearbeitet(mb)


def setze_woche(D, mb, w, neu_kg):
    alt = float(ss["plan"].loc[mb, w])
    ss["plan"].loc[mb, w] = max(neu_kg, 0.0)
    protokolliere("Woche angepasst", mb, kw_label(w), alt, neu_kg)
    bearbeitet(mb)


def darf_bearbeiten(mb) -> bool:
    return ansicht() == ANSICHT_PLAN and ss["status"].loc[mb, "Status"] in ("Offen", "Abgelehnt")


def geaendert(D, mb) -> bool:
    return not np.allclose(ss["plan"].loc[mb].to_numpy(float), D["prog"].loc[mb].to_numpy(float), rtol=1e-6, atol=1.0)


def abweichung(D, mb) -> float:
    p = D["prog"].loc[mb].sum()
    return ss["plan"].loc[mb].sum() / p - 1 if p > 0 else np.nan


def braucht_begruendung(D, mb) -> bool:
    a = abweichung(D, mb)
    return np.isfinite(a) and abs(a) > SCHWELLE and not str(ss["status"].loc[mb, "Begruendung"]).strip()


def mbs_mit_status(*status) -> list:
    st_ = ss["status"]
    return [mb for mb in st_.index if st_.loc[mb, "Status"] in status]


def einreichen(D, mb):
    st_ = ss["status"]
    st_.loc[mb, ["Status", "Eingereicht_am", "Kommentar_Freigabe"]] = ["Eingereicht", jetzt(), ""]
    protokolliere("Eingereicht", mb, "", D["prog"].loc[mb].sum(), ss["plan"].loc[mb].sum(),
                  st_.loc[mb, "Begruendung"])


def genehmigen(D, mb, kommentar=""):
    ss["status"].loc[mb, ["Status", "Kommentar_Freigabe", "Freigegeben_von", "Freigegeben_am"]] = [
        "Freigegeben", kommentar, nutzer_name(), jetzt()]
    protokolliere("Freigegeben", mb, "", D["prog"].loc[mb].sum(), ss["plan"].loc[mb].sum(), kommentar)


def ablehnen(D, mb, kommentar):
    ss["status"].loc[mb, ["Status", "Kommentar_Freigabe", "Freigegeben_von", "Freigegeben_am"]] = [
        "Abgelehnt", kommentar, nutzer_name(), jetzt()]
    protokolliere("Abgelehnt", mb, "", None, ss["plan"].loc[mb].sum(), kommentar)


# Prognoseguete: Genauigkeit = 100 % minus mittlere Abweichung je 4-Wochen-Block im Backtest
GUETE_HILFE = ("Backtest: Das Modell wurde nur mit Daten bis vor 26 KW gerechnet und dann mit dem tatsächlichen "
               "Absatz der letzten 26 KW verglichen. Genauigkeit = 100 % minus mittlere Abweichung je 4 Wochen. "
               "Vergleich: Vorjahresmethode (gleiche KW des Vorjahres als Prognose). "
               "Ab 80 % hoch, ab 65 % mittel, darunter gering.")


def guete_label(g: float) -> str:
    if not np.isfinite(g):
        return "n/a"
    return "hoch" if g >= 0.80 else "mittel" if g >= 0.65 else "gering"


def _zahl(x) -> float:
    try:
        return float(x)
    except (TypeError, ValueError):
        return np.nan


def genauigkeit(D, mbs) -> tuple:
    """(Modell, Vorjahresmethode). Gesamt aus dem Backtest der Summe, Teilauswahl mengengewichtet je MB."""
    if len(mbs) == len(D["mbs"]):
        g, gv = 1 - _zahl(D["meta"].get("Fehler_4W_Gesamt")), 1 - _zahl(D["meta"].get("Fehler_4W_Vorjahresmethode"))
        if np.isfinite(g):
            return max(g, 0.0), max(gv, 0.0)
    info = D["mbinfo"].loc[mbs]
    w = D["prog"].loc[mbs].sum(axis=1)
    out = []
    for spalte in ("Fehler_4W", "Fehler_4W_Vorjahresmethode"):
        f = pd.to_numeric(info[spalte], errors="coerce")
        ok = f.notna() & (w > 0)
        out.append(max(1 - float((f[ok] * w[ok]).sum() / w[ok].sum()), 0.0) if ok.any() else np.nan)
    return tuple(out)


def genauigkeit_mb(D, mb) -> float:
    return max(1 - _zahl(D["mbinfo"].loc[mb, "Fehler_4W"]), 0.0) if np.isfinite(_zahl(D["mbinfo"].loc[mb, "Fehler_4W"])) \
        else np.nan


# ============================================================================
# DIAGRAMME
# ============================================================================
def layout(fig: go.Figure, hoehe=360, y_titel="t", legende=True):
    fig.update_layout(
        height=hoehe, margin=dict(l=8, r=8, t=36 if legende else 12, b=8), paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF", font=dict(family=FONT, color=NAVY, size=13), hovermode="x unified",
        hoverlabel=dict(bgcolor="#FFFFFF", bordercolor=HELL, font=dict(family=FONT, color=NAVY)),
        showlegend=legende, legend=dict(orientation="h", x=0, y=1.1, xanchor="left", yanchor="bottom",
                                        font=dict(size=12), bgcolor="rgba(0,0,0,0)"))
    fig.update_xaxes(showgrid=False, linecolor=HELLGRAU, ticks="outside", tickcolor=HELLGRAU)
    fig.update_yaxes(gridcolor=ZUSATZ, zeroline=False, title=y_titel, title_font=dict(size=12, color=GRAU),
                     tickformat=",.0f", separatethousands=True)
    fig.update_layout(separators=",.")
    return fig


PLOT_CFG = {"displaylogo": False, "modeBarButtonsToRemove": ["select2d", "lasso2d", "autoScale2d"],
            "toImageButtonOptions": {"format": "png", "scale": 2}}


def deutsche_zeitachse(fig, x0, x1):
    """Deutsche Monatsbeschriftung; Abstand je nach Zeitspanne (Monat, Quartal, Halbjahr, Jahr)."""
    x0, x1 = pd.Timestamp(x0), pd.Timestamp(x1)
    monate = (x1.year - x0.year) * 12 + x1.month - x0.month
    schritt = 1 if monate <= 8 else 3 if monate <= 30 else 6 if monate <= 60 else 12
    werte = [d for d in pd.date_range(pd.Timestamp(x0.year, 1, 1), x1, freq="MS")
             if (d.month - 1) % schritt == 0 and d >= x0]
    text = [str(d.year) if schritt == 12 else f"{MONATE[d.month - 1]} {str(d.year)[2:]}" for d in werte]
    fig.update_xaxes(tickvals=werte, ticktext=text, hoverformat="KW %V/%G")


def chart_wochen(ist, prog, bu, bo, plan, hoehe=400, zeitraum=None, wochen_ist=104):
    """Wochenverlauf: Ist, Modellprognose mit Bandbreite und Plan; Planhorizont hinterlegt.
    zeitraum=(von, bis) als Monatsschluessel 'JJJJ-MM'; ohne Angabe die letzten 104 KW Ist."""
    if zeitraum is None:
        ist = ist.iloc[-wochen_ist:]
    else:
        von, bis = zeitraum

        def fenster(r):
            k = pd.Index([monat_key(w) for w in r.index])
            return r[(k >= von) & (k <= bis)]
        ist, prog, bu, bo, plan = (fenster(r) for r in (ist, prog, bu, bo, plan))
    ist, prog, bu, bo, plan = ist / T, prog / T, bu / T, bo / T, plan / T
    fig = go.Figure()
    zukunft = plan.index if len(plan) else prog.index
    if len(zukunft):
        h0, h1 = zukunft[0] - pd.Timedelta(days=3.5), zukunft[-1] + pd.Timedelta(days=3.5)
        fig.add_vrect(x0=h0, x1=h1, fillcolor=LILA, opacity=0.55, line_width=0, layer="below")
        fig.add_annotation(x=h0, y=1, xref="x", yref="paper", text="Planhorizont", showarrow=False,
                           xanchor="left", yanchor="top", xshift=6, yshift=-4, font=dict(size=11, color=MITTEL))
    if len(bu):
        fig.add_trace(go.Scatter(x=list(bu.index) + list(bo.index[::-1]), y=list(bu) + list(bo[::-1]),
                                 fill="toself", fillcolor="rgba(102,136,187,.18)", line=dict(width=0),
                                 name="Bandbreite 80 %", hoverinfo="skip"))
    if len(ist):
        fig.add_trace(go.Scatter(x=ist.index, y=ist, name="Ist", line=dict(color=NAVY, width=1.6),
                                 hovertemplate="%{y:,.0f} t"))
    anschluss = bool(len(ist) and len(plan) and plan.index[0] - ist.index[-1] <= pd.Timedelta(days=7))
    x0 = [ist.index[-1]] if anschluss else []
    y0 = [ist.iloc[-1]] if anschluss else []
    if len(prog):
        fig.add_trace(go.Scatter(x=x0 + list(prog.index), y=y0 + list(prog), name="Modellprognose",
                                 line=dict(color=MITTEL, width=2, dash="dash"), hovertemplate="%{y:,.0f} t"))
    if len(plan):
        fig.add_trace(go.Scatter(x=x0 + list(plan.index), y=y0 + list(plan), name="Plan",
                                 line=dict(color=NAVY, width=3.4), hovertemplate="%{y:,.0f} t"))
    fig = layout(fig, hoehe)
    alle = [r.index for r in (ist, prog, plan) if len(r)]
    if alle:
        deutsche_zeitachse(fig, min(i[0] for i in alle), max(i[-1] for i in alle))
    return fig


def chart_monate(D, plan_m, prog_m, vj_m, hoehe=320):
    x = [monat_label(m, D["n_kw"][m]) for m in D["monate"]]
    fig = go.Figure()
    for name, y, farbe in (("Vorjahr", vj_m, HELLGRAU), ("Modellprognose", prog_m, HELL), ("Plan", plan_m, NAVY)):
        fig.add_trace(go.Bar(x=x, y=np.asarray(y) / T, name=name, marker_color=farbe, hovertemplate="%{y:,.0f} t"))
    fig.update_layout(barmode="group", bargap=.28, bargroupgap=.06)
    return layout(fig, hoehe)


def chart_treiber(df, label, wert, hoehe=None, n=10):
    d = df.reindex(df[wert].abs().sort_values(ascending=False).index).head(n).iloc[::-1]
    fig = go.Figure(go.Bar(x=d[wert] / T, y=d[label], orientation="h",
                           marker_color=[NAVY if v >= 0 else ROT for v in d[wert]],
                           hovertemplate="%{x:+,.0f} t<extra></extra>"))
    fig = layout(fig, hoehe or 34 * len(d) + 60, y_titel="", legende=False)
    fig.update_layout(hovermode="closest")
    fig.update_xaxes(title="Veränderung Ist 52 KW ggü. Vorjahr (t)", title_font=dict(size=12, color=GRAU),
                     tickformat="+,.0f", gridcolor=ZUSATZ, zeroline=True, zerolinecolor=GRAU)
    fig.update_yaxes(type="category", gridcolor="rgba(0,0,0,0)")
    return fig


SEQUENZ = [[0, LILA], [0.35, HELL], [0.7, MITTEL], [1, NAVY]]
DIVERGIEREND = [[0, ROT], [0.25, "#DDA29E"], [0.5, "#F4F4F8"], [0.75, HELL], [1, NAVY]]
DIVERGIEREND_KARTE = [[0, ROT], [0.3, "#E2AAA6"], [0.5, "#C4C9D8"], [0.7, HELL], [1, NAVY]]
KARTE_LEER = "#EEF0F5"
LINIEN = [(NAVY, "solid"), (MITTEL, "solid"), (GRAU, "solid"), (NAVY, "dash"), (MITTEL, "dash"), (GRAU, "dash")]


def chart_kumuliert(W, plan, prog, vj, hoehe=380):
    """Kumulierter Verlauf ueber den Planhorizont mit Endwerten am rechten Rand."""
    x = [kw_label(w) for w in W]
    werte = {n: np.cumsum(np.asarray(y, float)) / T for n, y in (("Vorjahr", vj), ("Modellprognose", prog),
                                                                   ("Plan", plan))}
    fig = go.Figure()
    stil = {"Vorjahr": (GRAU, "dot", 1.8), "Modellprognose": (MITTEL, "dash", 2.2), "Plan": (NAVY, "solid", 3.2)}
    for name, (farbe, strich, breite) in stil.items():
        fig.add_trace(go.Scatter(x=x, y=werte[name], name=name, line=dict(color=farbe, width=breite, dash=strich),
                                 hovertemplate="%{y:,.0f} t"))
    # Endwerte rechts; Modell entfaellt, wenn deckungsgleich mit Plan; Abstand gegen Ueberlappung
    ende = {n: werte[n][-1] for n in stil}
    spanne = max(max(ende.values()), 1)
    etiketten = [n for n in ("Plan", "Vorjahr", "Modellprognose")
                 if not (n == "Modellprognose" and abs(ende[n] - ende["Plan"]) < 0.01 * spanne)]
    vorher = None
    for n in sorted(etiketten, key=lambda n: -ende[n]):
        pos = ende[n] / spanne * (hoehe - 80)          # ungefaehre Pixelhoehe
        ziel = pos if vorher is None else min(pos, vorher - 16)
        shift, vorher = ziel - pos, ziel
        fig.add_annotation(x=x[-1], y=ende[n], text=f"{'<b>' if n == 'Plan' else ''}{de(ende[n])} t",
                           showarrow=False, xanchor="left", xshift=8, yshift=shift,
                           font=dict(color=stil[n][0], size=12))
    fig = layout(fig, hoehe, y_titel="t kumuliert")
    fig.update_layout(margin=dict(l=8, r=80, t=36, b=8))
    fig.update_xaxes(nticks=9)
    return fig


def chart_saison(ist_w: pd.Series, plan_w: pd.Series, hoehe=360):
    """Saisonvergleich: Ø t je KW je Kalendermonat, eine Linie je Jahr, Planmonate gestrichelt.
    Mittel je KW statt Monatssumme, damit Monate mit 4 und 5 KW vergleichbar bleiben."""
    df = pd.DataFrame({"v": np.r_[ist_w.to_numpy(float), plan_w.to_numpy(float)],
                       "plan": np.r_[np.zeros(len(ist_w)), np.ones(len(plan_w))]},
                      index=list(ist_w.index) + list(plan_w.index))
    df["mon"] = [monat_key(w) for w in df.index]
    g = df.groupby("mon").agg(v=("v", "mean"), n=("v", "size"), plan=("plan", "mean"))
    g = g[g["n"] >= 3].copy()
    g["jahr"] = g.index.str[:4].astype(int)
    g["m"] = g.index.str[5:].astype(int)
    jahre = [j for j, k in g.groupby("jahr").size().items() if k >= 3]
    ist_j = sorted(g.loc[g["plan"] < 0.5, "jahr"].unique())
    akt = ist_j[-1] if ist_j else max(jahre)
    jahre = [j for j in jahre if j >= akt - 3]
    farben = {akt: NAVY, akt - 1: MITTEL, akt - 2: HELL, akt - 3: HELLGRAU}
    fig = go.Figure()
    plan_leg = False
    for j in sorted(jahre):
        d = g[g["jahr"] == j].sort_values("m")
        ist, pl = d[d["plan"] < 0.5], d[d["plan"] >= 0.5]
        if len(ist):
            fig.add_trace(go.Scatter(
                x=[MONATE[m - 1] for m in ist["m"]], y=ist["v"] / T, name=str(j),
                mode="lines+markers" if j == akt else "lines", marker=dict(size=6),
                line=dict(color=farben.get(j, HELLGRAU), width=3.2 if j == akt else 2.2),
                hovertemplate=f"{j}: %{{y:,.1f}} t"))
        if len(pl):
            if len(ist) and ist["m"].iloc[-1] == pl["m"].iloc[0] - 1:
                pl = pd.concat([ist.iloc[[-1]], pl])
            fig.add_trace(go.Scatter(
                x=[MONATE[m - 1] for m in pl["m"]], y=pl["v"] / T, name="Plan", legendgroup="plan",
                showlegend=not plan_leg, mode="lines+markers", marker=dict(size=6, symbol="circle-open"),
                line=dict(color=NAVY, width=3, dash="dash"), hovertemplate=f"Plan {j}: %{{y:,.1f}} t"))
            plan_leg = True
    fig = layout(fig, hoehe, y_titel="Ø t je KW")
    fig.update_xaxes(type="category", categoryorder="array", categoryarray=MONATE)
    fig.update_yaxes(tickformat=",.0f" if g["v"].max() / T >= 20 else ",.1f")
    return fig


def chart_zeitverlauf(daten: pd.DataFrame, monate: list, keys: list, labels: dict, hoehe=340, index=False):
    """Monatsverlauf (Kalendermonate) fuer ausgewaehlte Materialien, Laender oder Kunden.
    index=True: jede Linie relativ zu ihrem 24-Monats-Mittel (= 100), macht kleine Positionen vergleichbar."""
    piv = (daten.pivot_table(index="Schluessel", columns="Monat", values="Menge_kg", aggfunc="sum")
           .reindex(columns=monate).fillna(0.0))
    if index:
        piv = piv.div(piv.mean(axis=1).replace(0, np.nan), axis=0) * 100 * T
    x = [monat_label(m) for m in monate]
    fig = go.Figure()
    for i, k in enumerate(keys):
        if k not in piv.index:
            continue
        farbe, strich = LINIEN[i % len(LINIEN)]
        fig.add_trace(go.Scatter(x=x, y=piv.loc[k].to_numpy() / T, name=labels.get(k, k), mode="lines",
                                 line=dict(color=farbe, width=2.6 if i == 0 else 2.2, dash=strich),
                                 hovertemplate="%{y:,.1f} t"))
    fig = layout(fig, hoehe, y_titel="Index (Ø 24 Monate = 100)" if index else "t je Monat")
    fig.update_xaxes(type="category", nticks=12)
    if index:
        fig.update_traces(hovertemplate="%{y:,.0f}")
        fig.add_hline(y=100, line=dict(color=HELLGRAU, width=1, dash="dot"))
        fig.update_yaxes(tickformat=",.0f")
    else:
        fig.update_yaxes(tickformat=",.1f" if piv.loc[[k for k in keys if k in piv.index]].to_numpy().max(initial=0) / T
                         < 20 else ",.0f")
    return fig


def chart_heatmap(D, plan_m: pd.DataFrame, prog_m: pd.DataFrame):
    """Abweichung Plan zu Modell je MB und Monat."""
    mbs = D["mbs"]
    spalten = [monat_label(m, D["n_kw"][m]) for m in D["monate"]]
    pm, fm = plan_m.loc[mbs].to_numpy(float), prog_m.loc[mbs].to_numpy(float)
    with np.errstate(divide="ignore", invalid="ignore"):
        z = np.where(fm > 0, (pm / fm - 1) * 100, np.nan)
    gross = np.nanmax(np.abs(z)) if np.isfinite(z).any() else 0
    lim = float(min(max(5.0, np.ceil(gross / 5) * 5), 50.0))
    text = [["" if not np.isfinite(v) else ("0 %" if abs(v) < 0.5 else f"{v:+.0f} %") for v in zeile] for zeile in z]
    fig = go.Figure(go.Heatmap(
        z=np.clip(z, -lim, lim), x=spalten, y=[f"MB {m}" for m in mbs], text=text, texttemplate="%{text}",
        textfont=dict(size=12, family=FONT), customdata=np.dstack([pm / T, fm / T]), colorscale=DIVERGIEREND,
        zmid=0, zmin=-lim, zmax=lim, xgap=3, ygap=3,
        hovertemplate="%{y}, %{x}<br>Plan %{customdata[0]:,.1f} t<br>Modell %{customdata[1]:,.1f} t"
                      "<br>Abweichung %{text}<extra></extra>",
        colorbar=dict(thickness=10, len=0.9, ticksuffix=" %", outlinewidth=0, tickfont=dict(size=11))))
    fig = layout(fig, 60 + 36 * len(mbs), y_titel="", legende=False)
    fig.update_layout(hovermode="closest", margin=dict(l=8, r=8, t=30, b=8))
    fig.update_yaxes(autorange="reversed", type="category", gridcolor="rgba(0,0,0,0)")
    fig.update_xaxes(side="top", ticks="", showline=False)
    return fig


# ---------------------------------------------------------------------------
# Laenderkarte: Grenzen aus Natural Earth 1:50 Mio (gemeinfrei), eingebettet am Dateiende.
# Gezeichnet als Flaechen in einem normalen Diagramm, damit Plotly keine Kartendaten aus dem
# Internet nachlaedt. Europa in flaechentreuer Azimutalprojektion (wie ETRS89-LAEA), Welt in
# Natural Earth Projektion.
# ---------------------------------------------------------------------------
def _laea(lon, lat, lon0=10.0, lat0=52.0):
    lam, phi = np.radians(lon), np.radians(lat)
    l0, p0 = np.radians(lon0), np.radians(lat0)
    k = np.sqrt(2 / np.maximum(1 + np.sin(p0) * np.sin(phi) + np.cos(p0) * np.cos(phi) * np.cos(lam - l0), 1e-9))
    return (k * np.cos(phi) * np.sin(lam - l0),
            k * (np.cos(p0) * np.sin(phi) - np.sin(p0) * np.cos(phi) * np.cos(lam - l0)))


def _natural_earth(lon, lat):
    lam, phi = np.radians(lon), np.radians(lat)
    p2 = phi * phi
    p4 = p2 * p2
    x = lam * (0.8707 - 0.131979 * p2 + p4 * (-0.013791 + p4 * (0.003971 * p2 - 0.001529 * p4)))
    y = phi * (1.007226 + p2 * (0.015085 + p4 * (-0.044475 + 0.028874 * p2 - 0.005916 * p4)))
    return x, y


KARTE_RAHMEN = {
    "Europa": _laea(np.array([-10, -10, -24, 31, 44, 44, 35, 26]), np.array([36, 52, 64.5, 70.5, 57, 40, 34, 34.5])),
    "Welt": (np.array([-2.75, 2.75]), np.array([-1.0, 1.38])),
}


@st.cache_resource(show_spinner=False)
def weltkarte(ansicht: str) -> dict:
    gj = json.loads(gzip.decompress(base64.b64decode(WELT_B64)))
    out = {}
    for f in gj["features"]:
        geo = f["geometry"]
        polys = [geo["coordinates"]] if geo["type"] == "Polygon" else geo["coordinates"]
        xs, ys = [], []
        for poly in polys:
            for ring in poly:
                a = np.asarray(ring, float)
                if ansicht == "Europa":
                    if a[:, 0].max() < -35 or a[:, 0].min() > 65 or a[:, 1].max() < 25 or a[:, 1].min() > 82:
                        continue
                    x, y = _laea(a[:, 0], a[:, 1])
                else:
                    x, y = _natural_earth(a[:, 0], a[:, 1])
                # Ringe an der Datumsgrenze (Sprung > 180 Grad) auftrennen, sonst Linien quer ueber die Karte
                bruch = np.where(np.abs(np.diff(a[:, 0])) > 180)[0] + 1
                for tx, ty in zip(np.split(np.round(x, 4), bruch), np.split(np.round(y, 4), bruch)):
                    if len(tx) > 2:
                        xs += tx.tolist() + [None]
                        ys += ty.tolist() + [None]
        if xs:
            out[f["id"]] = (xs, ys)
    return out


def laender_werte(det: pd.DataFrame) -> pd.DataFrame:
    g = det.groupby("Land").agg(Landname=("Landname", "first"), Ist=("Ist_52W_kg", "sum"),
                                VJ=("Ist_VJ_52W_kg", "sum")).reset_index()
    g["Land"] = g["Land"].astype(str).str.strip().str.upper()
    return g[(g["Ist"] > 0) | (g["VJ"] > 0)]


def _nette_ticks(maximum: float) -> list:
    if maximum <= 0:
        return [0.0]
    basis = 10 ** np.floor(np.log10(maximum))
    kandidaten = [f * basis * 10 ** e for e in (-2, -1, 0) for f in (1, 2, 5)]
    t = sorted({0.0, *[k for k in kandidaten if k <= maximum]})
    return t[::max(1, len(t) // 4)][-5:] if len(t) > 5 else t


def chart_karte(werte: pd.DataFrame, metrik: str, ansicht: str, hoehe=460):
    geo = weltkarte(ansicht)
    w = werte[werte["Land"].isin(geo.keys())].copy()
    mit_daten = set(w["Land"])
    fig = go.Figure()
    bx, by = [], []
    for k, (xs, ys) in geo.items():
        if k not in mit_daten:
            bx += xs
            by += ys
    fig.add_trace(go.Scatter(x=bx, y=by, fill="toself", fillcolor=KARTE_LEER, mode="lines",
                             line=dict(color="#D9DCE6", width=0.5), hoverinfo="skip", showlegend=False))
    with np.errstate(divide="ignore", invalid="ignore"):
        w["pct"] = np.where(w["VJ"] > 0, (w["Ist"] / w["VJ"] - 1) * 100, np.nan)
    if metrik == "Menge":
        v = np.sqrt(w["Ist"].clip(lower=0) / T)
        cmax = float(v.max()) if len(v) and v.max() > 0 else 1.0
        anteil = (v / cmax).fillna(0).clip(0, 1)
        skala, cmin = SEQUENZ, 0.0
        tv_t = _nette_ticks(cmax ** 2)
        tickvals, ticktext, titel = [np.sqrt(t) for t in tv_t], [de(t) for t in tv_t], "t"
    else:
        roh = w["pct"].copy()
        lim = float(np.clip(np.ceil(np.nanpercentile(np.abs(roh), 90) / 5) * 5 if roh.notna().any() else 10, 10, 50))
        wert = roh.fillna(lim).clip(-lim, lim)      # neu (ohne Vorjahr) = oberes Ende
        anteil = (wert + lim) / (2 * lim)
        skala, cmin, cmax = DIVERGIEREND_KARTE, -lim, lim
        tickvals = [-lim, -lim / 2, 0, lim / 2, lim]
        ticktext = [f"{t:+.0f} %".replace("+0 %", "0 %") for t in tickvals]
        titel = "ggü. VJ"
    farben = sample_colorscale(skala, anteil.tolist()) if len(w) else []
    for (_, r), farbe in zip(w.iterrows(), farben):
        xs, ys = geo[r["Land"]]
        veraend = "neu" if not np.isfinite(r["pct"]) else pct(r["pct"] / 100)
        txt = (f"<b>{h(r['Landname'])}</b><br>Ist 52 KW {de(r['Ist'] / T, 1)} t<br>"
               f"Vorjahr {de(r['VJ'] / T, 1)} t<br>Veränderung {veraend}")
        fig.add_trace(go.Scatter(x=xs, y=ys, fill="toself", fillcolor=farbe, mode="lines", hoveron="fills",
                                 line=dict(color="#FFFFFF", width=0.7), name=txt, hoverinfo="name",
                                 showlegend=False))
    rx, ry = KARTE_RAHMEN[ansicht]
    fig.add_trace(go.Scatter(x=[float(rx.min())], y=[float(ry.min())], mode="markers", hoverinfo="skip",
                             showlegend=False, marker=dict(size=0.1, opacity=0, color=[cmin], cmin=cmin, cmax=cmax,
                                                           colorscale=skala, showscale=True, colorbar=dict(
                                 thickness=10, len=0.32 if ansicht == "Europa" else 0.55,
                                 x=0.015 if ansicht == "Europa" else 0.99,
                                 xanchor="left" if ansicht == "Europa" else "right",
                                 y=0.6 if ansicht == "Europa" else 0.45, yanchor="middle", outlinewidth=0,
                                 tickvals=tickvals, ticktext=ticktext, tickfont=dict(size=11, color=NAVY),
                                 title=dict(text=titel, side="top", font=dict(size=11, color=GRAU)),
                                 bgcolor="rgba(0,0,0,0)"))))
    pad_x, pad_y = (rx.max() - rx.min()) * 0.02, (ry.max() - ry.min()) * 0.02
    fig.update_layout(height=hoehe if ansicht == "Europa" else int(hoehe * 0.62), margin=dict(l=0, r=0, t=0, b=0),
                      paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF",
                      hovermode="closest", dragmode="pan", font=dict(family=FONT, color=NAVY),
                      hoverlabel=dict(bgcolor="#FFFFFF", bordercolor=HELL, namelength=-1,
                                      font=dict(family=FONT, color=NAVY, size=13)))
    fig.update_xaxes(visible=False, range=[rx.min() - pad_x, rx.max() + pad_x])
    fig.update_yaxes(visible=False, range=[ry.min() - pad_y, ry.max() + pad_y], scaleanchor="x", scaleratio=1)
    return fig


KARTE_CFG = {**PLOT_CFG, "scrollZoom": True,
             "modeBarButtonsToRemove": ["select2d", "lasso2d", "autoScale2d", "zoom2d"]}


def karte_block(werte: pd.DataFrame, key: str, hoehe=460):
    """Karte mit Umschaltern fuer Kennzahl und Ausschnitt; Hinweis auf nicht darstellbare Laender."""
    if werte.empty:
        st.info("Keine Länderdaten vorhanden.")
        return
    geo_codes = set(weltkarte("Welt").keys())
    europa = set(weltkarte("Europa").keys())
    anteil_eu = werte.loc[werte["Land"].isin(europa) & ~werte["Land"].isin({"US", "CA", "CN", "IN", "BR"}),
                          "Ist"].sum() / max(werte["Ist"].sum(), 1)
    c1, c2 = st.columns(2)
    metrik = c1.segmented_control("Kennzahl", ["Menge", "Veränderung"], default="Menge", key=f"{key}_m",
                                  label_visibility="collapsed") or "Menge"
    ansicht = c2.segmented_control("Ausschnitt", ["Europa", "Welt"], default="Europa" if anteil_eu >= 0.6 else "Welt",
                                   key=f"{key}_a", label_visibility="collapsed") or "Europa"
    st.plotly_chart(chart_karte(werte, metrik, ansicht, hoehe), config=KARTE_CFG, key=f"{key}_fig")
    fehlt = werte[~werte["Land"].isin(geo_codes)]
    zusatz = ""
    if len(fehlt):
        zusatz = (" Ohne Kartenfläche: " + ", ".join(f"{h(r['Landname'])} ({de(r['Ist'] / T, 1)} t)"
                                                      for _, r in fehlt.sort_values("Ist", ascending=False).head(6)
                                                      .iterrows()) + ".")
    st.caption("Ist der letzten 52 KW je Land Warenempfänger; Veränderung gegenüber den 52 KW davor. "
               "Mausrad zoomt, Ziehen verschiebt, Doppelklick setzt zurück." + zusatz)


# ============================================================================
# BAUSTEINE
# ============================================================================
SICHT_HINWEIS = {
    "Wochen": "Ist je KW, Modellprognose mit 80-%-Bandbreite und Plan; hinterlegt ist der Planhorizont. "
              "Zeitraum mit den beiden Griffen wählen, „Gesamt“ zeigt den kompletten Datensatz.",
    "Monate": "Monatssummen im Planhorizont nach KW-Zuordnung (Donnerstag der KW) im Vergleich zum Vorjahr.",
    "Kumuliert": "Aufsummierte Menge über den Planhorizont; rechts der Endwert nach 26 KW.",
    "Saisonvergleich": "Durchschnittliche Menge je KW pro Kalendermonat, eine Linie je Jahr; Planmonate gestrichelt.",
}
SICHTEN = ["Wochen", "Monate", "Kumuliert", "Saisonvergleich"]


def summen(D, mbs) -> dict:
    """Zeitreihen fuer eine Auswahl von Marktbereichen (alle MB: Bandbreite aus dem Backtest der Summe)."""
    alle = len(mbs) == len(D["mbs"])
    return dict(ist=D["ist"].loc[mbs].sum(), prog=D["prog"].loc[mbs].sum(), plan=ss["plan"].loc[mbs].sum(),
                vj=D["vj"].loc[mbs].sum(),
                bu=D["ges"]["Band_unten_kg"] if alle else D["bu"].loc[mbs].sum(),
                bo=D["ges"]["Band_oben_kg"] if alle else D["bo"].loc[mbs].sum())


def verlauf_chart(D, sicht, mbs, zeitraum=None, hoehe=400):
    r = summen(D, mbs)
    if sicht == "Monate":
        return chart_monate(D, plan_monat(D).loc[mbs].sum().values, wert_monat(D, D["prog"]).loc[mbs].sum().values,
                            wert_monat(D, D["vj"]).loc[mbs].sum().values, hoehe=hoehe)
    if sicht == "Kumuliert":
        return chart_kumuliert(D["W"], r["plan"].values, r["prog"].values, r["vj"].values, hoehe=hoehe)
    if sicht == "Saisonvergleich":
        return chart_saison(r["ist"], r["plan"], hoehe=hoehe)
    return chart_wochen(r["ist"], r["prog"], r["bu"], r["bo"], r["plan"], hoehe=hoehe, zeitraum=zeitraum)


def verlauf_karte(D, mbs, key, titel):
    with st.container(border=True):
        c1, c2 = st.columns([2, 3], vertical_alignment="center")
        c1.markdown(f"### {titel}")
        sicht = c2.segmented_control("Sicht", SICHTEN, default="Wochen", required=True, key=f"{key}_sicht",
                                     label_visibility="collapsed") or "Wochen"
        zr = zeitstrahl(D, f"{key}_zeit") if sicht == "Wochen" else None
        st.plotly_chart(verlauf_chart(D, sicht, mbs, zr), config=PLOT_CFG, key=f"{key}_fig")
        st.caption(SICHT_HINWEIS[sicht])


def runde_titel(D) -> str:
    rm = next((m for m in D["monate"] if D["n_kw"][m] >= 4), D["monate"][0])
    return f"Bedarfsplanung {MONATE_LANG[int(rm[5:]) - 1]} {rm[:4]}"


def mb_text(mbs) -> str:
    return ", ".join(f"MB {h(m)}" for m in mbs)


def freigabe_hinweis(D):
    """Hinweis fuer die Freigabe-Ansicht, sobald Einreichungen warten."""
    offen = mbs_mit_status("Eingereicht")
    if ansicht() != ANSICHT_FREIGABE or not offen:
        return
    c1, c2 = st.columns([6, 1.6], vertical_alignment="center")
    c1.markdown(f"<div class='wp-banner'><b>{len(offen)} {'Einreichung wartet' if len(offen) == 1 else 'Einreichungen warten'}"
                f"</b> auf Ihre Freigabe: {mb_text(offen)}.</div>", unsafe_allow_html=True)
    if c2.button("Zur Freigabe", type="primary", icon=":material/fact_check:", width="stretch", key="banner_fg"):
        st.switch_page(SEITEN["freigabe"])


def kernaussagen(D, mbs) -> list:
    st_ = ss["status"]
    plan, prog = ss["plan"].loc[mbs], D["prog"].loc[mbs]
    p6, f6, vj6 = plan.values.sum(), prog.values.sum(), D["vj"].loc[mbs].values.sum()
    s = []
    zu_modell = p6 / f6 - 1 if f6 else np.nan
    s.append(f"Plan <b>{de(p6 / T)} t</b> für die nächsten 6 Monate, <b>{pct(p6 / vj6 - 1) if vj6 else 'n/a'}</b> "
             f"ggü. Vorjahr; " + ("entspricht der Modellprognose." if abs(zu_modell) < 0.0005
                                  else f"{pct(zu_modell)} ggü. Modellprognose."))
    g, gv = genauigkeit(D, mbs)
    if np.isfinite(g):
        bewertung = {"hoch": "Zahlen belastbar", "mittel": "Zahlen eingeschränkt belastbar",
                     "gering": "Zahlen nur mit Marktwissen verwenden"}.get(guete_label(g), "")
        s.append(f"Prognosegenauigkeit im Backtest <b>{pct(g, False, 0)}</b> (Vorjahresmethode {pct(gv, False, 0)}); "
                 f"{bewertung}.")
    schwach = [mb for mb in mbs if guete_label(genauigkeit_mb(D, mb)) == "gering"
               and str(D["mbinfo"].loc[mb, "Verlaesslichkeit"]) != "nicht steuerungsrelevant"]
    drift = [mb for mb in mbs if str(D["mbinfo"].loc[mb, "Drift"]).lower() in ("true", "1", "wahr")]
    if schwach or drift:
        teile = []
        if schwach:
            teile.append(f"geringe Prognosegüte in {mb_text(schwach)}")
        if drift:
            teile.append(f"systematische Abweichung zuletzt in {mb_text(drift)}")
        s.append("Kritisch prüfen: " + "; ".join(teile) + ".")
    eingriffe = (plan.sum(axis=1) - prog.sum(axis=1)).abs().sort_values(ascending=False)
    if len(eingriffe) and eingriffe.iloc[0] > 0.005 * max(f6, 1):
        mb = eingriffe.index[0]
        grund = str(st_.loc[mb, "Begruendung"]).strip()
        s.append(f"Größter Eingriff: <b>MB {h(mb)}</b> mit {pct(abweichung(D, mb))} ggü. Modell"
                 + (f" ({h(grund[:90])})." if grund else ", noch ohne Begründung."))
    frei = sum(st_.loc[m, "Status"] == "Freigegeben" for m in mbs)
    offen = sum(st_.loc[m, "Status"] == "Eingereicht" for m in mbs)
    s.append(f"Freigabestand: <b>{frei} von {len(mbs)}</b> Marktbereichen freigegeben"
             + (f", {offen} {'Einreichung' if offen == 1 else 'Einreichungen'} offen." if offen else "."))
    return s


def kernaussagen_karte(D, mbs):
    saetze = kernaussagen(D, mbs)
    inhalt = "".join(f"<p class='wp-satz'>{x}</p>" for x in saetze)
    st.markdown(f"<div class='wp-kern'><div class='wp-kern-titel'>Kernaussagen</div>"
                f"<div class='wp-kern-grid'>{inhalt}</div></div>", unsafe_allow_html=True)


def kennzahlen(D, mbs, freigabe=True):
    plan, prog = ss["plan"].loc[mbs], D["prog"].loc[mbs]
    p6, f6, vj6 = plan.values.sum(), prog.values.sum(), D["vj"].loc[mbs].values.sum()
    l26 = D["ist"].loc[mbs].iloc[:, -26:].values.sum()
    g, gv = genauigkeit(D, mbs)
    k = st.columns(5 if freigabe else 4)
    k[0].metric("Plan 6 Monate", f"{de(p6 / T)} t", f"Modell {pct(p6 / f6 - 1) if f6 else 'n/a'}",
                border=True, delta_color="off", delta_arrow="off")
    k[1].metric("Ggü. Vorjahr", pct(p6 / vj6 - 1) if vj6 else "n/a", f"Vorjahr {de(vj6 / T)} t", border=True,
                delta_color="off", delta_arrow="off")
    k[2].metric("Ggü. letzte 26 KW", pct(p6 / l26 - 1) if l26 else "n/a", f"Ist {de(l26 / T)} t", border=True,
                delta_color="off", delta_arrow="off")
    k[3].metric("Genauigkeit", pct(g, False, 0) if np.isfinite(g) else "n/a",
                f"Güte {guete_label(g)}", border=True,
                delta_color="off", delta_arrow="off", help=GUETE_HILFE)
    if freigabe:
        st_ = ss["status"]
        frei = [m for m in mbs if st_.loc[m, "Status"] == "Freigegeben"]
        anteil = prog.loc[frei].values.sum() / f6 if f6 else 0
        k[4].metric("Freigegeben", f"{len(frei)} von {len(mbs)}", f"{pct(anteil, False, 0)} der Menge", border=True,
                    delta_color="off", delta_arrow="off")


# ============================================================================
# SEITE: START
# ============================================================================
def seite_start():
    st.markdown(
        "<div class='wp-start'><h1>Bedarfsplanung Marktbereiche</h1>"
        "<p class='lead'>Ein abgestimmter Bedarfsplan für die nächsten sechs Monate, einmal im Monat. "
        "Die Modellprognose ist der Startpunkt, Sonderwissen aus dem Markt wird begründet eingearbeitet, "
        "die Freigabe macht den Plan verbindlich.</p>"
        "<div class='wp-schritte'>"
        "<div><b>1  Prognose laden</b><span>Ergebnis-Excel des Prognoseskripts hochladen, optional den "
        "letzten Planstand dazu.</span></div>"
        "<div><b>2  Prüfen und anpassen</b><span>Übersicht und Detail je Marktbereich bis auf Produkt, Land "
        "und Kunde; Modellwerte bei Bedarf begründet überschreiben.</span></div>"
        "<div><b>3  Freigeben</b><span>Änderungen genehmigen oder ablehnen, finalen Bericht mit "
        "Änderungsprotokoll herunterladen.</span></div>"
        "</div></div>", unsafe_allow_html=True)
    c1, c2 = st.columns([3, 2], gap="large")
    with c1:
        with st.container(border=True):
            st.markdown("### Prognose hochladen")
            datei = st.file_uploader("Ergebnis-Excel des Prognoseskripts", type=["xlsx"], key="up_start")
            if datei is not None:
                lade_prognose(datei.getvalue(), datei.name)
            st.caption("Die Datei bleibt im Arbeitsspeicher dieser Sitzung. Nichts wird gespeichert.")
    with c2:
        with st.container(border=True):
            st.markdown("### Ohne Kundendaten testen")
            st.markdown("<div class='wp-hinweis'>Synthetische Daten mit acht Marktbereichen, um die Bedienung "
                        "zu zeigen.</div>", unsafe_allow_html=True)
            st.space("small")
            if st.button("Demodaten laden", icon=":material/science:", width="stretch"):
                ss["D"] = demo_daten()
                ss["quelle"] = "Demodaten"
                runde_starten(ss["D"])
                st.switch_page(SEITEN["uebersicht"])


def lade_prognose(inhalt: bytes, name: str):
    try:
        D = lese_prognose(inhalt)
    except Exception as e:
        st.error(str(e) if isinstance(e, ValueError) else f"Datei konnte nicht gelesen werden: {e}")
        return
    if ss.get("D") is None or lauf_id(D) != ss.get("lauf"):
        ss["D"], ss["quelle"] = D, name
        runde_starten(D)
        st.switch_page(SEITEN["uebersicht"])


# ============================================================================
# SEITE: UEBERSICHT
# ============================================================================
def mb_tabelle(D) -> pd.DataFrame:
    st_, plan, prog, mat = ss["status"], ss["plan"], D["prog"], D["mat"]
    f6 = prog.values.sum()
    top = {mb: g.nlargest(5, "Ist_52W_kg")["Material"].tolist() for mb, g in mat.groupby("MB")}
    return pd.DataFrame({
        "MB": D["mbs"],
        "Status": [st_.loc[mb, "Status"] + (" (geändert)" if geaendert(D, mb) and st_.loc[mb, "Status"] == "Offen"
                                            else "") for mb in D["mbs"]],
        "Ist 52 KW": [(D["ist"].loc[mb].iloc[-52:] / T).round(1).tolist() for mb in D["mbs"]],
        "Plan 6 M (t)": (plan.loc[D["mbs"]].sum(axis=1).values / T).round(0),
        "zu Modell": [abweichung(D, mb) * 100 for mb in D["mbs"]],
        "zu Vorjahr": [(plan.loc[mb].sum() / D["vj"].loc[mb].sum() - 1) * 100 if D["vj"].loc[mb].sum() > 0
                       else np.nan for mb in D["mbs"]],
        "Anteil": (prog.loc[D["mbs"]].sum(axis=1) / f6 * 100).values,
        "Genauigkeit": [genauigkeit_mb(D, mb) * 100 for mb in D["mbs"]],
        "Top-Produkte": [top.get(mb, []) for mb in D["mbs"]],
    })


MB_TABELLE_CFG = {
    "MB": st.column_config.TextColumn(width=55),
    "Status": st.column_config.TextColumn(width=140),
    "Ist 52 KW": st.column_config.AreaChartColumn(width=150, color=MITTEL),
    "Plan 6 M (t)": st.column_config.NumberColumn(format="localized", width=85),
    "zu Modell": st.column_config.NumberColumn(format="%+.1f %%", width=80),
    "zu Vorjahr": st.column_config.NumberColumn(format="%+.1f %%", width=80),
    "Anteil": st.column_config.ProgressColumn(format="%.0f %%", min_value=0, max_value=100, color=NAVY, width=85),
    "Genauigkeit": st.column_config.NumberColumn(format="%.0f %%", width=85,
                                                 help="Prognosegenauigkeit im Backtest (100 % minus mittlere "
                                                      "Abweichung je 4 Wochen)"),
    "Top-Produkte": st.column_config.ListColumn(width="large"),
}


def seite_uebersicht():
    D = ss.get("D")
    if D is None:
        return keine_daten()
    key = f"ue_tab_{ss.get('tab_v', 0)}"
    if key in ss and ss[key] is not None:                         # Auswahl aus der Tabelle (vor dem Rendern)
        ss["auswahl"] = [D["mbs"][i] for i in ss[key]["selection"]["rows"] if i < len(D["mbs"])]
    sel = [m for m in ss.get("auswahl", []) if m in D["mbs"]]
    mbs = sel or D["mbs"]
    plan, vj6 = ss["plan"].loc[mbs], D["vj"].loc[mbs].values.sum()
    p6 = plan.values.sum()
    kopf(runde_titel(D), f"Plan {de(p6 / T)} t für 6 Monate, {pct(p6 / vj6 - 1) if vj6 else 'n/a'} ggü. Vorjahr"
         + (f" | Auswahl {', '.join('MB ' + m for m in sel)}" if sel else ""), D)
    freigabe_hinweis(D)
    c1, c2 = st.columns([6, 1.6], vertical_alignment="center")       # feste Zeile, damit nichts springt
    c1.markdown(f"<div class='wp-filter'>Gefiltert auf <b>{mb_text(sel)}</b>. Kennzahlen, Verlauf, Karte und "
                "Produkte zeigen nur diese Auswahl.</div>" if sel else
                "<div class='wp-filter'><b>Alle Marktbereiche.</b> In der Tabelle Marktbereiche anhaken, um die "
                "Seite zu filtern.</div>", unsafe_allow_html=True)
    if c2.button("Filter aufheben", icon=":material/filter_alt_off:", width="stretch", key="filter_weg",
                 disabled=not sel):
        ss["auswahl"] = []
        ss["tab_v"] = ss.get("tab_v", 0) + 1
        st.rerun()
    kernaussagen_karte(D, mbs)
    kennzahlen(D, mbs)
    verlauf_karte(D, mbs, "ue", "Absatzverlauf: Ist, Modellprognose und Plan")

    with st.container(border=True):
        if True:
            st.markdown("### Marktbereiche")
            st.caption("Haken setzen filtert Kennzahlen, Verlauf, Karte und Produkte auf die Auswahl.")
            st.dataframe(mb_tabelle(D), hide_index=True, on_select="rerun", selection_mode="multi-row", key=key,
                         row_height=36, column_config=MB_TABELLE_CFG,
                         selection_default={"selection": {"rows": [D["mbs"].index(m) for m in sel]}})
            c1, c2 = st.columns([1, 4], vertical_alignment="center")
            if c1.button("Im Detail öffnen", icon=":material/travel_explore:", width="stretch",
                         disabled=len(sel) != 1, key="ue_detail"):
                ss["mb_aktiv"] = sel[0]
                st.switch_page(SEITEN["detail"])
            c2.markdown("<span class='wp-hinweis'>Genau einen Marktbereich anhaken, um ihn im Detail zu öffnen."
                        "</span>" if len(sel) != 1 else "", unsafe_allow_html=True)
    links, rechts = st.columns([1, 1], gap="medium")
    with links:
        with st.container(border=True, height="stretch"):
            st.markdown("### Absatz nach Ländern" + (f": {', '.join('MB ' + m for m in sel)}" if sel else ""))
            karte_block(laender_werte(D["det"][D["det"]["MB"].isin(mbs)]), key="ue_karte", hoehe=440)
    with rechts:
        produkte_uebersicht(D, sel)
    rueckblick(D)


def produkt_tabelle(D, mat: pd.DataFrame, mit_mb=False) -> None:
    p6 = ss["plan"].sum(axis=1)
    t = pd.DataFrame({
        **({"MB": mat["MB"].values} if mit_mb else {}),
        "Produkt": mat["Material"].values,
        "Anteil im MB": mat["Anteil"].values * 100,
        "Ist 52 KW (t)": mat["Ist_52W_kg"].values / T,
        "zu Vorjahr": (mat["Ist_52W_kg"] / mat["Ist_VJ_52W_kg"].replace(0, np.nan) - 1).values * 100,
        "Plan 6 M (t)": (mat["Anteil"] * mat["MB"].map(p6)).values / T,
    })
    f1 = st.column_config.NumberColumn(format="%.1f")
    st.dataframe(t, hide_index=True, row_height=34, column_config={
        "MB": st.column_config.TextColumn(width=50),
        "Anteil im MB": st.column_config.ProgressColumn(format="%.0f %%", min_value=0, max_value=100, color=MITTEL,
                                                        width=110),
        "Ist 52 KW (t)": f1, "Plan 6 M (t)": f1, "zu Vorjahr": st.column_config.NumberColumn(format="%+.1f %%"),
        })


def produkte_uebersicht(D, sel):
    with st.container(border=True, height="stretch"):
        st.markdown("### Wichtigste Produkte")
        if not sel:
            st.caption("Die zehn absatzstärksten Produkte der letzten 52 KW. Marktbereiche anhaken, um je "
                       "Marktbereich die fünf wichtigsten Produkte aufzuklappen.")
            produkt_tabelle(D, D["mat"].nlargest(10, "Ist_52W_kg"), mit_mb=True)
            return
        for i, mb in enumerate(sel):
            mat = D["mat"][D["mat"]["MB"] == mb].nlargest(5, "Ist_52W_kg")
            with st.expander(f"MB {mb}: fünf wichtigste Produkte", expanded=i == 0):
                produkt_tabelle(D, mat)
                if st.button(f"MB {mb} im Detail öffnen", key=f"prod_det_{mb}", icon=":material/travel_explore:"):
                    ss["mb_aktiv"] = mb
                    st.switch_page(SEITEN["detail"])


# ============================================================================
# SEITE: MARKTBEREICH IM DETAIL
# ============================================================================
def seite_detail():
    D = ss.get("D")
    if D is None:
        return keine_daten()
    st_, mbs = ss["status"], D["mbs"]
    if ss.get("mb_aktiv") not in mbs:
        ss["mb_aktiv"] = (ss.get("auswahl") or mbs)[0]
    mb = ss["mb_aktiv"]
    p6, f6, vj6 = ss["plan"].loc[mb].sum(), D["prog"].loc[mb].sum(), D["vj"].loc[mb].sum()
    kopf(f"Marktbereich {mb}", f"Plan {de(p6 / T)} t, {pct(p6 / f6 - 1) if f6 else 'n/a'} zum Modell, "
         f"{pct(p6 / vj6 - 1) if vj6 else 'n/a'} ggü. Vorjahr", D)
    freigabe_hinweis(D)
    c1, c2 = st.columns([2, 5], vertical_alignment="center")
    neu = c1.selectbox("Marktbereich", mbs, index=mbs.index(mb), label_visibility="collapsed",
                       format_func=lambda m: f"MB {m}  |  {st_.loc[m, 'Status']}")
    if neu != mb:
        ss["mb_aktiv"] = neu
        st.rerun()
    zusatz = ""
    if str(st_.loc[mb, "Kommentar_Freigabe"]).strip():
        zusatz = f"; Freigabe: {h(st_.loc[mb, 'Kommentar_Freigabe'])}"
    c2.markdown(f"{status_pill(st_.loc[mb, 'Status'])} &nbsp; <span class='wp-hinweis'>"
                f"{'Plan weicht vom Modell ab' if geaendert(D, mb) else 'Plan entspricht dem Modell'}"
                f"{'; zuletzt geändert ' + h(st_.loc[mb, 'Geaendert_am']) if st_.loc[mb, 'Geaendert_am'] else ''}"
                f"{zusatz}</span>", unsafe_allow_html=True)
    kennzahlen(D, [mb], freigabe=False)
    tabs = st.tabs(["Plan und Prognose", "Produkte", "Länder", "Kunden", "Prognosegüte"], key="de_tabs",
                   on_change="rerun")
    mat = D["mat"][D["mat"]["MB"] == mb].copy()
    det = D["det"][D["det"]["MB"] == mb].copy()
    if tabs[0].open:
        with tabs[0]:
            tab_plan(D, mb)
    if tabs[1].open:
        with tabs[1]:
            tab_produkte(D, mb, mat)
    if tabs[2].open:
        with tabs[2]:
            tab_dimension(D, det, mat, "Landname", "Land", mb)
    if tabs[3].open:
        with tabs[3]:
            tab_dimension(D, det, mat, "Auftraggeber", "Kunde", mb)
    if tabs[4].open:
        with tabs[4]:
            tab_guete(D, mb)


def tab_plan(D, mb):
    verlauf_karte(D, [mb], "de", "Absatzverlauf: Ist, Modellprognose und Plan")
    st_ = ss["status"]
    darf = darf_bearbeiten(mb)
    with st.container(border=True):
        st.markdown("### Planwerte je Monat")
        s = st_.loc[mb, "Status"]
        if ansicht() == ANSICHT_FREIGABE:
            st.caption("Ansicht Freigabe: Planwerte nur lesbar. Genehmigen oder ablehnen unter „Freigabe und Bericht“.")
        elif s == "Eingereicht":
            st.caption(f"Eingereicht am {st_.loc[mb, 'Eingereicht_am']}, wartet auf Freigabe. Änderungen erst nach "
                       "Entscheidung wieder möglich.")
        elif s == "Freigegeben":
            st.caption(f"Freigegeben am {st_.loc[mb, 'Freigegeben_am']} von {st_.loc[mb, 'Freigegeben_von']}.")
        elif s == "Abgelehnt":
            st.warning(f"Abgelehnt: {st_.loc[mb, 'Kommentar_Freigabe'] or 'ohne Kommentar'}. Bitte anpassen und erneut "
                       "einreichen.", icon=":material/undo:")
        else:
            st.caption(f"Modellwerte bei Sonderwissen in der Spalte „Plan“ überschreiben (Tonnen). Ab "
                       f"{pct(SCHWELLE, False, 0)} Abweichung zum Modell ist eine Begründung Pflicht.")
        pm, fm, vm = plan_monat(D).loc[mb], wert_monat(D, D["prog"]).loc[mb], wert_monat(D, D["vj"]).loc[mb]
        df = pd.DataFrame({"Monat": [monat_label(m, D["n_kw"][m]) for m in D["monate"]],
                           "Vorjahr (t)": vm.values / T, "Modell (t)": fm.values / T,
                           "Plan (t)": (pm.values / T).round(1),
                           "Plan zu Modell": (pm.values / np.where(fm.values > 0, fm.values, np.nan) - 1) * 100})
        num = st.column_config.NumberColumn(format="%.1f", disabled=True)
        ed = st.data_editor(df, hide_index=True, disabled=not darf, key=f"ed_mon_{mb}_{ss['editor_v']}",
                            row_height=36, column_config={
                                "Monat": st.column_config.TextColumn(disabled=True),
                                "Vorjahr (t)": num, "Modell (t)": num,
                                "Plan (t)": st.column_config.NumberColumn(format="%.1f", min_value=0.0, step=0.1,
                                                                          disabled=not darf),
                                "Plan zu Modell": st.column_config.NumberColumn(format="%+.1f %%", disabled=True)})
        if darf:
            aend = False
            for i, m in enumerate(D["monate"]):
                v = ed.loc[i, "Plan (t)"]
                if pd.notna(v) and abs(float(v) - df.loc[i, "Plan (t)"]) > 0.049:
                    setze_monat(D, mb, m, float(v) * T)
                    aend = True
            if aend:
                st.rerun()
        p6, f6 = pm.sum(), fm.sum()
        st.markdown(f"<p class='wp-satz'>Summe 6 Monate: Modell <b>{de(f6 / T, 1)} t</b>, Plan <b>{de(p6 / T, 1)} t"
                    f"</b> ({pct(p6 / f6 - 1) if f6 else 'n/a'}).</p>", unsafe_allow_html=True)
        grund = st.text_area("Begründung (Sonderwissen)", value=str(st_.loc[mb, "Begruendung"]), disabled=not darf,
                             key=f"grund_{mb}_{ss['editor_v']}", height=80,
                             placeholder="Was weiß der Markt, was das Modell nicht weiß? Zum Beispiel Neukunde, "
                                         "Projektgeschäft, Preiserhöhung, Lageraufbau beim Kunden")
        if darf and grund.strip() != str(st_.loc[mb, "Begruendung"]).strip():
            st_.loc[mb, "Begruendung"] = grund.strip()
            protokolliere("Begründung", mb, "", None, None, grund.strip())
            bearbeitet(mb)
            st.rerun()
        if darf:
            fehlt = braucht_begruendung(D, mb)
            b1, b2, b3 = st.columns([1.4, 1.4, 3], vertical_alignment="center")
            if b1.button("Zur Freigabe einreichen", type="primary", icon=":material/send:", width="stretch",
                         disabled=fehlt, key=f"einr_{mb}"):
                einreichen(D, mb)
                ss["editor_v"] += 1
                st.toast(f"MB {mb} eingereicht.", icon=":material/check_circle:")
                st.rerun()
            if b2.button("Auf Modell zurücksetzen", icon=":material/restart_alt:", width="stretch",
                         disabled=not geaendert(D, mb), key=f"reset_{mb}"):
                alt = ss["plan"].loc[mb].sum()
                ss["plan"].loc[mb] = D["prog"].loc[mb]
                protokolliere("Auf Modell zurückgesetzt", mb, "", alt, ss["plan"].loc[mb].sum())
                bearbeitet(mb)
                ss["editor_v"] += 1
                st.rerun()
            b3.markdown("<span class='wp-hinweis'>" + (
                f"Begründung erforderlich, Abweichung {pct(abweichung(D, mb))}." if fehlt else
                "Ohne Änderung bestätigt das Einreichen die Modellprognose.") + "</span>", unsafe_allow_html=True)
        with st.expander("Feinjustierung je KW"):
            wochen_editor(D, mb, darf)


def wochen_editor(D, mb, darf):
    W = D["W"]
    df = pd.DataFrame({"KW": [kw_label(w) for w in W], "Monat": [monat_label(D["monat"][w]) for w in W],
                       "Vorjahr": D["vj"].loc[mb, W].values / T, "Modell": D["prog"].loc[mb, W].values / T,
                       "Band unten": D["bu"].loc[mb, W].values / T, "Band oben": D["bo"].loc[mb, W].values / T,
                       "Plan": (ss["plan"].loc[mb, W].values / T).round(1)})
    num = st.column_config.NumberColumn(format="%.1f", disabled=True)
    ed = st.data_editor(df, hide_index=True, disabled=not darf, key=f"ed_kw_{mb}_{ss['editor_v']}", height=420,
                        row_height=32, column_config={
                            "KW": st.column_config.TextColumn(disabled=True), "Monat": st.column_config.TextColumn(
                                disabled=True), "Vorjahr": num, "Modell": num, "Band unten": num, "Band oben": num,
                            "Plan": st.column_config.NumberColumn("Plan (t)", format="%.1f", min_value=0.0, step=0.1,
                                                                  disabled=not darf)})
    if darf:
        aend = False
        for i, w in enumerate(W):
            v = ed.loc[i, "Plan"]
            if pd.notna(v) and abs(float(v) - df.loc[i, "Plan"]) > 0.049:
                setze_woche(D, mb, w, float(v) * T)
                aend = True
        if aend:
            st.rerun()
    st.caption("Werte in Tonnen. Band: 80 % der im Backtest beobachteten Abweichungen.")


def tab_produkte(D, mb, mat):
    if mat.empty:
        st.info("Keine Produktdaten für diesen Marktbereich.")
        return
    pm = plan_monat(D).loc[mb]
    mat["Delta_kg"] = mat["Ist_52W_kg"] - mat["Ist_VJ_52W_kg"]
    mat = mat.sort_values("Ist_52W_kg", ascending=False)
    n_80 = int((mat["Anteil"].sort_values(ascending=False).cumsum() < 0.8).sum()) + 1
    top, tief = mat.loc[mat["Delta_kg"].idxmax()], mat.loc[mat["Delta_kg"].idxmin()]
    st.markdown(f"<p class='wp-satz'><b>{n_80} von {len(mat)} Produkten</b> tragen 80 % des Plans. Stärkster Zuwachs "
                f"ggü. Vorjahr: <b>{h(top['Material'])}</b> ({de(top['Delta_kg'] / T, 1)} t), stärkster Rückgang: "
                f"<b>{h(tief['Material'])}</b> ({de(tief['Delta_kg'] / T, 1)} t).</p>", unsafe_allow_html=True)
    with st.container(border=True):
        c1, c2 = st.columns([3, 2], vertical_alignment="center")
        c1.markdown("**Prognose und Plan je Produkt**")
        such = c2.text_input("Produkt suchen", key=f"such_mat_{mb}", placeholder="Produktnummer suchen",
                             label_visibility="collapsed")
        t = mat[mat["Material"].str.contains(such.strip(), case=False, regex=False)] if such else mat
        out = pd.DataFrame({"Produkt": t["Material"].values, "Anteil": t["Anteil"].values * 100,
                            "Ist 52 KW": t["Ist_52W_kg"].values / T,
                            "zu Vorjahr": (t["Ist_52W_kg"] / t["Ist_VJ_52W_kg"].replace(0, np.nan) - 1).values * 100,
                            "Modell 6 M": t["Prognose_26W_kg"].values / T,
                            "Plan 6 M": t["Anteil"].values * pm.sum() / T})
        for m in D["monate"]:
            out[monat_label(m, D["n_kw"][m])] = t["Anteil"].values * pm[m] / T
        f1 = st.column_config.NumberColumn(format="%.1f")
        st.dataframe(out, hide_index=True, height=380, row_height=34, column_config={
            "Produkt": st.column_config.TextColumn(pinned=True),
            "Anteil": st.column_config.ProgressColumn(format="%.1f %%", min_value=0,
                                                      max_value=float(max(out["Anteil"].max(), 1)), color=MITTEL),
            "zu Vorjahr": st.column_config.NumberColumn(format="%+.1f %%"),
            **{c: f1 for c in out.columns if c not in ("Produkt", "Anteil", "zu Vorjahr")}})
        st.caption("Plan je Produkt und Monat = Plan des Marktbereichs x gewichteter Produktanteil der letzten 52 KW. "
                   "Werte in Tonnen.")
    c1, c2 = st.columns([2, 3], gap="medium")
    with c1:
        with st.container(border=True, height="stretch"):
            st.markdown("**Treiber der Veränderung**")
            st.plotly_chart(chart_treiber(mat, "Material", "Delta_kg"), config=PLOT_CFG, key=f"tr_mat_{mb}")
    with c2:
        zeitverlauf_block(D, mb, "Material", mat["Material"].tolist(), {}, key=f"zv_mat_{mb}")


def chart_guete(D, mb):
    g = pd.Series({m: genauigkeit_mb(D, m) * 100 for m in D["mbs"]}).iloc[::-1]
    fig = go.Figure(go.Bar(x=g.values, y=[f"MB {m}" for m in g.index], orientation="h",
                           marker_color=[NAVY if m == mb else HELL for m in g.index],
                           text=[f"{v:.0f} %" if np.isfinite(v) else "" for v in g.values], textposition="outside",
                           cliponaxis=False, hovertemplate="%{y}: %{x:.0f} %<extra></extra>"))
    fig = layout(fig, 60 + 30 * len(g), y_titel="", legende=False)
    for schwelle, text, anker in ((80, "hoch", "left"), (65, "mittel", "right")):
        fig.add_vline(x=schwelle, line=dict(color=GRAU, width=1, dash="dot"))
        fig.add_annotation(x=schwelle, y=1, yref="paper", text=f"{text} ab {schwelle} %", showarrow=False,
                           yanchor="bottom", xanchor=anker, xshift=4 if anker == "left" else -4,
                           font=dict(size=11, color=GRAU))
    fig.update_xaxes(range=[0, 108], ticksuffix=" %", gridcolor=ZUSATZ)
    fig.update_yaxes(type="category", gridcolor="rgba(0,0,0,0)")
    fig.update_layout(hovermode="closest", margin=dict(l=8, r=30, t=30, b=8))
    return fig


def tab_guete(D, mb):
    info = D["mbinfo"].loc[mb]
    f4, fv, bias = _zahl(info.get("Fehler_4W")), _zahl(info.get("Fehler_4W_Vorjahresmethode")), _zahl(info.get("Bias"))
    g, gv = (max(1 - f4, 0) if np.isfinite(f4) else np.nan), (max(1 - fv, 0) if np.isfinite(fv) else np.nan)
    drift = str(info.get("Drift", "")).lower() in ("true", "1", "wahr")
    label = guete_label(g)
    if str(info.get("Verlaesslichkeit")) == "nicht steuerungsrelevant":
        einordnung = "Kleinstmenge: für die Steuerung nicht relevant, Prognose nur als Orientierung."
    else:
        einordnung = {"hoch": "Die Zahlen sind belastbar.", "mittel": "Die Zahlen sind eingeschränkt belastbar; "
                      "größere Abweichungen mit dem Markt abgleichen.", "gering": "Die Zahlen nur mit Marktwissen "
                      "verwenden; Plan hier besonders begründen."}.get(label, "")
    vergleich = (f" Das Modell ist damit {pct(1 - f4 / fv, False, 0)} genauer als die Vorjahresmethode."
                 if np.isfinite(f4) and np.isfinite(fv) and fv > 0 and f4 < fv else
                 " Die Vorjahresmethode war hier nicht schlechter; Modell und Markt kritisch abwägen."
                 if np.isfinite(f4) and np.isfinite(fv) else "")
    st.markdown(f"<p class='wp-satz'>Im Backtest lag das Modell für MB {h(mb)} je 4 Wochen im Mittel "
                f"<b>{pct(f4, False)}</b> neben dem tatsächlichen Absatz, die Vorjahresmethode {pct(fv, False)}."
                f"{vergleich} <b>{einordnung}</b></p>", unsafe_allow_html=True)
    k = st.columns(4)
    k[0].metric("Prognosegenauigkeit", pct(g, False, 0) if np.isfinite(g) else "n/a", label.capitalize(),
                border=True, delta_color="off", delta_arrow="off", help=GUETE_HILFE)
    k[1].metric("Vorjahresmethode", pct(gv, False, 0) if np.isfinite(gv) else "n/a", "Vergleichsmaßstab",
                border=True, delta_color="off", delta_arrow="off")
    k[2].metric("Systematische Abweichung", pct(bias) if np.isfinite(bias) else "n/a",
                "Modell zu hoch" if bias > 0.02 else "Modell zu niedrig" if bias < -0.02 else "keine Tendenz",
                border=True, delta_color="off", delta_arrow="off",
                help="Mittlere Über- oder Unterschätzung im Backtest. Positiv: Modell lag zu hoch.")
    k[3].metric("Aktuelle Drift", "ja" if drift else "nein",
                "zuletzt daneben" if drift else "trifft das Niveau", border=True,
                delta_color="off", delta_arrow="off",
                help="Prüft, ob das Modell in den letzten 13 KW dauerhaft in eine Richtung abweicht.")
    c1, c2 = st.columns([3, 2], gap="medium")
    with c1:
        with st.container(border=True, height="stretch"):
            st.markdown("**Prognosegenauigkeit je Marktbereich**")
            st.plotly_chart(chart_guete(D, mb), config=PLOT_CFG, key=f"guete_{mb}")
    with c2:
        with st.container(border=True, height="stretch"):
            st.markdown("**So wird die Güte gemessen**")
            st.markdown(
                f"<p class='wp-hinweis'>{GUETE_HILFE}</p><p class='wp-hinweis'>Gewähltes Modell: "
                f"<b>{h(info.get('Modell', 'n/a'))}</b>, Verbrauchsklasse {h(info.get('XYZ', 'n/a'))}. Die Bandbreite "
                "im Verlauf zeigt den Bereich, in dem 80 % der Backtest-Abweichungen lagen.</p>",
                unsafe_allow_html=True)


# ============================================================================
# SEITE: EXTERNE FAKTOREN (folgt)
# ============================================================================
def seite_extern():
    D = ss.get("D")
    if D is None:
        return keine_daten()
    kopf("Externe Faktoren", "Korrelation des Absatzes mit externen Treibern je Marktbereich", D)
    with st.container(border=True):
        st.markdown(
            "<div class='wp-leer'><div class='wp-gross'>In Vorbereitung</div>"
            "<p class='wp-satz' style='max-width:46rem'>Hier folgt die Analyse externer Einflussgrößen je "
            "Marktbereich: Welche Indikatoren laufen dem Absatz voraus, wie stark ist der Zusammenhang und wie "
            "verändert er die Prognose.</p>"
            "<p class='wp-hinweis'>Geplant: Korrelation und Vorlauf je Indikator, Einfluss auf die Prognose, "
            "Szenarien für die nächsten sechs Monate.</p></div>", unsafe_allow_html=True)


# ============================================================================
# SEITE: FREIGABE UND BERICHT
# ============================================================================
def seite_freigabe():
    D = ss.get("D")
    if D is None:
        return keine_daten()
    st_ = ss["status"]
    frei, eingereicht = mbs_mit_status("Freigegeben"), mbs_mit_status("Eingereicht")
    kopf("Freigabe und Bericht", f"{len(frei)} von {len(st_)} Marktbereichen freigegeben, {len(eingereicht)} "
         f"{'Einreichung' if len(eingereicht) == 1 else 'Einreichungen'} offen", D)
    ist_fg = ansicht() == ANSICHT_FREIGABE
    if not ist_fg:
        c1, c2 = st.columns([6, 1.6], vertical_alignment="center")
        c1.markdown("<div class='wp-filter'>Sie sind in der Ansicht <b>Planung</b>. Genehmigen und Ablehnen ist in "
                    "der Ansicht <b>Freigabe</b> möglich.</div>", unsafe_allow_html=True)
        if c2.button("Zur Ansicht Freigabe", type="primary", width="stretch", key="zu_fg"):
            ss["ansicht"] = ANSICHT_FREIGABE
            ss["ansicht_v"] = ss.get("ansicht_v", 0) + 1
            st.rerun()
    freigabeband(D)
    st.space("small")

    # 1 Offene Einreichungen
    st.markdown("### Offene Einreichungen")
    if not eingereicht:
        st.caption("Keine offenen Einreichungen.")
    for mb in eingereicht:
        einreichung_karte(D, mb, ist_fg)

    # 2 Unveraenderte und noch offene Marktbereiche
    unveraendert = [mb for mb in mbs_mit_status("Offen") if not geaendert(D, mb)]
    in_arbeit = [mb for mb in mbs_mit_status("Offen", "Abgelehnt") if mb not in unveraendert]
    with st.container(border=True):
        c1, c2 = st.columns([5, 2], vertical_alignment="center")
        with c1:
            st.markdown("**Modellprognose unverändert**")
            st.markdown(f"<p class='wp-satz'>{mb_text(unveraendert) if unveraendert else 'Keine'}"
                        f"{' übernehmen die Modellprognose ohne Änderung.' if unveraendert else '.'}</p>"
                        + (f"<p class='wp-hinweis'>Noch in Bearbeitung, nicht eingereicht: {mb_text(in_arbeit)}.</p>"
                           if in_arbeit else ""), unsafe_allow_html=True)
        if c2.button("Modellprognose freigeben", type="primary", icon=":material/done_all:", width="stretch",
                     disabled=not (ist_fg and unveraendert), key="fg_unver"):
            for mb in unveraendert:
                genehmigen(D, mb, "Modellprognose übernommen")
            st.toast(f"{len(unveraendert)} Marktbereiche freigegeben.", icon=":material/check_circle:")
            st.rerun()

    # 3 Abweichungen und Bericht
    c1, c2 = st.columns([3, 2], gap="medium")
    with c1:
        with st.container(border=True, height="stretch"):
            st.markdown("### Abweichung Plan zu Modell")
            st.plotly_chart(chart_heatmap(D, plan_monat(D), wert_monat(D, D["prog"])), config=PLOT_CFG,
                            key="fg_heat")
            st.caption("Je Marktbereich und Monat. Blau: Plan über Modell, rot: Plan unter Modell.")
    with c2:
        with st.container(border=True, height="stretch"):
            final = len(frei) == len(st_)
            st.markdown("### Bericht")
            st.markdown("<p class='wp-satz'>" + (
                "Alle Marktbereiche sind freigegeben. Der Bericht ist die <b>finale Bedarfsplanung</b> der nächsten "
                "sechs Monate." if final else f"Noch nicht final: {len(st_) - len(frei)} Marktbereiche ohne Freigabe. "
                "Der Download ist ein Zwischenstand.") + "</p>", unsafe_allow_html=True)
            st.download_button("Finalen Bericht herunterladen" if final else "Zwischenstand herunterladen",
                               data=planstand_excel(D, final), type="primary", icon=":material/download:",
                               file_name=f"excel_w&p_{app_name().split(' ', 1)[-1].lower().replace(' ', '_')}"
                                         f"_{'final' if final else 'zwischenstand'}_{datetime.now():%y%m%d_%H%M}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               on_click="ignore", width="stretch")
            st.markdown("<p class='wp-hinweis'>Inhalt: Übersicht mit Kennzahlen und Prognosegüte, Plan je "
                        "Marktbereich und Monat sowie je KW, Änderungen mit Begründung und Freigabe, "
                        "Änderungsprotokoll. Die Datei lässt sich hier wieder hochladen, um weiterzuarbeiten.</p>",
                        unsafe_allow_html=True)

    # 4 Freigegebene Marktbereiche
    with st.expander(f"Freigegebene Marktbereiche ({len(frei)})"):
        if frei:
            t = pd.DataFrame({"MB": frei, "Plan 6 M (t)": [ss["plan"].loc[m].sum() / T for m in frei],
                              "zu Modell": [abweichung(D, m) * 100 for m in frei],
                              "Freigegeben von": st_.loc[frei, "Freigegeben_von"].values,
                              "am": st_.loc[frei, "Freigegeben_am"].values,
                              "Kommentar": st_.loc[frei, "Kommentar_Freigabe"].values})
            st.dataframe(t, hide_index=True, column_config={
                "Plan 6 M (t)": st.column_config.NumberColumn(format="localized"),
                "zu Modell": st.column_config.NumberColumn(format="%+.1f %%")})
            if ist_fg:
                c1, c2 = st.columns([2, 3], vertical_alignment="bottom")
                mb = c1.selectbox("Freigabe zurücknehmen für", frei, key="fg_zurueck")
                if c2.button("Zurücknehmen", icon=":material/lock_open:", key="fg_zurueck_go"):
                    st_.loc[mb, ["Status", "Freigegeben_von", "Freigegeben_am"]] = ["Offen", "", ""]
                    protokolliere("Freigabe zurückgenommen", mb, "", None, None, "")
                    st.rerun()
        else:
            st.caption("Noch keine Freigaben.")
    with st.expander("Änderungsprotokoll"):
        log = pd.DataFrame(ss["log"])
        if log.empty:
            st.caption("Noch keine Einträge.")
        else:
            st.dataframe(log.iloc[::-1], hide_index=True, height=420, column_config={
                "Vorher_t": st.column_config.NumberColumn("Vorher (t)", format="localized"),
                "Nachher_t": st.column_config.NumberColumn("Nachher (t)", format="localized"),
                "Kommentar": st.column_config.TextColumn(width="large")})


OHNE_GRUND = "Ohne Begründung (Abweichung unter der Schwelle)"


def einreichung_karte(D, mb, ist_fg):
    st_ = ss["status"]
    p6, f6 = ss["plan"].loc[mb].sum(), D["prog"].loc[mb].sum()
    with st.container(border=True):
        c1, c2, c3 = st.columns([2.2, 2.6, 1.8], gap="medium")
        with c1:
            g = genauigkeit_mb(D, mb)
            st.markdown(f"<div class='wp-gross'>MB {h(mb)}</div>"
                        f"<p class='wp-satz'>Modell {de(f6 / T)} t, Plan <b>{de(p6 / T)} t</b> "
                        f"({pct(p6 / f6 - 1) if f6 else 'n/a'})</p>"
                        f"<div class='wp-zitat'>{h(st_.loc[mb, 'Begruendung']) or OHNE_GRUND}</div>"
                        f"<p class='wp-hinweis'>Eingereicht {h(st_.loc[mb, 'Eingereicht_am'])}; "
                        f"Prognosegenauigkeit {pct(g, False, 0) if np.isfinite(g) else 'n/a'} ({guete_label(g)})</p>",
                        unsafe_allow_html=True)
        with c2:
            pm, fm = plan_monat(D).loc[mb], wert_monat(D, D["prog"]).loc[mb]
            t = pd.DataFrame({"Monat": [monat_label(m) for m in D["monate"]], "Modell (t)": fm.values / T,
                              "Plan (t)": pm.values / T,
                              "Abweichung": (pm.values / np.where(fm.values > 0, fm.values, np.nan) - 1) * 100})
            st.dataframe(t, hide_index=True, row_height=30, column_config={
                "Modell (t)": st.column_config.NumberColumn(format="%.1f"),
                "Plan (t)": st.column_config.NumberColumn(format="%.1f"),
                "Abweichung": st.column_config.NumberColumn(format="%+.1f %%")})
        with c3:
            kommentar = st.text_input("Kommentar", key=f"fg_kom_{mb}", placeholder="Pflicht bei Ablehnung",
                                      disabled=not ist_fg)
            if st.button("Genehmigen", type="primary", icon=":material/check:", width="stretch",
                         disabled=not ist_fg, key=f"fg_ok_{mb}"):
                genehmigen(D, mb, kommentar)
                st.toast(f"MB {mb} freigegeben.", icon=":material/check_circle:")
                st.rerun()
            if st.button("Ablehnen", icon=":material/close:", width="stretch",
                         disabled=not (ist_fg and kommentar.strip()), key=f"fg_nein_{mb}"):
                ablehnen(D, mb, kommentar.strip())
                st.rerun()
            if st.button("Im Detail ansehen", icon=":material/travel_explore:", width="stretch", key=f"fg_det_{mb}"):
                ss["mb_aktiv"] = mb
                st.switch_page(SEITEN["detail"])


# ============================================================================
# BERICHT (Excel)
# ============================================================================
def planstand_excel(D, final: bool) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter as L

    wb = Workbook()
    kopf_fill, band = PatternFill("solid", fgColor=NAVY[1:]), PatternFill("solid", fgColor=LILA[1:])
    kf, nf, bf = (Font(name="Calibri", bold=True, color="FFFFFF"), Font(name="Calibri", color=NAVY[1:]),
                  Font(name="Calibri", bold=True, color=NAVY[1:]))
    linie = Border(top=Side(style="medium", color=NAVY[1:]))
    TON, PROZ = '#,##0" t"', "+0.0%;-0.0%;0.0%"

    def kopfzeilen(ws, titel, kern):
        ws.sheet_view.showGridLines = False
        ws["A1"], ws["A2"] = titel, kern
        ws["A1"].font = Font(name="Calibri", size=16, bold=True, color=NAVY[1:])
        ws["A2"].font = Font(name="Calibri", size=11, color=MITTEL[1:])

    def tabelle(ws, df, start=4, fmt=None, summe=False, breite=None):
        for j, c in enumerate(df.columns, 1):
            z = ws.cell(row=start, column=j, value=str(c))
            z.fill, z.font = kopf_fill, kf
            z.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for i, row in enumerate(df.itertuples(index=False), start + 1):
            for j, v in enumerate(row, 1):
                if isinstance(v, float) and not np.isfinite(v):
                    v = None
                z = ws.cell(row=i, column=j, value=v.item() if isinstance(v, np.generic) else v)
                z.font = nf
                z.alignment = Alignment(vertical="top", wrap_text=isinstance(v, str) and len(v) > 40)
                if (i - start) % 2 == 0:
                    z.fill = band
                if fmt and df.columns[j - 1] in fmt:
                    z.number_format = fmt[df.columns[j - 1]]
                elif isinstance(v, (float, np.floating)):
                    z.number_format = "#,##0.0"
        ende = start + len(df)
        if summe:
            for j in range(1, len(df.columns) + 1):
                ws.cell(row=ende, column=j).font = bf
                ws.cell(row=ende, column=j).border = linie
        for j, c in enumerate(df.columns, 1):
            ws.column_dimensions[L(j)].width = (breite or {}).get(c, max(11, min(48, len(str(c)) + 4)))
        return ende

    st_ = ss["status"]
    mbs = D["mbs"]
    plan_m, prog_m, vj_m = plan_monat(D), wert_monat(D, D["prog"]), wert_monat(D, D["vj"])
    spalten = [monat_label(m, D["n_kw"][m]) for m in D["monate"]]
    stand = (f"Datenstand {D['meta'].get('Datenstand', '')}; Planhorizont {kw_label(D['W'][0])} bis "
             f"{kw_label(D['W'][-1])}; erstellt {jetzt()} von {nutzer_name()}")
    titel = f"{app_name()} | {'Finaler Plan' if final else 'Zwischenstand'}"

    # Blatt 1: Uebersicht
    ws = wb.active
    ws.title = "Übersicht"
    kopfzeilen(ws, titel, stand)
    p6, f6, vj6 = ss["plan"].values.sum(), D["prog"].values.sum(), D["vj"].values.sum()
    g, gv = genauigkeit(D, mbs)
    frei = mbs_mit_status("Freigegeben")
    k = pd.DataFrame([("Plan 6 Monate", p6 / T, ""), ("Modellprognose 6 Monate", f6 / T, ""),
                      ("Vorjahr gleiche KW", vj6 / T, ""), ("Plan zu Modell", p6 / f6 - 1 if f6 else None, ""),
                      ("Plan zu Vorjahr", p6 / vj6 - 1 if vj6 else None, ""),
                      ("Prognosegenauigkeit (Backtest)", g, guete_label(g)),
                      ("Vorjahresmethode (Backtest)", gv, "Vergleichsmaßstab"),
                      ("Freigegebene Marktbereiche", f"{len(frei)} von {len(mbs)}", "final" if final else "offen")],
                     columns=["Kennzahl", "Wert", "Einordnung"])
    ende = tabelle(ws, k, breite={"Kennzahl": 34, "Wert": 16, "Einordnung": 22})
    for r, f in zip(range(5, ende + 1), [TON, TON, TON, PROZ, PROZ, "0.0%", "0.0%", "@"]):
        ws.cell(row=r, column=2).number_format = f
        ws.cell(row=r, column=2).alignment = Alignment(horizontal="right")
        ws.cell(row=r, column=3).alignment = Alignment(horizontal="left", indent=1)
    ws.cell(row=ende + 2, column=1, value="Gesamt je Monat").font = bf
    vg = pd.DataFrame({"Monat": spalten, "Plan": plan_m.sum().values / T, "Modell": prog_m.sum().values / T,
                       "Vorjahr": vj_m.sum().values / T})
    vg["Plan zu Modell"] = vg["Plan"] / vg["Modell"].replace(0, np.nan) - 1
    vg["Plan zu Vorjahr"] = vg["Plan"] / vg["Vorjahr"].replace(0, np.nan) - 1
    tabelle(ws, vg, start=ende + 3, fmt={"Plan": TON, "Modell": TON, "Vorjahr": TON, "Plan zu Modell": PROZ,
                                         "Plan zu Vorjahr": PROZ}, breite={"Monat": 34, "Plan": 16, "Modell": 22})

    # Blatt 2: Plan MB x Monat
    df = pd.DataFrame({"Marktbereich": mbs, "Status": st_.loc[mbs, "Status"].values})
    for m, lbl in zip(D["monate"], spalten):
        df[lbl] = plan_m.loc[mbs, m].values / T
    df["Plan 6 M"] = ss["plan"].loc[mbs].sum(axis=1).values / T
    df["Modell 6 M"] = D["prog"].loc[mbs].sum(axis=1).values / T
    df["Vorjahr"] = D["vj"].loc[mbs].sum(axis=1).values / T
    df["Plan zu Modell"] = df["Plan 6 M"] / df["Modell 6 M"].replace(0, np.nan) - 1
    df["Plan zu Vorjahr"] = df["Plan 6 M"] / df["Vorjahr"].replace(0, np.nan) - 1
    summe = {c: df[c].sum() if c in spalten + ["Plan 6 M", "Modell 6 M", "Vorjahr"] else "" for c in df.columns}
    summe.update({"Marktbereich": "Gesamt", "Plan zu Modell": p6 / f6 - 1 if f6 else None,
                  "Plan zu Vorjahr": p6 / vj6 - 1 if vj6 else None})
    df = pd.concat([df, pd.DataFrame([summe])], ignore_index=True)
    ws = wb.create_sheet("Plan MB Monat")
    kopfzeilen(ws, "Bedarfsplan je Marktbereich und Monat", stand)
    tabelle(ws, df, fmt={**{c: TON for c in spalten + ["Plan 6 M", "Modell 6 M", "Vorjahr"]},
                         "Plan zu Modell": PROZ, "Plan zu Vorjahr": PROZ}, summe=True)
    ws.freeze_panes = "B5"

    # Blatt 3: Plan MB x KW
    kw = pd.DataFrame({"Marktbereich": mbs})
    for w in D["W"]:
        kw[kw_label(w)] = ss["plan"].loc[mbs, w].values / T
    kw = pd.concat([kw, pd.DataFrame([{"Marktbereich": "Gesamt", **{kw_label(w): ss["plan"][w].sum() / T
                                                                     for w in D["W"]}}])], ignore_index=True)
    ws = wb.create_sheet("Plan MB KW")
    kopfzeilen(ws, "Bedarfsplan je Marktbereich und KW (t)", "Grundlage für Kapazitäts- und Rohstoffplanung")
    tabelle(ws, kw, summe=True)
    ws.freeze_panes = "B5"

    # Blatt 4: Aenderungen und Freigabe
    ae = pd.DataFrame({"Marktbereich": mbs,
                       "Modell 6 M": D["prog"].loc[mbs].sum(axis=1).values / T,
                       "Plan 6 M": ss["plan"].loc[mbs].sum(axis=1).values / T,
                       "Abweichung": [abweichung(D, m) for m in mbs],
                       "Geändert": ["ja" if geaendert(D, m) else "nein" for m in mbs],
                       "Begründung": st_.loc[mbs, "Begruendung"].values,
                       "Status": st_.loc[mbs, "Status"].values,
                       "Freigabe von": st_.loc[mbs, "Freigegeben_von"].values,
                       "Freigabe am": st_.loc[mbs, "Freigegeben_am"].values,
                       "Kommentar Freigabe": st_.loc[mbs, "Kommentar_Freigabe"].values})
    ws = wb.create_sheet("Änderungen und Freigabe")
    kopfzeilen(ws, "Änderungen ggü. Modellprognose und Freigabe", stand)
    tabelle(ws, ae, fmt={"Modell 6 M": TON, "Plan 6 M": TON, "Abweichung": PROZ},
            breite={"Begründung": 60, "Kommentar Freigabe": 40, "Freigabe von": 18, "Freigabe am": 18})

    # Blatt 5: Prognoseguete
    info = D["mbinfo"]
    pg = pd.DataFrame({"Marktbereich": mbs, "Modell": info.loc[mbs, "Modell"].values,
                       "Genauigkeit": [genauigkeit_mb(D, m) for m in mbs],
                       "Vorjahresmethode": [max(1 - _zahl(info.loc[m, "Fehler_4W_Vorjahresmethode"]), 0) for m in mbs],
                       "Systematische Abweichung": [_zahl(info.loc[m, "Bias"]) for m in mbs],
                       "Drift": ["ja" if str(info.loc[m, "Drift"]).lower() in ("true", "1", "wahr") else "nein"
                                 for m in mbs],
                       "Einordnung": [guete_label(genauigkeit_mb(D, m)) for m in mbs]})
    ws = wb.create_sheet("Prognosegüte")
    kopfzeilen(ws, "Prognosegüte im Backtest", GUETE_HILFE[:220])
    tabelle(ws, pg, fmt={"Genauigkeit": "0.0%", "Vorjahresmethode": "0.0%", "Systematische Abweichung": PROZ},
            breite={"Modell": 34})

    # Blatt 6: Protokoll
    ws = wb.create_sheet("Änderungsprotokoll")
    kopfzeilen(ws, "Änderungsprotokoll", "Alle Änderungen und Entscheidungen dieser Runde")
    tabelle(ws, pd.DataFrame(ss["log"]).fillna(""), breite={"Kommentar": 60, "Zeit": 18})

    # Maschinenlesbar fuer das Wiederhochladen
    def roh(name, d):
        w_ = wb.create_sheet(name)
        w_.sheet_state = "hidden"
        w_.append(list(d.columns))
        for row in d.itertuples(index=False):
            w_.append([None if (isinstance(v, float) and not np.isfinite(v)) else
                       (v.to_pydatetime() if isinstance(v, pd.Timestamp) else
                        v.item() if isinstance(v, np.generic) else v) for v in row])
    lang = ss["plan"].stack().rename("Plan_kg").reset_index()
    lang.columns = ["MB", "KW_Start", "Plan_kg"]
    roh("APP_PLAN", lang)
    roh("APP_STATUS", st_.reset_index())
    roh("APP_LOG", pd.DataFrame(ss["log"]))
    roh("APP_PLANMETA", pd.DataFrame([("Lauf_ID", ss["lauf"]), ("Exportiert", jetzt()),
                                      ("Final", "ja" if final else "nein"), ("App_Version", APP_VERSION),
                                      ("Quelle", ss.get("quelle", ""))], columns=["Schluessel", "Wert"]))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ============================================================================
# SEITENLEISTE
# ============================================================================
def seitenleiste():
    with st.sidebar:
        D = ss.get("D")
        if D is None:
            st.markdown(f"<div class='wp-marke'><b>W&amp;P</b><span>{h(app_name())}</span></div>",
                        unsafe_allow_html=True)
        else:
            st.markdown("**Ansicht**")
            optionen = [ANSICHT_PLAN, ANSICHT_FREIGABE]
            wahl = st.radio("Ansicht", optionen, index=optionen.index(ansicht()), horizontal=True,
                            key=f"w_ansicht_{ss.get('ansicht_v', 0)}", label_visibility="collapsed")
            if wahl and wahl != ansicht():
                ss["ansicht"] = wahl
                offen = mbs_mit_status("Eingereicht")
                if wahl == ANSICHT_FREIGABE and offen:
                    st.toast(f"{len(offen)} {'Einreichung wartet' if len(offen) == 1 else 'Einreichungen warten'} "
                             "auf Freigabe.", icon=":material/notifications:")
                st.rerun()
            st.caption("Demo: Planung und Freigabe in einer Sitzung umschalten. Feste Rollen folgen mit dem "
                       "Mehrbenutzerbetrieb.")
            offen = mbs_mit_status("Eingereicht")
            if offen:
                st.markdown(f"<div class='wp-badge'>{len(offen)} "
                            f"{'Einreichung' if len(offen) == 1 else 'Einreichungen'} offen</div>",
                            unsafe_allow_html=True)
        if ZUGANG["aktiv"]:
            st.markdown(f"<div class='wp-schutz' style='margin-top:.6rem'>Angemeldet als <b>{h(ss['nutzer'])}</b></div>",
                        unsafe_allow_html=True)
            st.space("small")
            if st.button("Abmelden", icon=":material/logout:", width="stretch", key="abmelden"):
                abmelden("Sie wurden abgemeldet.")
        if D is None:
            if not ZUGANG["aktiv"]:
                st.markdown("<div class='wp-schutz'>Läuft nur auf diesem Rechner. Dateien bleiben im Arbeitsspeicher, "
                            "nichts wird gespeichert oder übertragen.</div>", unsafe_allow_html=True)
            return
        st.divider()
        st.caption(f"Prognose: {ss.get('quelle', '')}")
        with st.expander("Datei wechseln oder Planstand laden"):
            neu = st.file_uploader("Neue Prognose", type=["xlsx"], key="up_prog")
            if neu is not None:
                lade_prognose(neu.getvalue(), neu.name)
            ps = st.file_uploader("Gesicherter Planstand", type=["xlsx"], key="up_plan")
            if ps is not None:
                lade_planstand(ps.getvalue(), D)
        hinweis = ("Dateien bleiben nur für diese Sitzung im Arbeitsspeicher und sind nach dem Abmelden weg."
                   if ZUGANG["aktiv"] else "Nur lokal. Nichts wird gespeichert, außer Sie laden den Bericht herunter.")
        st.markdown(f"<div class='wp-schutz'>{hinweis}</div>", unsafe_allow_html=True)


# ----------------------------------------------------------------------------
# Uebernommene Bausteine
# ----------------------------------------------------------------------------
def _monat_text(key: str) -> str:
    return f"{MONATE[int(key[5:]) - 1]} {key[:4]}"


def zeitstrahl(D, key) -> tuple:
    """Zeitstrahl mit zwei Griffen: von Monat bis Monat, ueber die gesamte Ist-Historie plus Planhorizont.
    Standard: ab Januar des Vorjahres bis Ende Planhorizont. Auswahl gilt fuer Cockpit und Marktbereich."""
    erster = monat_key(D["ist"].columns[0])
    letzter = monat_key(D["W"][-1])
    alle = [str(p) for p in pd.period_range(erster, letzter, freq="M")]
    vorjahr = f"{(D['ist'].columns[-1] + pd.Timedelta(days=3)).year - 1}-01"
    standard = (vorjahr if vorjahr in alle else alle[0], alle[-1])
    if ss.get("zeitraum") is None or not set(ss["zeitraum"]) <= set(alle):
        ss["zeitraum"] = standard
    c1, c2 = st.columns([8, 1.4], vertical_alignment="center")
    with c1:
        wahl = st.select_slider("Zeitraum", options=alle, value=ss["zeitraum"], format_func=_monat_text,
                                key=key, label_visibility="collapsed")
        st.markdown(f"<div class='wp-hinweis' style='margin-top:-.4rem'>Datenbestand {_monat_text(alle[0])} bis "
                    f"{_monat_text(alle[-1])} (Ist bis {_monat_text(monat_key(D['ist'].columns[-1]))}, danach Plan)"
                    "</div>", unsafe_allow_html=True)
    with c2:
        if st.button("Gesamt", key=f"{key}_ges", width="stretch", help="Gesamten Datensatz anzeigen"):
            ss["zeitraum"] = (alle[0], alle[-1])
            ss.pop(key, None)
            st.rerun()
    ss["zeitraum"] = tuple(wahl)
    return ss["zeitraum"]


def kopf(titel: str, kernsatz: str, D: dict | None = None):
    meta = ""
    if D is not None:
        m = D["meta"]
        try:
            stand = pd.Timestamp(m.get("Datenstand")).strftime("%d.%m.%Y")
            ab = kw_label(D["W"][0])
            bis = kw_label(D["W"][-1])
            meta = f"Datenstand {stand}<br>Planhorizont {ab} bis {bis}"
        except Exception:
            meta = ""
    st.markdown(f"<div class='wp-kopf'><div><h1>{h(titel)}</h1><div class='kern'>{h(kernsatz)}</div></div>"
                f"<div class='meta'>{meta}</div></div>", unsafe_allow_html=True)


def status_pill(s: str) -> str:
    return (f"<span class='wp-pill' style='background:{STATUS_FARBE[s]};color:{STATUS_TEXT[s]}'>{h(s)}</span>")


def freigabeband(D):
    """Jeder Marktbereich als Segment, Breite = Mengenanteil, Farbe = Status."""
    st_ = ss["status"]
    anteil = D["prog"].sum(axis=1) / D["prog"].values.sum()
    seg = []
    for mb in D["mbs"]:
        s = st_.loc[mb, "Status"]
        w = max(anteil[mb] * 100, 1.2)
        txt = h(mb) if anteil[mb] > 0.035 else ""
        tip = f"MB {mb}: {s}, {pct(anteil[mb], False, 0)} der Menge"
        seg.append(f"<div title='{h(tip)}' style='flex:{w} 1 0;background:{STATUS_FARBE[s]};color:{STATUS_TEXT[s]}'>"
                   f"{txt}</div>")
    zaehl = st_["Status"].value_counts()
    leg = "".join(f"<span><i style='background:{STATUS_FARBE[s]}'></i>{s} {zaehl.get(s, 0)}</span>"
                  for s in STATUS if zaehl.get(s, 0) or s in ("Offen", "Freigegeben"))
    st.markdown(f"<div class='wp-band'>{''.join(seg)}</div><div class='wp-legende'>{leg}</div>",
                unsafe_allow_html=True)


def keine_daten():
    st.switch_page(SEITEN["start"])


def rueckblick(D):
    vp = ss.get("vorplan")
    if vp is None:
        return
    gemeinsam = [w for w in vp.columns if w in D["ist"].columns]
    if not gemeinsam:
        return
    with st.container(border=True):
        st.markdown("### Rückblick: Wie gut war der Vormonatsplan?")
        ist = D["ist"][gemeinsam].reindex(vp.index).fillna(0)
        pl = vp[gemeinsam]
        r = pd.DataFrame({"Marktbereich": vp.index, "Plan (t)": pl.sum(axis=1).values / T,
                          "Ist (t)": ist.sum(axis=1).values / T})
        r["Abweichung"] = (r["Ist (t)"] / r["Plan (t)"].replace(0, np.nan) - 1) * 100
        gp, gi = r["Plan (t)"].sum(), r["Ist (t)"].sum()
        st.markdown(f"<p class='wp-satz'>In {len(gemeinsam)} abgeschlossenen KW lag das Ist bei {de(gi)} t, "
                    f"der Plan bei {de(gp)} t ({pct(gi / gp - 1) if gp else 'n/a'}).</p>", unsafe_allow_html=True)
        st.dataframe(r, hide_index=True, column_config={
            "Plan (t)": st.column_config.NumberColumn(format="localized"),
            "Ist (t)": st.column_config.NumberColumn(format="localized"),
            "Abweichung": st.column_config.NumberColumn("Ist zu Plan", format="%+.1f %%")})


def zeitverlauf_block(D, mb, dim: str, kandidaten: list, labels: dict, key: str, hinweis: str = ""):
    """Liniendiagramm Monatsverlauf (24 Kalendermonate) fuer die groessten Positionen einer Dimension."""
    with st.container(border=True):
        c1, c2, c3 = st.columns([2, 3, 1.3], vertical_alignment="center")
        c1.markdown("**Monatsverlauf der letzten 24 Monate**")
        darst = c3.segmented_control("Darstellung", ["Tonnen", "Index"], default="Tonnen", key=f"{key}_d",
                                     label_visibility="collapsed") or "Tonnen"
        z = D.get("zeit")
        if z is None:
            st.caption("Monatsverlauf ist ab Prognoseskript 0.6.1 enthalten. Bitte eine aktuelle Ergebnis-Excel laden.")
            return
        zz = z[(z["MB"] == mb) & (z["Dimension"] == dim)]
        vorhanden = set(zz["Schluessel"])
        opts = [k for k in kandidaten if k in vorhanden]
        if not opts:
            st.caption("Keine Monatsdaten für diese Auswahl.")
            return
        wahl = c2.multiselect("Im Verlauf zeigen", opts, default=opts[:5], max_selections=6, key=key,
                              format_func=lambda k: labels.get(k, k), label_visibility="collapsed",
                              placeholder="Positionen wählen")
        if wahl:
            st.plotly_chart(chart_zeitverlauf(zz, sorted(z["Monat"].unique()), wahl, labels,
                                              index=darst == "Index"), config=PLOT_CFG, key=f"{key}_fig")
        st.caption("Fakturierte Menge je Kalendermonat. Vorauswahl: die fünf größten Positionen der letzten 52 KW. "
                   "Index: jede Linie relativ zu ihrem eigenen Durchschnitt, um Muster kleiner Positionen zu sehen."
                   + (" " + hinweis if hinweis else ""))


def tab_dimension(D, det, mat, spalte, name, mb):
    if det.empty:
        st.info(f"Keine {name}daten für diesen Marktbereich.")
        return
    c0, _ = st.columns([2, 3])
    mats = c0.multiselect("Auf Materialien eingrenzen", sorted(det["Material"].unique()), key=f"f_{spalte}_{mb}",
                          placeholder="Alle Materialien")
    if mats:
        det = det[det["Material"].isin(mats)]
    g = det.groupby(spalte).agg(Ist=("Ist_52W_kg", "sum"), VJ=("Ist_VJ_52W_kg", "sum"),
                                Materialien=("Material", "nunique"), Positionen=("Pos_52W", "sum"))
    if spalte == "Auftraggeber":
        g["Länder"] = det.groupby(spalte)["Landname"].agg(lambda s: ", ".join(sorted(set(s)))[:60])
    g["Delta_kg"] = g["Ist"] - g["VJ"]
    g = g.sort_values("Ist", ascending=False)
    gesamt = g["Ist"].sum()
    if gesamt > 0:
        top5 = g["Ist"].head(5).sum() / gesamt
        neu = g[(g["VJ"] <= 0) & (g["Ist"] > 0)]
        weg = g[(g["Ist"] <= 0) & (g["VJ"] > 0)]
        satz = (f"<b>Top 5 {'Kunden' if name == 'Kunde' else 'Länder'}</b> stehen für {pct(top5, False, 0)} der Menge "
                f"der letzten 52 KW ({len(g)} insgesamt).")
        if len(neu) or len(weg):
            satz += (f" Neu ggü. Vorjahr: <b>{len(neu)}</b> ({de(neu['Ist'].sum() / T)} t), "
                     f"weggefallen: <b>{len(weg)}</b> ({de(weg['VJ'].sum() / T)} t).")
        st.markdown(f"<p class='wp-satz'>{satz}</p>", unsafe_allow_html=True)
    if spalte == "Landname":
        k1, k2 = st.columns([3, 2], gap="medium")
        with k1:
            karte_block(laender_werte(det), key=f"karte_{mb}", hoehe=430)
        with k2:
            st.markdown("**Treiber der Veränderung**")
            st.plotly_chart(chart_treiber(g.reset_index(), spalte, "Delta_kg"), config=PLOT_CFG,
                            key=f"tr_{spalte}_{mb}")
        codes = det.groupby("Land")["Ist_52W_kg"].sum().sort_values(ascending=False)
        namen = det.drop_duplicates("Land").set_index("Land")["Landname"].to_dict()
        zeitverlauf_block(D, mb, "Land", codes.index.tolist(), namen, key=f"zv_land_{mb}",
                          hinweis="Gilt für den gesamten Marktbereich, unabhängig vom Materialfilter." if mats else "")
    c1, c2 = st.columns([3, 2], gap="medium") if spalte != "Landname" else (st.container(), None)
    with c1:
        out = pd.DataFrame({name: g.index, "Ist 52 KW (t)": g["Ist"].values / T,
                            "Vorjahr 52 KW (t)": g["VJ"].values / T,
                            "Veränderung (t)": g["Delta_kg"].values / T,
                            "Ist zu Vorjahr": (g["Ist"] / g["VJ"].replace(0, np.nan) - 1).values * 100,
                            "Anteil": (g["Ist"] / gesamt * 100).values if gesamt else 0})
        if "Länder" in g:
            out.insert(1, "Länder", g["Länder"].values)
        out["Materialien"] = g["Materialien"].values
        f1 = st.column_config.NumberColumn(format="%.1f")
        st.dataframe(out, hide_index=True, height=420, column_config={
            "Ist 52 KW (t)": f1, "Vorjahr 52 KW (t)": f1,
            "Veränderung (t)": st.column_config.NumberColumn(format="%+.1f"),
            "Ist zu Vorjahr": st.column_config.NumberColumn(format="%+.1f %%"),
            "Anteil": st.column_config.ProgressColumn(format="%.1f %%", min_value=0,
                                                      max_value=float(max(out["Anteil"].max(), 1)), color=MITTEL),
            "Materialien": st.column_config.NumberColumn(width=90)})
    if c2 is not None:
        with c2:
            st.markdown("**Treiber der Veränderung**")
            st.plotly_chart(chart_treiber(g.reset_index(), spalte, "Delta_kg"), config=PLOT_CFG,
                            key=f"tr_{spalte}_{mb}")
    st.caption("Ist-Werte der letzten 52 KW gegenüber den 52 KW davor; Basis Fakturadatum, Einheit KG.")
    if spalte == "Auftraggeber":
        zeitverlauf_block(D, mb, "Auftraggeber", g.index.tolist(), {}, key=f"zv_kunde_{mb}",
                          hinweis="Gilt für den gesamten Marktbereich, unabhängig vom Materialfilter." if mats else "")


def lade_planstand(inhalt: bytes, D):
    digest = hashlib.sha256(inhalt).hexdigest()
    if ss.get("planstand_geladen") == digest:
        return
    try:
        vs = lese_planstand(inhalt)
    except Exception as e:
        st.error(str(e))
        return
    if vs["meta"].get("Lauf_ID") == lauf_id(D):
        runde_starten(D, vs)
        ss["planstand_geladen"] = digest
        st.toast("Planstand übernommen, Sie arbeiten weiter.", icon=":material/check_circle:")
        st.rerun()
    else:
        startwert_dialog(vs, digest)


@st.dialog("Planstand aus einer früheren Runde")
def startwert_dialog(vs, digest):
    D = ss["D"]
    st.markdown(f"Der Planstand vom {vs['meta'].get('Exportiert', '')} gehört zu einer früheren Prognose. "
                "Zuständigkeiten werden übernommen, der Vormonatsplan erscheint im Cockpit als Rückblick gegen das Ist.")
    wahl = st.radio("Startwert für die neue Runde", ["Modellprognose", "Vormonatsplan"],
                    captions=["Neue Prognose als Basis (Standard)",
                              "Überlappende KW aus dem letzten Plan, übrige aus der neuen Prognose"])
    if st.button("Übernehmen", type="primary"):
        runde_starten(D, vs, startwert=wahl)
        ss["planstand_geladen"] = digest
        st.rerun()


# ============================================================================
# KARTENDATEN: Laendergrenzen Natural Earth 1:50 Mio (gemeinfrei), vereinfacht, ISO-3166-Alpha-2,
# gzip + base64. Wird nur fuer die Laenderkarte gelesen; keine Kundendaten.
# ============================================================================
WELT_B64 = (
    "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"
    "e+6L88X++fc/T3bCZJLNBwA="
)


SEITEN = {
    "start": st.Page(seite_start, title="Start", url_path="start"),
    "uebersicht": st.Page(seite_uebersicht, title="Übersicht", url_path="uebersicht"),
    "detail": st.Page(seite_detail, title="Marktbereich im Detail", url_path="detail"),
    "extern": st.Page(seite_extern, title="Externe Faktoren", url_path="externe-faktoren"),
    "freigabe": st.Page(seite_freigabe, url_path="freigabe",
                        title=(f"Freigabe und Bericht ({len(mbs_mit_status('Eingereicht'))})"
                               if ss.get("status") is not None and len(mbs_mit_status("Eingereicht"))
                               else "Freigabe und Bericht")),
}

# ============================================================================
# ZUGANGSSCHUTZ
# ============================================================================
# Betrieb mit Anmeldung, sobald in den Secrets Nutzer hinterlegt sind (Pflicht in der Cloud):
#
#   kunde = "Kundenname"                  # optional; erscheint erst nach der Anmeldung
#   sitzung_minuten = 60                  # optional; Abmeldung nach Inaktivitaet
#   [nutzer."D. Karges"]
#   passwort_hash = "pbkdf2_sha256$..."   # erzeugen mit passwort_hash_erzeugen.py
#   rolle = "koordination"                # oder "mb"
#
# Ohne Nutzer startet die App nur auf localhost (Einzelplatz) und verweigert sonst den Start.
# Passwoerter liegen nie im Klartext vor, nur als PBKDF2-Hash mit Salt.
import hmac
import time

MAX_FEHL_SITZUNG = 5          # danach Wartezeit je Sitzung
MAX_FEHL_GLOBAL = 15          # Fehlversuche aller Sitzungen in 15 Minuten, danach Sperre
SPERRE_SEKUNDEN = 900


def _secrets() -> dict:
    try:
        return {k: st.secrets[k] for k in st.secrets}
    except Exception:
        return {}


def _nutzer() -> dict:
    roh = _secrets().get("nutzer", {}) or {}
    out = {}
    for name, eintrag in dict(roh).items():
        e = dict(eintrag)
        if e.get("passwort_hash"):
            out[str(name).strip()] = {"hash": str(e["passwort_hash"]),
                                      "rolle": ROLLE_KOORD if str(e.get("rolle", "mb")).lower().startswith("koord")
                                      else ROLLE_MB}
    return out


def ist_lokal() -> bool:
    try:
        host = (st.context.headers.get("Host") or "").rsplit(":", 1)[0].strip("[]").lower()
    except Exception:
        return False
    return host in ("localhost", "127.0.0.1", "::1")


def app_name() -> str:
    k = str(_secrets().get("kunde", "")).strip()
    if ZUGANG["aktiv"] and not ss.get("angemeldet"):
        k = ""                       # vor der Anmeldung kein Kundenname
    return f"Bedarfsplanung {k}".strip()


def pruefe_hash(passwort: str, gespeichert: str) -> bool:
    try:
        verfahren, runden, salt, soll = gespeichert.split("$")
        if verfahren != "pbkdf2_sha256":
            return False
        ist = hashlib.pbkdf2_hmac("sha256", passwort.encode("utf-8"), bytes.fromhex(salt), int(runden)).hex()
        return hmac.compare_digest(ist, soll)
    except Exception:
        return False


@st.cache_resource
def _fehlversuche() -> dict:
    """Prozessweiter Zaehler gegen Durchprobieren von Passwoertern ueber viele Sitzungen."""
    return {"zeiten": [], "gesperrt_bis": 0.0}


def abmelden(meldung: str = ""):
    ss.clear()
    if meldung:
        ss["abmelde_meldung"] = meldung
    st.rerun()


def zugang_pruefen():
    nutzer = _nutzer()
    ZUGANG["aktiv"] = bool(nutzer)
    if not nutzer:
        if ist_lokal():
            return                                   # Einzelplatz auf diesem Rechner
        st.error("Zugang nicht eingerichtet. Diese Instanz startet ohne hinterlegte Nutzer nicht. "
                 "Nutzer mit Passwort-Hash in den Secrets eintragen.", icon=":material/lock:")
        st.stop()

    jetzt_s = time.time()
    if ss.get("angemeldet"):
        grenze = float(_secrets().get("sitzung_minuten", 60)) * 60
        if jetzt_s - ss.get("aktiv_seit", jetzt_s) > grenze:
            abmelden("Aus Sicherheitsgründen nach Inaktivität abgemeldet. Bitte erneut anmelden.")
        ss["aktiv_seit"] = jetzt_s
        return

    g = _fehlversuche()
    g["zeiten"] = [t for t in g["zeiten"] if jetzt_s - t < SPERRE_SEKUNDEN]
    _, mitte, _ = st.columns([1, 1.15, 1])
    with mitte:
        st.space("large")
        st.markdown("<div class='wp-start'><h1 style='font-size:2.1rem'>Bedarfsplanung</h1>"
                    "<p class='lead' style='font-size:1.02rem'>Zugang nur für das Projektteam.</p></div>",
                    unsafe_allow_html=True)
        if ss.get("abmelde_meldung"):
            st.info(ss.pop("abmelde_meldung"), icon=":material/info:")
        with st.container(border=True):
            gesperrt = g["gesperrt_bis"] > jetzt_s
            warten = ss.get("warten_bis", 0) > jetzt_s
            with st.form("login", border=False):
                name = st.text_input("Name", placeholder="Vorname Nachname wie vereinbart", autocomplete="username")
                pw = st.text_input("Passwort", type="password", autocomplete="current-password")
                los = st.form_submit_button("Anmelden", type="primary", width="stretch",
                                            disabled=gesperrt)
            if gesperrt:
                st.error("Zu viele Fehlversuche. Anmeldung vorübergehend gesperrt, bitte später erneut versuchen.")
            elif los:
                if warten:
                    st.warning("Bitte einen Moment warten und erneut versuchen.")
                else:
                    treffer = next((n for n in nutzer if n.lower() == name.strip().lower()), None)
                    ok = pruefe_hash(pw, nutzer[treffer]["hash"]) if treffer else \
                        pruefe_hash(pw, "pbkdf2_sha256$600000$00$00")   # gleiche Laufzeit ohne Treffer
                    if ok:
                        ss.clear()
                        ss.update({"angemeldet": True, "nutzer": treffer, "rolle": nutzer[treffer]["rolle"],
                                   "aktiv_seit": time.time()})
                        st.rerun()
                    g["zeiten"].append(jetzt_s)
                    if len(g["zeiten"]) >= MAX_FEHL_GLOBAL:
                        g["gesperrt_bis"] = jetzt_s + SPERRE_SEKUNDEN
                    n = ss.get("fehl", 0) + 1
                    ss["fehl"] = n
                    if n >= MAX_FEHL_SITZUNG:
                        ss["warten_bis"] = jetzt_s + min(30 * 2 ** (n - MAX_FEHL_SITZUNG), 900)
                    st.error("Name oder Passwort falsch.")
        st.caption("Verbindung verschlüsselt. Abmeldung automatisch nach Inaktivität.")
    st.stop()


ZUGANG = {"aktiv": False}
zugang_pruefen()
seitenleiste()
if ss.get("D") is None:
    nav = st.navigation([SEITEN["start"]], position="hidden")
else:
    nav = st.navigation([SEITEN["uebersicht"], SEITEN["detail"], SEITEN["extern"], SEITEN["freigabe"]],
                        position="sidebar")
nav.run()
