"""
ABSATZPLANUNG  |  W&P  |  S&OP Planungs-App auf Marktbereichsebene
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

import hashlib
import io
import warnings
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.graph_objects as go
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

STATUS = ["Offen", "In Bearbeitung", "Eingereicht", "Freigegeben", "Zurückgewiesen"]
STATUS_FARBE = {"Offen": HELLGRAU, "In Bearbeitung": HELL, "Eingereicht": MITTEL,
                "Freigegeben": NAVY, "Zurückgewiesen": ROT}
STATUS_TEXT = {"Offen": NAVY, "In Bearbeitung": NAVY, "Eingereicht": "#FFFFFF",
               "Freigegeben": "#FFFFFF", "Zurückgewiesen": "#FFFFFF"}
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
[data-testid="stSidebarNav"]::before {{ content:"W&P   Absatzplanung"; display:block;
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
    return dict(meta=meta, mbinfo=mbinfo.set_index("MB").reindex(reihenfolge), prog=prog.loc[reihenfolge],
                bu=bu.loc[reihenfolge], bo=bo.loc[reihenfolge], ist=ist.loc[reihenfolge],
                vj=vj.loc[reihenfolge].fillna(0.0), ges=g, W=W, monat=monat, monate=monate, n_kw=n_kw,
                mat=mat, det=det, mbs=list(reihenfolge))


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
               ("ES", "Spanien"), ("NL", "Niederlande"), ("TR", "Türkei"), ("US", "USA")]

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
                code, name = laender[int(rng.choice(8, p=[.36, .14, .12, .1, .09, .08, .06, .05]))]
                kunde = f"{100000 + int(rng.integers(1, 60))}"
                a = rng.uniform(.05, .5)
                det.append(dict(MB=mb, Material=m, Land=code, Landname=name, Auftraggeber=kunde,
                                Ist_52W_kg=ist52 * a, Ist_VJ_52W_kg=vj52 * a * rng.lognormal(0, .25),
                                Pos_52W=int(rng.integers(3, 60))))
    for j, w in enumerate(prog_w):
        kw.append(("GESAMT", w, "Prognose", ges[j], ges[j] * (0.93 - j * 0.002), ges[j] * (1.07 + j * 0.003)))
    meta = pd.DataFrame([("App_Version", "1"), ("System_Version", "Demo"), ("Datenstand", "2026-10-29"),
                         ("Letzte_Ist_KW", f"{stand:%Y-%m-%d}"), ("Erste_Prognose_KW", f"{prog_w[0]:%Y-%m-%d}"),
                         ("Horizont_Wochen", "26"), ("Quelldatei", "Demodaten (synthetisch)"),
                         ("Daten_SHA256", "demo"), ("Erstellt", "2026-10-30 08:00"), ("Zeitachse", "Fakturadatum"),
                         ("Einheit", "KG"), ("Fehler_4W_Gesamt", "0.084"), ("Fehler_4W_Vorjahresmethode", "0.142"),
                         ("Bias_Gesamt", "0.011")], columns=["Schluessel", "Wert"])
    return aufbereiten({"APP_META": meta, "APP_MB": pd.DataFrame(mbz),
                        "APP_KW": pd.DataFrame(kw, columns=["MB", "KW_Start", "Typ", "Menge_kg", "Band_unten_kg",
                                                           "Band_oben_kg"]),
                        "APP_MATERIAL": pd.DataFrame(mat), "APP_DETAIL": pd.DataFrame(det)})


# ============================================================================
# SITZUNGSZUSTAND: Plan, Status, Zustaendigkeiten, Protokoll
# ============================================================================
ss = st.session_state


def jetzt() -> str:
    return datetime.now().strftime("%d.%m.%Y %H:%M")


def lauf_id(D) -> str:
    return f"{D['meta'].get('Daten_SHA256', '')[:16]}|{D['meta'].get('Erste_Prognose_KW', '')}"


def runde_starten(D: dict, vorstand: dict | None = None, startwert: str = "Modellprognose"):
    """Neue Planungsrunde oder Fortsetzung eines gesicherten Planstands."""
    plan = D["prog"].copy()
    status = pd.DataFrame({"MB": D["mbs"], "Verantwortlich": OHNE, "Status": "Offen", "Begruendung": "",
                           "Kommentar_Koordination": "", "Geaendert_von": "", "Geaendert_am": ""}).set_index("MB")
    if "team" in ss:                       # Zustaendigkeiten bleiben ueber Laeufe erhalten
        for mb, v in ss["team"].items():
            if mb in status.index:
                status.loc[mb, "Verantwortlich"] = v
    log = []
    ss["vorplan"] = None
    if vorstand is not None:
        vs = vorstand["status"].set_index("MB")
        for mb in status.index.intersection(vs.index):
            status.loc[mb, "Verantwortlich"] = str(vs.loc[mb, "Verantwortlich"] or OHNE)
        gleich = vorstand["meta"].get("Lauf_ID") == lauf_id(D)
        vp = vorstand["plan"].pivot(index="MB", columns="KW_Start", values="Plan_kg")
        if gleich:
            plan.update(vp.reindex(index=plan.index, columns=plan.columns))
            for c in ("Status", "Begruendung", "Kommentar_Koordination", "Geaendert_von", "Geaendert_am"):
                if c in vs:
                    status.loc[status.index.intersection(vs.index), c] = vs[c].fillna("").astype(str)
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
    ss["plan"], ss["status"], ss["log"] = plan, status, log
    ss["team"] = status["Verantwortlich"].to_dict()
    ss["lauf"] = lauf_id(D)
    ss["editor_v"] = ss.get("editor_v", 0) + 1


def eintrag(aktion, mb, objekt, vorher, nachher, kommentar="") -> dict:
    return {"Zeit": jetzt(), "Nutzer": ss.get("nutzer") or "unbekannt", "Rolle": ss.get("rolle", ""),
            "Marktbereich": mb, "Aktion": aktion, "Objekt": objekt,
            "Vorher_t": None if vorher is None else round(vorher / T, 1),
            "Nachher_t": None if nachher is None else round(nachher / T, 1), "Kommentar": kommentar}


def protokolliere(*a, **k):
    ss["log"].append(eintrag(*a, **k))


def plan_monat(D) -> pd.DataFrame:
    p = ss["plan"]
    return p.T.groupby(pd.Series(D["monat"])).sum().T.reindex(columns=D["monate"])


def wert_monat(D, df) -> pd.DataFrame:
    return df.T.groupby(pd.Series(D["monat"])).sum().T.reindex(columns=D["monate"])


def bearbeitet(mb):
    if ss["status"].loc[mb, "Status"] == "Offen":
        ss["status"].loc[mb, "Status"] = "In Bearbeitung"
    ss["status"].loc[mb, ["Geaendert_von", "Geaendert_am"]] = [ss.get("nutzer") or "unbekannt", jetzt()]


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
    s = ss["status"].loc[mb, "Status"]
    if ss.get("rolle") == ROLLE_KOORD:
        return s != "Freigegeben"
    return (bool(ss.get("nutzer")) and ss["status"].loc[mb, "Verantwortlich"] == ss.get("nutzer")
            and s in ("Offen", "In Bearbeitung", "Zurückgewiesen"))


def abweichung(D, mb) -> float:
    p = D["prog"].loc[mb].sum()
    return ss["plan"].loc[mb].sum() / p - 1 if p > 0 else np.nan


def braucht_begruendung(D, mb) -> bool:
    a = abweichung(D, mb)
    return np.isfinite(a) and abs(a) > SCHWELLE and not str(ss["status"].loc[mb, "Begruendung"]).strip()


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


def chart_wochen(ist, prog, bu, bo, plan, vj=None, wochen_ist=104, hoehe=380):
    ist = ist.iloc[-wochen_ist:] / T
    prog, bu, bo, plan = prog / T, bu / T, bo / T, plan / T
    x0 = ist.index[-1]
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=list(bu.index) + list(bo.index[::-1]), y=list(bu) + list(bo[::-1]), fill="toself",
                             fillcolor="rgba(187,204,238,.45)", line=dict(width=0), name="Bandbreite 80 %",
                             hoverinfo="skip"))
    if vj is not None:
        fig.add_trace(go.Scatter(x=vj.index, y=vj / T, name="Vorjahr", line=dict(color=GRAU, width=1.4, dash="dot"),
                                 hovertemplate="%{y:,.0f} t"))
    fig.add_trace(go.Scatter(x=ist.index, y=ist, name="Ist", line=dict(color=NAVY, width=1.8),
                             hovertemplate="%{y:,.0f} t"))
    fig.add_trace(go.Scatter(x=[x0] + list(prog.index), y=[ist.iloc[-1]] + list(prog), name="Modellprognose",
                             line=dict(color=MITTEL, width=2, dash="dash"), hovertemplate="%{y:,.0f} t"))
    fig.add_trace(go.Scatter(x=[x0] + list(plan.index), y=[ist.iloc[-1]] + list(plan), name="Plan",
                             line=dict(color=NAVY, width=3.2), hovertemplate="%{y:,.0f} t"))
    fig.add_vline(x=x0 + pd.Timedelta(days=3.5), line=dict(color=HELLGRAU, width=1))
    return layout(fig, hoehe)


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


# ============================================================================
# BAUSTEINE
# ============================================================================
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
        tip = f"{mb} | {s} | {pct(anteil[mb], False, 0)} der Menge | {st_.loc[mb, 'Verantwortlich']}"
        seg.append(f"<div title='{h(tip)}' style='flex:{w} 1 0;background:{STATUS_FARBE[s]};color:{STATUS_TEXT[s]}'>"
                   f"{txt}</div>")
    zaehl = st_["Status"].value_counts()
    leg = "".join(f"<span><i style='background:{STATUS_FARBE[s]}'></i>{s} {zaehl.get(s, 0)}</span>"
                  for s in STATUS if zaehl.get(s, 0) or s in ("Offen", "Freigegeben"))
    st.markdown(f"<div class='wp-band'>{''.join(seg)}</div><div class='wp-legende'>{leg}</div>",
                unsafe_allow_html=True)


def keine_daten():
    st.switch_page(SEITEN["start"])


# ============================================================================
# SEITE: START
# ============================================================================
def seite_start():
    st.markdown(
        "<div class='wp-start'><h1>Absatzplanung Marktbereiche</h1>"
        "<p class='lead'>Ein abgestimmter Absatzplan für die nächsten sechs Monate, einmal im Monat. "
        "Die Modellprognose ist der Startpunkt, die Marktbereiche bringen ihr Wissen ein, "
        "die S&amp;OP-Koordination gibt frei.</p>"
        "<div class='wp-schritte'>"
        "<div><b>1  Prognose laden</b><span>Ergebnis-Excel des Prognoseskripts hochladen, optional "
        "den letzten Planstand dazu.</span></div>"
        "<div><b>2  Plan abstimmen</b><span>Je Marktbereich Monatswerte anpassen, Abweichungen begründen, "
        "bei Fragen bis auf Material, Land und Kunde eintauchen.</span></div>"
        "<div><b>3  Freigeben und sichern</b><span>Koordination gibt frei, der Planstand wird als Excel "
        "gesichert und ist Startpunkt der nächsten Runde.</span></div>"
        "</div></div>", unsafe_allow_html=True)
    c1, c2 = st.columns([3, 2], gap="large")
    with c1:
        with st.container(border=True):
            st.markdown("### Prognose hochladen")
            datei = st.file_uploader("Ergebnis-Excel des Prognoseskripts",
                                     type=["xlsx"], key="up_start")
            if datei is not None:
                lade_prognose(datei.getvalue(), datei.name)
            st.caption("Die Datei bleibt im Arbeitsspeicher dieses Rechners. Nichts wird gespeichert oder übertragen.")
    with c2:
        with st.container(border=True):
            st.markdown("### Ohne Kundendaten testen")
            st.markdown("<div class='wp-hinweis'>Synthetische Daten mit acht Marktbereichen, um die Bedienung "
                        "zu zeigen oder Planer einzuweisen.</div>", unsafe_allow_html=True)
            st.space("small")
            if st.button("Demodaten laden", icon=":material/science:", width="stretch"):
                ss["D"] = demo_daten()
                ss["quelle"] = "Demodaten"
                ss["team"] = {"110": "A. Becker", "120": "A. Becker", "130": "M. Wagner", "210": "M. Wagner",
                              "220": "S. Hoffmann", "310": "S. Hoffmann", "320": "S. Hoffmann", "410": "S. Hoffmann"}
                runde_starten(ss["D"])
                st.switch_page(SEITEN["cockpit"])


def lade_prognose(inhalt: bytes, name: str):
    try:
        D = lese_prognose(inhalt)
    except Exception as e:
        st.error(str(e) if isinstance(e, ValueError) else f"Datei konnte nicht gelesen werden: {e}")
        return
    if ss.get("D") is None or lauf_id(D) != ss.get("lauf"):
        ss["D"], ss["quelle"] = D, name
        runde_starten(D)
        st.switch_page(SEITEN["cockpit"])


# ============================================================================
# SEITE: COCKPIT
# ============================================================================
def seite_cockpit():
    D = ss.get("D")
    if D is None:
        return keine_daten()
    plan, prog = ss["plan"], D["prog"]
    p6, f6, vj6 = plan.values.sum(), prog.values.sum(), D["vj"].values.sum()
    l26 = D["ist"].iloc[:, -26:].values.sum()
    st_ = ss["status"]
    n_frei = int((st_["Status"] == "Freigegeben").sum())
    anteil_frei = prog.loc[st_.index[st_["Status"] == "Freigegeben"]].values.sum() / f6 if f6 else 0
    rm = next((m for m in D["monate"] if D["n_kw"][m] >= 4), D["monate"][0])
    kopf(f"Planungsrunde {MONATE_LANG[int(rm[5:]) - 1]} {rm[:4]}",
         f"Plan {de(p6 / T)} t für 6 Monate, {pct(p6 / vj6 - 1) if vj6 else 'n/a'} ggü. Vorjahr", D)

    freigabeband(D)
    st.space("small")

    k = st.columns(4)
    k[0].metric("Plan 6 Monate", f"{de(p6 / T)} t", f"{pct(p6 / f6 - 1)} zur Modellprognose", border=True,
                delta_color="off", delta_arrow="off")
    k[1].metric("Ggü. Vorjahr (gleiche KW)", pct(p6 / vj6 - 1) if vj6 else "n/a", f"Vorjahr {de(vj6 / T)} t",
                border=True, delta_color="off", delta_arrow="off")
    k[2].metric("Ggü. letzte 26 KW", pct(p6 / l26 - 1) if l26 else "n/a", f"Ist {de(l26 / T)} t", border=True,
                delta_color="off", delta_arrow="off")
    k[3].metric("Freigegeben", f"{n_frei} von {len(st_)}", f"{pct(anteil_frei, False, 0)} der Menge", border=True,
                delta_color="off", delta_arrow="off")

    # Kernaussagen
    abw = pd.Series({mb: abweichung(D, mb) for mb in D["mbs"]})
    groesste = (plan.sum(axis=1) - prog.sum(axis=1)).abs().sort_values(ascending=False)
    saetze = []
    if groesste.iloc[0] > 0.005 * f6:
        mb = groesste.index[0]
        saetze.append(f"Größter Eingriff in <b>MB {h(mb)}</b>: {pct(abw[mb])} ggü. Modell "
                      f"({de((plan.loc[mb].sum() - prog.loc[mb].sum()) / T)} t).")
    offen = st_.index[st_["Status"].isin(["Offen", "In Bearbeitung", "Zurückgewiesen"])]
    if len(offen):
        saetze.append(f"Noch in Arbeit: <b>{len(offen)} Marktbereiche</b> mit "
                      f"{pct(prog.loc[offen].values.sum() / f6, False, 0)} der Menge.")
    ohne_grund = [mb for mb in D["mbs"] if braucht_begruendung(D, mb)]
    if ohne_grund:
        saetze.append(f"Begründung fehlt bei Abweichung über {pct(SCHWELLE, False, 0)}: <b>MB "
                      f"{', '.join(map(h, ohne_grund))}</b>.")
    drift = D["mbinfo"].index[D["mbinfo"]["Drift"].astype(str).str.lower().isin(["true", "1", "wahr"])]
    if len(drift):
        saetze.append(f"Modell meldet systematische Abweichung (Drift) in <b>MB {', '.join(map(h, drift))}</b>; "
                      "Plan dort besonders kritisch prüfen.")

    links, rechts = st.columns([5, 2], gap="medium")
    with links:
        with st.container(border=True):
            c1, c2 = st.columns([3, 2], vertical_alignment="center")
            c1.markdown("### Gesamtabsatz: Ist, Modellprognose und Plan")
            sicht = c2.segmented_control("Sicht", ["Wochen", "Monate"], default="Wochen", key="ck_sicht",
                                         label_visibility="collapsed")
            if sicht == "Monate":
                fig = chart_monate(D, plan_monat(D).sum().values, wert_monat(D, prog).sum().values,
                                   wert_monat(D, D["vj"]).sum().values)
            else:
                g = D["ges"]
                fig = chart_wochen(D["ist"].sum(), g["Menge_kg"], g["Band_unten_kg"], g["Band_oben_kg"], plan.sum(),
                                   vj=D["vj"].sum())
            st.plotly_chart(fig, config=PLOT_CFG, key="ck_chart")
    with rechts:
        with st.container(border=True, height="stretch"):
            st.markdown("### Kernaussagen")
            if not saetze:
                saetze = ["Plan entspricht der Modellprognose; noch keine Anpassungen."]
            for s in saetze:
                st.markdown(f"<p class='wp-satz'>{s}</p>", unsafe_allow_html=True)
            fm = D["meta"]
            try:
                st.markdown(f"<p class='wp-hinweis'>Modellgüte gesamt: Fehler je 4 Wochen "
                            f"{pct(float(fm['Fehler_4W_Gesamt']), False)}, Vorjahresmethode "
                            f"{pct(float(fm['Fehler_4W_Vorjahresmethode']), False)}.</p>", unsafe_allow_html=True)
            except (KeyError, ValueError):
                pass

    # Tabelle Marktbereiche
    with st.container(border=True):
        st.markdown("### Marktbereiche")
        st.caption("Zeile anklicken, um den Marktbereich im Detail zu öffnen.")
        info = D["mbinfo"]
        tab = pd.DataFrame({
            "Marktbereich": D["mbs"],
            "Verantwortlich": st_.loc[D["mbs"], "Verantwortlich"].values,
            "Status": st_.loc[D["mbs"], "Status"].values,
            "Verlauf Ist 52 KW": [(D["ist"].loc[mb].iloc[-52:] / T).round(1).tolist() for mb in D["mbs"]],
            "Plan 6 M (t)": (plan.loc[D["mbs"]].sum(axis=1).values / T).round(0),
            "Modell 6 M (t)": (prog.loc[D["mbs"]].sum(axis=1).values / T).round(0),
            "Plan zu Modell": [abweichung(D, mb) * 100 for mb in D["mbs"]],
            "Plan zu Vorjahr": [(plan.loc[mb].sum() / D["vj"].loc[mb].sum() - 1) * 100
                                if D["vj"].loc[mb].sum() > 0 else np.nan for mb in D["mbs"]],
            "Anteil": (prog.loc[D["mbs"]].sum(axis=1) / f6 * 100).values,
            "Prognosegüte": info["Verlaesslichkeit"].fillna("").values,
            "Begründung": st_.loc[D["mbs"], "Begruendung"].values,
        })
        auswahl = st.dataframe(
            tab, hide_index=True, on_select="rerun", selection_mode="single-row", key="ck_tab",
            row_height=38,
            column_config={
                "Marktbereich": st.column_config.TextColumn("MB", width=60),
                "Verantwortlich": st.column_config.TextColumn(width=130),
                "Verlauf Ist 52 KW": st.column_config.AreaChartColumn(width=150, color=MITTEL),
                "Plan 6 M (t)": st.column_config.NumberColumn(format="localized", width=95),
                "Modell 6 M (t)": st.column_config.NumberColumn(format="localized", width=95),
                "Plan zu Modell": st.column_config.NumberColumn(format="%+.1f %%", width=95),
                "Plan zu Vorjahr": st.column_config.NumberColumn(format="%+.1f %%", width=95),
                "Anteil": st.column_config.ProgressColumn(format="%.0f %%", min_value=0, max_value=100,
                                                          color=NAVY, width=90),
                "Prognosegüte": st.column_config.TextColumn(width=100),
                "Status": st.column_config.TextColumn(width=110),
                "Begründung": st.column_config.TextColumn(width="large"),
            })
        if auswahl and auswahl.selection.rows:
            ss["mb_aktiv"] = D["mbs"][auswahl.selection.rows[0]]
            st.switch_page(SEITEN["mb"])

    rueckblick(D)


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


# ============================================================================
# SEITE: PLANUNG (MB x Monat)
# ============================================================================
def seite_planung():
    D = ss.get("D")
    if D is None:
        return keine_daten()
    st_ = ss["status"]
    plan_m, prog_m = plan_monat(D), wert_monat(D, D["prog"])
    p6, f6 = ss["plan"].values.sum(), D["prog"].values.sum()
    kopf("Planung je Marktbereich und Monat",
         f"Plan {de(p6 / T)} t, {pct(p6 / f6 - 1)} zur Modellprognose", D)

    if not ss.get("nutzer"):
        st.info("Bitte links Rolle und Namen wählen, um den Plan zu bearbeiten.", icon=":material/person:")

    eigene = [mb for mb in D["mbs"] if darf_bearbeiten(mb)]
    andere = [mb for mb in D["mbs"] if mb not in eigene]
    spalten = [monat_label(m, D["n_kw"][m]) for m in D["monate"]]

    with st.container(border=True):
        c1, c2 = st.columns([3, 2], vertical_alignment="bottom")
        c1.markdown("### Ihr Plan in Tonnen")
        if eigene:
            pe, fe = ss["plan"].loc[eigene].values.sum(), D["prog"].loc[eigene].values.sum()
            c2.markdown(f"<p class='wp-satz' style='text-align:right'>Ihre {len(eigene)} Marktbereiche: <b>{de(pe / T)} t</b> "
                        f"Plan, {pct(pe / fe - 1) if fe else 'n/a'} zum Modell</p>", unsafe_allow_html=True)
        c1.caption("Monatswerte überschreiben, die App verteilt auf die Kalenderwochen im Verhältnis des bisherigen "
                   f"Wochenprofils. Ab {pct(SCHWELLE, False, 0)} Abweichung zur Modellprognose ist eine Begründung "
                   "Pflicht. Wochenwerte lassen sich im Marktbereich feinjustieren.")
        if not eigene:
            st.markdown("<p class='wp-hinweis'>Aktuell sind keine Marktbereiche für Sie bearbeitbar "
                        "(nicht zugeordnet, eingereicht oder freigegeben).</p>", unsafe_allow_html=True)
        else:
            df = pd.DataFrame({"Marktbereich": eigene})
            for m, lbl in zip(D["monate"], spalten):
                df[lbl] = (plan_m.loc[eigene, m] / T).round(1).values
            df["Plan 6 M"] = (ss["plan"].loc[eigene].sum(axis=1) / T).round(1).values
            df["Modell 6 M"] = (D["prog"].loc[eigene].sum(axis=1) / T).round(1).values
            df["Zum Modell"] = [abweichung(D, mb) * 100 for mb in eigene]
            df["Begründung"] = st_.loc[eigene, "Begruendung"].astype(str).values
            cfg = {"Marktbereich": st.column_config.TextColumn("MB", pinned=True, width=60, disabled=True),
                   "Plan 6 M": st.column_config.NumberColumn(format="%.1f", disabled=True, width=76),
                   "Modell 6 M": st.column_config.NumberColumn(format="%.1f", disabled=True, width=76),
                   "Zum Modell": st.column_config.NumberColumn(format="%+.1f %%", disabled=True, width=76),
                   "Begründung": st.column_config.TextColumn(width=220, max_chars=400)}
            for lbl in spalten:
                cfg[lbl] = st.column_config.NumberColumn(lbl, format="%.1f", min_value=0.0, step=0.1, width=76)
            ed = st.data_editor(df, hide_index=True, num_rows="fixed", column_config=cfg, row_height=38,
                                key=f"ed_plan_{ss['editor_v']}")
            geaendert = False
            for _, r in ed.iterrows():
                mb = r["Marktbereich"]
                for m, lbl in zip(D["monate"], spalten):
                    neu = r[lbl]
                    if pd.notna(neu) and abs(float(neu) - round(plan_m.loc[mb, m] / T, 1)) > 0.049:
                        setze_monat(D, mb, m, float(neu) * T)
                        geaendert = True
                grund = "" if pd.isna(r["Begründung"]) else str(r["Begründung"]).strip()
                if grund != str(st_.loc[mb, "Begruendung"]).strip():
                    st_.loc[mb, "Begruendung"] = grund
                    protokolliere("Begründung", mb, "", None, None, grund)
                    bearbeitet(mb)
                    geaendert = True
            if geaendert:
                st.rerun()

            aktionen_planer(D, eigene)

    if ss.get("rolle") == ROLLE_KOORD:
        freigabe_koordination(D)

    if andere:
        with st.container(border=True):
            st.markdown("### Weitere Marktbereiche")
            st.caption("Nur lesend: anderen Verantwortlichen zugeordnet, eingereicht oder freigegeben.")
            df = pd.DataFrame({"Marktbereich": andere, "Verantwortlich": st_.loc[andere, "Verantwortlich"].values,
                               "Status": st_.loc[andere, "Status"].values})
            for m, lbl in zip(D["monate"], spalten):
                df[lbl] = plan_m.loc[andere, m].values / T
            df["Plan 6 M"] = ss["plan"].loc[andere].sum(axis=1).values / T
            df["Zum Modell"] = [abweichung(D, mb) * 100 for mb in andere]
            df["Begründung"] = st_.loc[andere, "Begruendung"].values
            st.dataframe(df, hide_index=True, row_height=36, column_config={
                **{lbl: st.column_config.NumberColumn(format="localized") for lbl in spalten + ["Plan 6 M"]},
                "Zum Modell": st.column_config.NumberColumn(format="%+.1f %%"),
                "Begründung": st.column_config.TextColumn(width="large")})

    with st.container(border=True):
        st.markdown("### Gesamt je Monat")
        st.plotly_chart(chart_monate(D, plan_m.sum().values, prog_m.sum().values, wert_monat(D, D["vj"]).sum().values),
                        config=PLOT_CFG, key="pl_chart")


def aktionen_planer(D, eigene):
    st_ = ss["status"]
    c1, c2, c3 = st.columns([2, 2, 3], vertical_alignment="center")
    einreichbar = [mb for mb in eigene if st_.loc[mb, "Status"] in ("Offen", "In Bearbeitung", "Zurückgewiesen")]
    fehlt = [mb for mb in einreichbar if braucht_begruendung(D, mb)]
    with c1:
        if st.button("Zur Freigabe einreichen", type="primary", icon=":material/send:", width="stretch",
                     disabled=not einreichbar or bool(fehlt)):
            for mb in einreichbar:
                st_.loc[mb, "Status"] = "Eingereicht"
                st_.loc[mb, ["Geaendert_von", "Geaendert_am"]] = [ss.get("nutzer"), jetzt()]
                protokolliere("Eingereicht", mb, "", ss["plan"].loc[mb].sum(), None, st_.loc[mb, "Begruendung"])
            ss["editor_v"] += 1
            st.toast(f"{len(einreichbar)} Marktbereiche eingereicht.", icon=":material/check_circle:")
            st.rerun()
    with c2:
        with st.popover("Auf Modell zurücksetzen", icon=":material/restart_alt:", width="stretch"):
            mb = st.selectbox("Marktbereich", eigene, key="reset_mb")
            if st.button("Plan dieses Marktbereichs zurücksetzen", key="reset_go"):
                alt = ss["plan"].loc[mb].sum()
                ss["plan"].loc[mb] = D["prog"].loc[mb]
                protokolliere("Auf Modell zurückgesetzt", mb, "", alt, ss["plan"].loc[mb].sum())
                bearbeitet(mb)
                ss["editor_v"] += 1
                st.rerun()
    with c3:
        if fehlt:
            st.markdown(f"<span class='wp-hinweis'>Einreichen erst mit Begründung für MB "
                        f"{', '.join(map(h, fehlt))}.</span>", unsafe_allow_html=True)
        elif einreichbar:
            st.markdown(f"<span class='wp-hinweis'>Reicht {len(einreichbar)} Marktbereiche zur Freigabe ein; "
                        "danach sind sie für Sie gesperrt.</span>", unsafe_allow_html=True)


def freigabe_koordination(D):
    st_ = ss["status"]
    eingereicht = [mb for mb in D["mbs"] if st_.loc[mb, "Status"] == "Eingereicht"]
    frei = [mb for mb in D["mbs"] if st_.loc[mb, "Status"] == "Freigegeben"]
    with st.container(border=True):
        st.markdown("### Freigabe")
        if not eingereicht:
            st.markdown("<p class='wp-hinweis'>Keine Einreichungen offen. Als Koordination können Sie Marktbereiche "
                        "auch direkt oben bearbeiten und hier freigeben.</p>", unsafe_allow_html=True)
        kandidaten = [mb for mb in D["mbs"] if st_.loc[mb, "Status"] != "Freigegeben"]
        c1, c2 = st.columns([3, 2])
        with c1:
            wahl = st.multiselect("Marktbereiche", kandidaten, default=eingereicht, key=f"fg_wahl_{ss['editor_v']}",
                                  placeholder="Marktbereiche wählen",
                                  format_func=lambda mb: f"{mb}  |  {st_.loc[mb, 'Status']}  |  "
                                                         f"{pct(abweichung(D, mb))} zum Modell")
        kommentar = c2.text_input("Kommentar der Koordination", key="fg_kommentar",
                                  placeholder="Pflicht bei Zurückweisung")
        b1, b2, b3 = st.columns(3)
        if b1.button("Freigeben", type="primary", icon=":material/verified:", width="stretch", disabled=not wahl):
            fehlt = [mb for mb in wahl if braucht_begruendung(D, mb)]
            if fehlt:
                st.error(f"Begründung fehlt bei MB {', '.join(fehlt)}.")
            else:
                for mb in wahl:
                    st_.loc[mb, ["Status", "Kommentar_Koordination", "Geaendert_von", "Geaendert_am"]] = [
                        "Freigegeben", kommentar, ss.get("nutzer"), jetzt()]
                    protokolliere("Freigegeben", mb, "", None, ss["plan"].loc[mb].sum(), kommentar)
                ss["editor_v"] += 1
                st.rerun()
        if b2.button("Zurückweisen", icon=":material/undo:", width="stretch", disabled=not wahl or not kommentar):
            for mb in wahl:
                st_.loc[mb, ["Status", "Kommentar_Koordination", "Geaendert_von", "Geaendert_am"]] = [
                    "Zurückgewiesen", kommentar, ss.get("nutzer"), jetzt()]
                protokolliere("Zurückgewiesen", mb, "", None, None, kommentar)
            ss["editor_v"] += 1
            st.rerun()
        with b3.popover("Freigabe aufheben", icon=":material/lock_open:", width="stretch", disabled=not frei):
            mb = st.selectbox("Marktbereich", frei, key="fg_auf")
            if st.button("Wieder zur Bearbeitung öffnen", key="fg_auf_go"):
                st_.loc[mb, "Status"] = "In Bearbeitung"
                protokolliere("Freigabe aufgehoben", mb, "", None, None, "")
                ss["editor_v"] += 1
                st.rerun()


# ============================================================================
# SEITE: MARKTBEREICH (Detail und Drilldown)
# ============================================================================
def seite_mb():
    D = ss.get("D")
    if D is None:
        return keine_daten()
    st_ = ss["status"]
    mbs = D["mbs"]
    if ss.get("mb_aktiv") not in mbs:
        ss["mb_aktiv"] = mbs[0]
    mb = ss["mb_aktiv"]
    info = D["mbinfo"].loc[mb]
    plan, prog = ss["plan"].loc[mb], D["prog"].loc[mb]
    p6, f6, vj6 = plan.sum(), prog.sum(), D["vj"].loc[mb].sum()
    kopf(f"Marktbereich {mb}",
         f"Plan {de(p6 / T)} t, {pct(p6 / f6 - 1) if f6 else 'n/a'} zum Modell, "
         f"{pct(p6 / vj6 - 1) if vj6 else 'n/a'} ggü. Vorjahr", D)

    c1, c2, c3 = st.columns([2, 3, 2], vertical_alignment="center")
    neu = c1.selectbox("Marktbereich", mbs, index=mbs.index(mb), label_visibility="collapsed",
                       format_func=lambda m: f"MB {m}  |  {st_.loc[m, 'Verantwortlich']}")
    if neu != mb:
        ss["mb_aktiv"] = neu
        st.rerun()
    c2.markdown(f"{status_pill(st_.loc[mb, 'Status'])} &nbsp; <span class='wp-hinweis'>Verantwortlich "
                f"{h(st_.loc[mb, 'Verantwortlich'])}; zuletzt {h(st_.loc[mb, 'Geaendert_von']) or 'unverändert'} "
                f"{h(st_.loc[mb, 'Geaendert_am'])}</span>", unsafe_allow_html=True)
    if str(st_.loc[mb, "Kommentar_Koordination"]).strip():
        c3.markdown(f"<span class='wp-hinweis'>Koordination: {h(st_.loc[mb, 'Kommentar_Koordination'])}</span>",
                    unsafe_allow_html=True)

    k = st.columns(4)
    l26 = D["ist"].loc[mb].iloc[-26:].sum()
    k[0].metric("Plan 6 Monate", f"{de(p6 / T)} t", f"Modell {de(f6 / T)} t", border=True, delta_color="off", delta_arrow="off")
    k[1].metric("Ggü. Vorjahr", pct(p6 / vj6 - 1) if vj6 else "n/a", f"Vorjahr {de(vj6 / T)} t", border=True,
                delta_color="off", delta_arrow="off")
    k[2].metric("Ggü. letzte 26 KW", pct(p6 / l26 - 1) if l26 else "n/a", f"Ist {de(l26 / T)} t", border=True,
                delta_color="off", delta_arrow="off")
    verl = str(info.get("Verlaesslichkeit", ""))
    k[3].metric("Prognosegüte", verl.capitalize() if verl else "n/a",
                f"Fehler je 4 Wochen {pct(float(info['Fehler_4W']), False)}" if pd.notna(info.get("Fehler_4W"))
                else None, border=True, delta_color="off", delta_arrow="off")

    with st.container(border=True):
        st.markdown("### Verlauf je KW")
        st.plotly_chart(chart_wochen(D["ist"].loc[mb], prog, D["bu"].loc[mb], D["bo"].loc[mb], plan,
                                     vj=D["vj"].loc[mb]), config=PLOT_CFG, key="mb_chart")

    tabs = st.tabs(["Wochenplan", "Material", "Länder", "Kunden", "Prognosemodell"], key="mb_tabs")
    with tabs[0]:
        tab_wochen(D, mb)
    mat = D["mat"][D["mat"]["MB"] == mb].copy()
    det = D["det"][D["det"]["MB"] == mb].copy()
    with tabs[1]:
        tab_material(D, mb, mat)
    with tabs[2]:
        tab_dimension(det, mat, "Landname", "Land", mb)
    with tabs[3]:
        tab_dimension(det, mat, "Auftraggeber", "Kunde", mb)
    with tabs[4]:
        tab_modell(D, mb, info)


def tab_wochen(D, mb):
    darf = darf_bearbeiten(mb)
    W = D["W"]
    df = pd.DataFrame({"KW": [kw_label(w) for w in W], "Monat": [monat_label(D["monat"][w]) for w in W],
                       "Vorjahr": D["vj"].loc[mb, W].values / T, "Modell": D["prog"].loc[mb, W].values / T,
                       "Band unten": D["bu"].loc[mb, W].values / T, "Band oben": D["bo"].loc[mb, W].values / T,
                       "Plan": (ss["plan"].loc[mb, W].values / T).round(1)})
    df["Plan zu Modell"] = (df["Plan"] / df["Modell"].replace(0, np.nan) - 1) * 100
    num = st.column_config.NumberColumn(format="localized", disabled=True)
    cfg = {"KW": st.column_config.TextColumn(disabled=True, width="small"),
           "Monat": st.column_config.TextColumn(disabled=True, width="small"),
           "Vorjahr": num, "Modell": num, "Band unten": num, "Band oben": num,
           "Plan": st.column_config.NumberColumn("Plan (t)", format="localized", min_value=0.0, step=0.1,
                                                 disabled=not darf),
           "Plan zu Modell": st.column_config.NumberColumn(format="%+.1f %%", disabled=True)}
    st.caption("Werte in Tonnen. " + ("Plan je KW direkt überschreibbar; Monatssummen passen sich an."
                                      if darf else "Nur lesend: Marktbereich nicht für Sie bearbeitbar."))
    ed = st.data_editor(df, hide_index=True, column_config=cfg, key=f"ed_kw_{mb}_{ss['editor_v']}",
                        disabled=not darf, height=min(38 * 27 + 40, 760), row_height=34)
    if darf:
        aend = False
        for i, w in enumerate(W):
            v = ed.loc[i, "Plan"]
            if pd.notna(v) and abs(float(v) - df.loc[i, "Plan"]) > 0.049:
                setze_woche(D, mb, w, float(v) * T)
                aend = True
        if aend:
            st.rerun()
    st_ = ss["status"]
    if darf:
        g = st.text_area("Begründung für den Marktbereich", value=str(st_.loc[mb, "Begruendung"]),
                         key=f"grund_{mb}_{ss['editor_v']}", height=80,
                         placeholder="Was weiß der Markt, was das Modell nicht weiß? z. B. Neukunde, Projektgeschäft, "
                                     "Preiserhöhung, Lageraufbau beim Kunden")
        if g.strip() != str(st_.loc[mb, "Begruendung"]).strip():
            st_.loc[mb, "Begruendung"] = g.strip()
            protokolliere("Begründung", mb, "", None, None, g.strip())
            bearbeitet(mb)
            st.rerun()


def tab_material(D, mb, mat):
    if mat.empty:
        st.info("Keine Materialdaten für diesen Marktbereich.")
        return
    p6 = ss["plan"].loc[mb].sum()
    mat["Plan 6 M"] = mat["Anteil"] * p6
    mat["Delta_kg"] = mat["Ist_52W_kg"] - mat["Ist_VJ_52W_kg"]
    n_80 = int((mat.sort_values("Prognose_26W_kg", ascending=False)["Anteil"].cumsum() < 0.8).sum()) + 1
    top = mat.loc[mat["Delta_kg"].idxmax()]
    tief = mat.loc[mat["Delta_kg"].idxmin()]
    st.markdown(f"<p class='wp-satz'><b>{n_80} von {len(mat)} Materialien</b> tragen 80 % des Plans. "
                f"Stärkster Zuwachs Ist ggü. Vorjahr: <b>{h(top['Material'])}</b> ({de(top['Delta_kg'] / T, 1)} t), "
                f"stärkster Rückgang: <b>{h(tief['Material'])}</b> ({de(tief['Delta_kg'] / T, 1)} t).</p>",
                unsafe_allow_html=True)
    c1, c2 = st.columns([3, 2], gap="medium")
    with c1:
        such = st.text_input("Material suchen", key=f"such_mat_{mb}", placeholder="Materialnummer",
                             label_visibility="collapsed")
        t = mat.sort_values("Plan 6 M", ascending=False)
        if such:
            t = t[t["Material"].str.contains(such.strip(), case=False, regex=False)]
        out = pd.DataFrame({"Material": t["Material"], "Anteil im MB": t["Anteil"] * 100,
                            "Plan 6 M (t)": t["Plan 6 M"] / T, "Modell 6 M (t)": t["Prognose_26W_kg"] / T,
                            "Ist 52 KW (t)": t["Ist_52W_kg"] / T, "Vorjahr 52 KW (t)": t["Ist_VJ_52W_kg"] / T,
                            "Ist zu Vorjahr": (t["Ist_52W_kg"] / t["Ist_VJ_52W_kg"].replace(0, np.nan) - 1) * 100})
        f1 = st.column_config.NumberColumn(format="%.1f")
        st.dataframe(out, hide_index=True, height=420, column_config={
            "Anteil im MB": st.column_config.ProgressColumn(format="%.1f %%", min_value=0,
                                                            max_value=float(max(out["Anteil im MB"].max(), 1)),
                                                            color=MITTEL),
            "Plan 6 M (t)": f1, "Modell 6 M (t)": f1, "Ist 52 KW (t)": f1, "Vorjahr 52 KW (t)": f1,
            "Ist zu Vorjahr": st.column_config.NumberColumn(format="%+.1f %%")})
        st.caption("Plan je Material = Plan des Marktbereichs x gewichteter Materialanteil der letzten 52 KW "
                   "(rechnerische Verteilung, keine eigene Materialprognose).")
    with c2:
        st.markdown("**Treiber der Veränderung**")
        st.plotly_chart(chart_treiber(mat, "Material", "Delta_kg"), config=PLOT_CFG, key=f"tr_mat_{mb}")


def tab_dimension(det, mat, spalte, name, mb):
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
    c1, c2 = st.columns([3, 2], gap="medium")
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
    with c2:
        st.markdown("**Treiber der Veränderung**")
        st.plotly_chart(chart_treiber(g.reset_index(), spalte, "Delta_kg"), config=PLOT_CFG,
                        key=f"tr_{spalte}_{mb}")
    st.caption("Ist-Werte der letzten 52 KW gegenüber den 52 KW davor; Basis Fakturadatum, Einheit KG.")


def tab_modell(D, mb, info):
    def f(x):
        try:
            return float(x)
        except (TypeError, ValueError):
            return np.nan
    c = st.columns(4)
    c[0].metric("Gewähltes Modell", str(info.get("Modell", "n/a")), f"XYZ-Klasse {info.get('XYZ', 'n/a')}",
                border=True, delta_color="off", delta_arrow="off")
    c[1].metric("Fehler je 4 Wochen", pct(f(info.get("Fehler_4W")), False),
                f"Vorjahresmethode {pct(f(info.get('Fehler_4W_Vorjahresmethode')), False)}", border=True,
                delta_color="off", delta_arrow="off")
    c[2].metric("Systematische Abweichung", pct(f(info.get("Bias"))), "Holdout 26 KW", border=True, delta_color="off", delta_arrow="off")
    drift = str(info.get("Drift", "")).lower() in ("true", "1", "wahr")
    c[3].metric("Drift", "ja" if drift else "nein", f"t-Wert {de(f(info.get('Drift_t')), 1)}", border=True,
                delta_color="off", delta_arrow="off")
    st.markdown(
        "<p class='wp-hinweis'>Die Güte stammt aus einem ehrlichen Backtest: Modellwahl auf früheren Zeitpunkten, "
        "Messung auf den jüngsten 26 KW, die die Modellwahl nicht kannte. Bias größer null heißt: Modell lag zu hoch. "
        "Drift heißt: in den letzten 13 KW systematisch daneben, Plan dort mit dem Markt besonders prüfen. "
        "Die Bandbreite im Diagramm umfasst 80 % der im Backtest beobachteten Abweichungen.</p>",
        unsafe_allow_html=True)


# ============================================================================
# SEITE: ABSCHLUSS (Sichern, Zustaendigkeiten, Protokoll)
# ============================================================================
def seite_abschluss():
    D = ss.get("D")
    if D is None:
        return keine_daten()
    st_ = ss["status"]
    n_frei = int((st_["Status"] == "Freigegeben").sum())
    kopf("Planstand sichern", f"{n_frei} von {len(st_)} Marktbereichen freigegeben", D)
    t1, t2, t3 = st.tabs(["Sichern", "Zuständigkeiten", "Protokoll"])
    with t1:
        c1, c2 = st.columns([3, 2], gap="large")
        with c1:
            with st.container(border=True):
                st.markdown("### Planstand als Excel")
                final = n_frei == len(st_)
                st.markdown(
                    "<p class='wp-satz'>" + ("Alle Marktbereiche sind freigegeben. Diese Datei ist der <b>finale Plan</b> "
                                             "der Runde und Startpunkt der nächsten." if final else
                                             "Zwischenstand: Datei sichern, weitergeben und später hier wieder hochladen, "
                                             "um weiterzuarbeiten.") + "</p>", unsafe_allow_html=True)
                name = ("final" if final else "zwischenstand")
                st.download_button("Planstand herunterladen", data=planstand_excel(D, final),
                                   file_name=f"excel_w&p_{app_name().split(' ', 1)[-1].lower().replace(' ', '_') or 's&op'}_absatzplan_{name}_"
                                             f"{datetime.now():%y%m%d_%H%M}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   type="primary", icon=":material/download:", on_click="ignore")
                st.caption("Enthält Plan je Monat und KW, Vergleich mit Modell und Vorjahr, Status, Begründungen und "
                           "das vollständige Änderungsprotokoll.")
        with c2:
            with st.container(border=True):
                st.markdown("### Wo liegen die Daten?")
                st.markdown(
                    "<p class='wp-hinweis'>Die App läuft nur auf diesem Rechner. Hochgeladene Dateien liegen im "
                    "Arbeitsspeicher und sind nach dem Schließen weg. Gespeichert wird ausschließlich die Excel, "
                    "die Sie hier herunterladen, an dem Ort, den Sie wählen.</p>", unsafe_allow_html=True)
    with t2:
        with st.container(border=True):
            st.markdown("### Wer plant welchen Marktbereich?")
            if ss.get("rolle") != ROLLE_KOORD:
                st.caption("Nur die S&OP-Koordination kann Zuständigkeiten ändern.")
            df = pd.DataFrame({"Marktbereich": D["mbs"],
                               "Verantwortlich": [ss["team"].get(mb, OHNE) for mb in D["mbs"]],
                               "Anteil": (D["prog"].sum(axis=1) / D["prog"].values.sum() * 100).values})
            ed = st.data_editor(df, hide_index=True, disabled=ss.get("rolle") != ROLLE_KOORD,
                                key=f"ed_team_{ss['editor_v']}", column_config={
                                    "Marktbereich": st.column_config.TextColumn(disabled=True),
                                    "Verantwortlich": (st.column_config.SelectboxColumn(
                                        options=[OHNE] + sorted(n for n, e in _nutzer().items() if e["rolle"] == ROLLE_MB))
                                        if ZUGANG["aktiv"] else st.column_config.TextColumn(max_chars=60)),
                                    "Anteil": st.column_config.ProgressColumn(format="%.0f %%", min_value=0,
                                                                              max_value=100, color=MITTEL)})
            for _, r in ed.iterrows():
                v = str(r["Verantwortlich"] or "").strip() or OHNE
                if v != ss["team"].get(r["Marktbereich"]):
                    protokolliere("Zuständigkeit", r["Marktbereich"], "", None, None,
                                  f"{ss['team'].get(r['Marktbereich'])} zu {v}")
                    ss["team"][r["Marktbereich"]] = v
                    st_.loc[r["Marktbereich"], "Verantwortlich"] = v
    with t3:
        log = pd.DataFrame(ss["log"])
        if log.empty:
            st.info("Noch keine Änderungen.")
        else:
            c1, c2 = st.columns([2, 3])
            f = c1.multiselect("Marktbereich", sorted(x for x in log["Marktbereich"].unique() if x),
                               placeholder="Alle", key="log_mb")
            if f:
                log = log[log["Marktbereich"].isin(f)]
            st.dataframe(log.iloc[::-1], hide_index=True, height=520, column_config={
                "Vorher_t": st.column_config.NumberColumn("Vorher (t)", format="localized"),
                "Nachher_t": st.column_config.NumberColumn("Nachher (t)", format="localized"),
                "Kommentar": st.column_config.TextColumn(width="large")})


def planstand_excel(D, final: bool) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter as L

    wb = Workbook()
    kopf_fill, band = PatternFill("solid", fgColor=NAVY[1:]), PatternFill("solid", fgColor=LILA[1:])
    kf, nf, bf = (Font(name="Calibri", bold=True, color="FFFFFF"), Font(name="Calibri", color=NAVY[1:]),
                  Font(name="Calibri", bold=True, color=NAVY[1:]))
    linie = Border(top=Side(style="medium", color=NAVY[1:]))

    def tabelle(ws, df, start=4, fmt=None, titel="", kern=""):
        ws.sheet_view.showGridLines = False
        ws["A1"], ws["A2"] = titel, kern
        ws["A1"].font = Font(name="Calibri", size=16, bold=True, color=NAVY[1:])
        ws["A2"].font = Font(name="Calibri", size=11, color=MITTEL[1:])
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
                if (i - start) % 2 == 0:
                    z.fill = band
                if fmt and df.columns[j - 1] in fmt:
                    z.number_format = fmt[df.columns[j - 1]]
                elif isinstance(v, (float, np.floating)):
                    z.number_format = "#,##0.0"
        for j, c in enumerate(df.columns, 1):
            ws.column_dimensions[L(j)].width = max(11, min(48, len(str(c)) + 4))
        ws.freeze_panes = ws.cell(row=start + 1, column=2)
        return start + len(df)

    st_ = ss["status"]
    plan_m, prog_m, vj_m = plan_monat(D), wert_monat(D, D["prog"]), wert_monat(D, D["vj"])
    spalten = [monat_label(m, D["n_kw"][m]) for m in D["monate"]]
    # Blatt 1: Plan MB x Monat
    df = pd.DataFrame({"Marktbereich": D["mbs"], "Verantwortlich": st_.loc[D["mbs"], "Verantwortlich"].values,
                       "Status": st_.loc[D["mbs"], "Status"].values})
    for m, lbl in zip(D["monate"], spalten):
        df[lbl] = plan_m.loc[D["mbs"], m].values / T
    df["Plan 6 M"] = ss["plan"].loc[D["mbs"]].sum(axis=1).values / T
    df["Modell 6 M"] = D["prog"].loc[D["mbs"]].sum(axis=1).values / T
    df["Vorjahr"] = D["vj"].loc[D["mbs"]].sum(axis=1).values / T
    df["Plan zu Modell"] = df["Plan 6 M"] / df["Modell 6 M"].replace(0, np.nan) - 1
    df["Plan zu Vorjahr"] = df["Plan 6 M"] / df["Vorjahr"].replace(0, np.nan) - 1
    df["Begründung"] = st_.loc[D["mbs"], "Begruendung"].values
    summe = {c: df[c].sum() if c in spalten + ["Plan 6 M", "Modell 6 M", "Vorjahr"] else "" for c in df.columns}
    summe["Marktbereich"] = "Gesamt"
    summe["Plan zu Modell"] = summe["Plan 6 M"] / summe["Modell 6 M"] - 1
    summe["Plan zu Vorjahr"] = summe["Plan 6 M"] / summe["Vorjahr"] - 1 if summe["Vorjahr"] else None
    df = pd.concat([df, pd.DataFrame([summe])], ignore_index=True)
    ws = wb.active
    ws.title = "Plan MB Monat"
    fmt = {c: '#,##0" t"' for c in spalten + ["Plan 6 M", "Modell 6 M", "Vorjahr"]}
    fmt.update({"Plan zu Modell": "+0.0%;-0.0%;0.0%", "Plan zu Vorjahr": "+0.0%;-0.0%;0.0%"})
    ende = tabelle(ws, df, fmt=fmt, titel=app_name() + " | " + ("Finaler Plan" if final else "Zwischenstand"),
                   kern=f"Datenstand {D['meta'].get('Datenstand', '')}; Planhorizont {kw_label(D['W'][0])} bis "
                        f"{kw_label(D['W'][-1])}; exportiert {jetzt()} von {ss.get('nutzer') or 'unbekannt'}")
    for j in range(1, len(df.columns) + 1):
        ws.cell(row=ende, column=j).font = bf
        ws.cell(row=ende, column=j).border = linie
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions[L(len(df.columns))].width = 60

    # Blatt 2: Plan MB x KW
    kw = pd.DataFrame({"Marktbereich": D["mbs"]})
    for w in D["W"]:
        kw[kw_label(w)] = ss["plan"].loc[D["mbs"], w].values / T
    kw = pd.concat([kw, pd.DataFrame([{"Marktbereich": "Gesamt", **{kw_label(w): ss["plan"][w].sum() / T
                                                                     for w in D["W"]}}])], ignore_index=True)
    ws = wb.create_sheet("Plan MB KW")
    ende = tabelle(ws, kw, titel="Plan je Marktbereich und KW (t)", kern="Grundlage für Kapazitäts- und Rohstoffplanung")
    for j in range(1, len(kw.columns) + 1):
        ws.cell(row=ende, column=j).font = bf
        ws.cell(row=ende, column=j).border = linie

    # Blatt 3: Vergleich je Monat
    vg = pd.DataFrame({"Monat": spalten, "Plan": plan_m.sum().values / T, "Modell": prog_m.sum().values / T,
                       "Vorjahr": vj_m.sum().values / T})
    vg["Plan zu Modell"] = vg["Plan"] / vg["Modell"].replace(0, np.nan) - 1
    vg["Plan zu Vorjahr"] = vg["Plan"] / vg["Vorjahr"].replace(0, np.nan) - 1
    ws = wb.create_sheet("Vergleich Monat")
    tabelle(ws, vg, fmt={"Plan": '#,##0" t"', "Modell": '#,##0" t"', "Vorjahr": '#,##0" t"',
                         "Plan zu Modell": "+0.0%;-0.0%;0.0%", "Plan zu Vorjahr": "+0.0%;-0.0%;0.0%"},
            titel="Gesamt je Monat", kern="Plan, Modellprognose und gleiche KW des Vorjahres")

    ws = wb.create_sheet("Protokoll")
    tabelle(ws, pd.DataFrame(ss["log"]).fillna(""), titel="Änderungsprotokoll", kern="Alle Änderungen dieser Runde")

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
# SEITENLEISTE UND NAVIGATION
# ============================================================================
def seitenleiste():
    with st.sidebar:
        D = ss.get("D")
        if D is None:
            st.markdown(f"<div class='wp-marke'><b>W&amp;P</b><span>{h(app_name())}</span></div>",
                        unsafe_allow_html=True)
        if ZUGANG["aktiv"]:
            st.markdown(f"<div class='wp-schutz'>Angemeldet als <b>{h(ss['nutzer'])}</b><br>{h(ss['rolle'])}</div>",
                        unsafe_allow_html=True)
            st.space("small")
            if st.button("Abmelden", icon=":material/logout:", width="stretch", key="abmelden"):
                abmelden("Sie wurden abgemeldet.")
        if D is None:
            if not ZUGANG["aktiv"]:
                st.markdown("<div class='wp-schutz'>Läuft nur auf diesem Rechner. Dateien bleiben im Arbeitsspeicher, "
                            "nichts wird gespeichert oder übertragen.</div>", unsafe_allow_html=True)
            return
        if not ZUGANG["aktiv"]:
            st.markdown("**Ich arbeite als**")
            rolle = st.selectbox("Rolle", [ROLLE_MB, ROLLE_KOORD],
                                 index=[ROLLE_MB, ROLLE_KOORD].index(ss.get("rolle", ROLLE_MB)),
                                 key="w_rolle", label_visibility="collapsed")
            ss["rolle"] = rolle or ROLLE_MB
            if ss["rolle"] == ROLLE_MB:
                namen = sorted({v for v in ss["team"].values() if v and v != OHNE})
                if namen:
                    idx = namen.index(ss["nutzer"]) if ss.get("nutzer") in namen else None
                    ss["nutzer"] = st.selectbox("Name", namen, index=idx, placeholder="Namen wählen",
                                                label_visibility="collapsed", key="w_name")
                else:
                    st.caption("Noch keine Zuständigkeiten hinterlegt. Die Koordination pflegt sie unter "
                               "Planstand sichern.")
                    ss["nutzer"] = None
            else:
                ss["nutzer"] = st.text_input("Name", value=ss.get("nutzer_k", ""), placeholder="Ihr Name",
                                             label_visibility="collapsed", key="w_name_k") or None
                ss["nutzer_k"] = ss["nutzer"] or ""
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
                   if ZUGANG["aktiv"] else "Nur lokal. Nichts wird gespeichert, außer Sie laden den Planstand herunter.")
        st.markdown(f"<div class='wp-schutz'>{hinweis}</div>", unsafe_allow_html=True)


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


SEITEN = {
    "start": st.Page(seite_start, title="Start", url_path="start"),
    "cockpit": st.Page(seite_cockpit, title="Cockpit", url_path="cockpit"),
    "planung": st.Page(seite_planung, title="Planung", url_path="planung"),
    "mb": st.Page(seite_mb, title="Marktbereich", url_path="marktbereich"),
    "abschluss": st.Page(seite_abschluss, title="Planstand sichern", url_path="sichern"),
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
    return f"Absatzplanung {k}".strip()


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
        st.markdown("<div class='wp-start'><h1 style='font-size:2.1rem'>Absatzplanung</h1>"
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
    nav = st.navigation([SEITEN["cockpit"], SEITEN["planung"], SEITEN["mb"], SEITEN["abschluss"]],
                        position="sidebar")
nav.run()
