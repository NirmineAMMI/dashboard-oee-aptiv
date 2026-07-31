"""
Dashboard OEE - Aptiv (version design "cockpit" sombre + stockage persistant Supabase)
=======================================================================================
Lit automatiquement tous les fichiers Excel journaliers déposés dans un bucket
Supabase Storage (chacun avec les onglets "DATA 1"/"DATA 2"/"DT" -- OU
"SUMMARY"/"DT_LEGEND", les deux noms sont acceptés), calcule Availability,
Performance, Quality et OEE, et affiche un dashboard interactif filtrable par
Date, Shift (A/B/C) et CC.

Lancer avec :  streamlit run app.py

IMPORTANT — Persistance des fichiers :
Sur Streamlit Community Cloud, le disque local est éphémère (il est effacé à
chaque mise en veille / redéploiement). Les fichiers uploadés sont donc
stockés dans un bucket Supabase Storage plutôt que sur le disque, ce qui les
rend permanents. Voir les instructions de configuration fournies séparément
(secrets.toml + requirements.txt).

Design :
--------
Cette version reprend l'agencement d'un dashboard "cockpit" professionnel
(cartes KPI en tête, jauges/donuts circulaires, panneau de répartition,
mini-graphiques de tendance) afin que l'essentiel soit visible en un seul
écran. Les analyses détaillées (Pareto des arrêts, cycle time, données
brutes) restent disponibles dans des volets dépliables juste en dessous.
"""

import io
import os
import re
import unicodedata

import pandas as pd  # type: ignore
import plotly.express as px  # type: ignore
import plotly.graph_objects as go  # type: ignore
import streamlit as st  # type: ignore
from supabase import create_client  # type: ignore

# ----------------------------------------------------------------------------
# Configuration générale de la page
# ----------------------------------------------------------------------------
st.set_page_config(page_title="Dashboard OEE - Aptiv", layout="wide", page_icon="📊")

# ---- Palette "cockpit" sombre ----
BG = "#0e1330"
CARD_BG = "#161b3d"
CARD_BG_ALT = "#12173a"
CARD_BORDER = "#2a3160"
TEXT_MUTED = "#9aa4c7"
TEXT_MUTED_2 = "#7d87ad"
TRACK = "#232a55"

APTIV_BLUE = "#4A90D9"
APTIV_BLUE_DARK = "#0033A0"
GREEN = "#2ECC71"
ORANGE = "#F5A623"
RED = "#E74C3C"

st.markdown(
    f"""
    <style>
    .block-container {{padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1500px;}}
    #MainMenu, footer {{visibility: hidden;}}

    .dash-title {{
        font-size: 30px; font-weight: 800; color: #ffffff; letter-spacing: .3px;
        margin-bottom: 0px;
    }}
    .dash-subtitle {{
        font-size: 13px; color: {TEXT_MUTED}; margin-top: 2px; margin-bottom: 14px;
    }}
    .section-title {{
        font-size: 13px; font-weight: 700; color: #ffffff; text-transform: uppercase;
        letter-spacing: .8px; margin: 6px 0 10px 0; padding-left: 10px;
        border-left: 3px solid {APTIV_BLUE};
    }}

    /* ---- Cartes KPI (hero) ---- */
    .kpi-card {{
        background: {CARD_BG}; border: 1px solid {CARD_BORDER}; border-radius: 14px;
        padding: 16px 18px; height: 108px; display:flex; flex-direction:column; justify-content:space-between;
    }}
    .kpi-title {{
        font-size: 12px; color: {TEXT_MUTED}; font-weight: 700; text-transform: uppercase; letter-spacing: .6px;
    }}
    .kpi-row {{display:flex; align-items:center; justify-content:space-between;}}
    .kpi-value {{font-size: 32px; font-weight: 800; color: #ffffff; line-height: 1;}}
    .kpi-badge {{
        padding: 5px 11px; border-radius: 8px; font-weight: 700; font-size: 13px; white-space: nowrap;
    }}
    .kpi-subtitle {{font-size: 11.5px; color: {TEXT_MUTED_2};}}

    /* ---- Bloc "santé machines" (façon Promoters/Passives/Detractors) ---- */
    .health-card {{
        background: {CARD_BG}; border: 1px solid {CARD_BORDER}; border-radius: 14px;
        padding: 14px 18px; height: 100%;
    }}
    .health-row {{
        display:flex; align-items:center; gap: 12px; padding: 9px 0;
        border-bottom: 1px solid {TRACK};
    }}
    .health-row:last-child {{border-bottom:none;}}
    .health-dot {{
        width: 38px; height: 38px; min-width:38px; border-radius: 50%;
        display:flex; align-items:center; justify-content:center; font-size: 18px;
    }}
    .health-value {{font-size: 21px; font-weight: 800; color: #ffffff; line-height:1.1;}}
    .health-label {{font-size: 11.5px; color: {TEXT_MUTED};}}

    /* ---- Cadre neutre autour d'un graphique Plotly pour l'effet "carte" ---- */
    div[data-testid="stPlotlyChart"] {{
        background: {CARD_BG}; border: 1px solid {CARD_BORDER}; border-radius: 14px;
        padding: 6px 6px 0 6px;
    }}
    div[data-testid="stMetric"] {{
        background-color: {CARD_BG};
        border: 1px solid {CARD_BORDER};
        border-radius: 10px;
        padding: 12px 16px;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ----------------------------------------------------------------------------
# Fonctions de lecture / nettoyage des données (inchangées)
# ----------------------------------------------------------------------------
def _clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _code_to_int(code) -> "int | None":
    """Convertit un code DT en entier, quel que soit son format d'origine :
    - chiffre simple : '0'..'9'          -> 0..9
    - lettre simple   : 'A'..'Z'         -> 10..35   (A=10, B=11, ... comme dans DT_LEGEND)
    - nombre complet  : '11', '21', '31' -> 11, 21, 31 (comme dans DATA 1)
    Cela permet de faire correspondre les deux façons d'écrire le même code
    entre l'onglet DATA 1 (toujours en nombre) et l'onglet DT_LEGEND
    (0-9 puis lettres A-U pour 10-30, et de nouveau en nombre au-delà, ex. 31).
    """
    if pd.isna(code):
        return None
    s = str(code).strip().upper()
    if s == "":
        return None
    if s.isdigit():
        return int(s)
    if len(s) == 1 and s.isalpha():
        return ord(s) - ord("A") + 10
    return None


def _find_col(df: pd.DataFrame, *keywords: str):
    """Trouve la 1ère colonne dont le nom (en minuscules) contient tous les mots-clés donnés."""
    for col in df.columns:
        cl = col.lower()
        if all(k in cl for k in keywords):
            return col
    return None


def _parse_fr_float(x):
    """Convertit un nombre écrit à la française ('15,32') en float (15.32)."""
    if pd.isna(x):
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace(",", ".")
    m = re.search(r"[-+]?\d*\.?\d+", s)
    return float(m.group()) if m else None


def _parse_target_cycle(x):
    """Parse une cellule 'Target - cycle time' du type '15,00 ± 1,50 s'.
    Retourne (cible, tolerance) en secondes, ou (None, None) si illisible."""
    if pd.isna(x):
        return None, None
    s = str(x).strip().replace(",", ".")
    m = re.search(r"([\d.]+)\s*±\s*([\d.]+)", s)
    if m:
        return float(m.group(1)), float(m.group(2))
    m2 = re.search(r"([\d.]+)", s)
    if m2:
        return float(m2.group(1)), 0.0
    return None, None


def _name_of(f) -> str:
    """Nom du fichier, qu'il s'agisse d'un chemin (str) ou d'un fichier uploadé (UploadedFile / BytesIO)."""
    return getattr(f, "name", None) or str(f)


def _reset(f):
    """Remet le curseur au début pour un fichier uploadé (relecture possible)."""
    if hasattr(f, "seek"):
        try:
            f.seek(0)
        except Exception:
            pass


def _first_matching_sheet(filepath, candidates: list) -> str | None:
    """Retourne le 1er nom d'onglet existant parmi une liste de noms possibles.
    Compare en ignorant les espaces en début/fin (les exports du logiciel machine
    ajoutent parfois une espace en trop, ex. 'DATA 1 ' au lieu de 'DATA 1')."""
    try:
        _reset(filepath)
        xls = pd.ExcelFile(filepath)
        stripped_map = {name.strip(): name for name in xls.sheet_names}
        for candidate in candidates:
            if candidate in stripped_map:
                return stripped_map[candidate]
    except Exception:
        pass
    return None


def read_raw_log_from_file(filepath) -> pd.DataFrame:
    """Lit le journal d'événements (onglet 'DATA 1', accepte aussi 'RAW_LOG')."""
    sheet = _first_matching_sheet(filepath, ["DATA 1", "RAW_LOG"])
    if sheet is None:
        st.warning(f"Aucun onglet journal (DATA 1 / RAW_LOG) trouvé dans {_name_of(filepath)}")
        return pd.DataFrame()

    _reset(filepath)
    df = pd.read_excel(filepath, sheet_name=sheet)
    df = _clean_columns(df)

    required = ["Machine", "Start", "End", "Downtime reason",
                "Downtime name", "run [h]", "confessed [h]", "Commentaire"]
    df = df[[c for c in required if c in df.columns]].copy()
    if "Commentaire" not in df.columns:
        df["Commentaire"] = ""
    df["Commentaire"] = df["Commentaire"].fillna("")

    df["Start"] = pd.to_datetime(df["Start"], format="%d.%m.%Y %H:%M:%S", errors="coerce")
    df["End"] = pd.to_datetime(df["End"], format="%d.%m.%Y %H:%M:%S", errors="coerce")
    df["run [h]"] = pd.to_numeric(df["run [h]"], errors="coerce").fillna(0)
    df["confessed [h]"] = pd.to_numeric(df["confessed [h]"], errors="coerce").fillna(0)

    # "Downtime reason" contient le CODE numérique (0, 1, 4, ... 31) dans cet onglet.
    # On le normalise en entier pour pouvoir le relier à DT_LEGEND plus tard.
    df["Downtime reason"] = df["Downtime reason"].apply(_code_to_int)

    df = df.dropna(subset=["Start"])
    df["Date"] = df["Start"].dt.date

    def _shift(hour):
        if 6 <= hour < 14:
            return "A"
        elif 14 <= hour < 22:
            return "B"
        return "C"

    df["Shift"] = df["Start"].dt.hour.apply(_shift)
    df["Fichier_Source"] = _name_of(filepath)
    return df


def read_summary_from_file(filepath) -> pd.DataFrame:
    """Lit le résumé quotidien (onglet 'DATA 2', accepte aussi 'SUMMARY').
    Tolérant aux variations de colonnes : 'Yield' ou 'Yield quantity', 'CC' et 'Date'
    optionnels, colonnes 'Article number'/'Article name' optionnelles."""
    sheet = _first_matching_sheet(filepath, ["DATA 2", "SUMMARY"])
    if sheet is None:
        st.warning(f"Aucun onglet résumé (DATA 2 / SUMMARY) trouvé dans {_name_of(filepath)}")
        return pd.DataFrame()

    _reset(filepath)
    df = pd.read_excel(filepath, sheet_name=sheet)
    df = _clean_columns(df)

    # Détection d'un éventuel tableau annexe Machine -> CC, placé à côté du tableau
    # principal dans le même onglet (colonnes 'machine' en minuscule + 'CC').
    # On le repère par la présence de DEUX colonnes nommées "machine" (insensible à la
    # casse) : la 1ère est le tableau principal, la 2nde est le tableau de correspondance.
    machine_cols = [c for c in df.columns if c.lower().strip() == "machine"]
    col_machine = machine_cols[0] if machine_cols else _find_col(df, "machine")
    col_yield = _find_col(df, "yield")
    col_scrap = next((c for c in df.columns if c.lower().strip() == "scrap"), None)
    col_efficiency = _find_col(df, "efficiency")
    col_target = _find_col(df, "target", "cycle")
    col_avg = _find_col(df, "average", "cycle")
    col_cc = next((c for c in df.columns if c.lower().strip() == "cc"), None)
    col_date = next((c for c in df.columns if c.lower().strip() == "date"), None)
    col_art_num = _find_col(df, "article", "number")
    col_art_name = _find_col(df, "article", "name")

    if col_machine is None or col_yield is None:
        st.warning(f"Colonnes Machine/Yield introuvables dans {_name_of(filepath)}")
        return pd.DataFrame()

    cc_lookup = None
    if len(machine_cols) > 1 and col_cc:
        side_machine_col = machine_cols[1]
        lookup_df = df[[side_machine_col, col_cc]].dropna()
        cc_lookup = dict(zip(
            lookup_df[side_machine_col].astype(str).str.strip(), lookup_df[col_cc],
        ))

    out = pd.DataFrame()
    out["Machine"] = df[col_machine]
    out["Yield"] = pd.to_numeric(df[col_yield], errors="coerce").fillna(0)
    out["Scrap"] = pd.to_numeric(df[col_scrap], errors="coerce").fillna(0) if col_scrap else 0
    out["Efficiency"] = pd.to_numeric(df[col_efficiency], errors="coerce") if col_efficiency else float("nan")
    if cc_lookup:
        out["CC"] = out["Machine"].astype(str).str.strip().map(cc_lookup).fillna("N/A")
    elif col_cc:
        out["CC"] = df[col_cc]
    else:
        out["CC"] = "N/A"
    out["Date"] = pd.to_datetime(df[col_date], errors="coerce").dt.date if col_date else pd.NaT
    out["Article number"] = df[col_art_num] if col_art_num else None
    out["Article name"] = df[col_art_name] if col_art_name else None

    # Parsing des temps de cycle ('15,00 ± 1,50 s' -> cible + tolérance en secondes)
    if col_target:
        parsed = df[col_target].apply(_parse_target_cycle)
        out["Target cycle (s)"] = [p[0] for p in parsed]
        out["Tolerance cycle (s)"] = [p[1] for p in parsed]
    else:
        out["Target cycle (s)"] = None
        out["Tolerance cycle (s)"] = None
    out["Average cycle (s)"] = df[col_avg].apply(_parse_fr_float) if col_avg else None

    out = out.dropna(subset=["Machine"])
    out = out[out["Machine"].astype(str).str.strip() != ""]
    out["Fichier_Source"] = _name_of(filepath)
    return out


def read_dt_legend_from_file(filepath) -> pd.DataFrame:
    """Lit la légende des codes d'arrêt (onglet 'DT', accepte aussi 'DT_LEGEND')."""
    sheet = _first_matching_sheet(filepath, ["DT", "DT_LEGEND"])
    if sheet is None:
        return pd.DataFrame(columns=["CodeDT", "CodeDT_num", "Libelle", "Categorie", "ColorHEX"])

    # L'onglet "DT" du logiciel machine a son en-tête en ligne 4 (pas en ligne 1)
    try:
        _reset(filepath)
        df = pd.read_excel(filepath, sheet_name=sheet, header=3)
        if not {"Downtime name", "Downtime reason"}.issubset({c.strip() if isinstance(c, str) else c for c in df.columns}):
            _reset(filepath)
            df = pd.read_excel(filepath, sheet_name=sheet)  # sinon, en-tête normale en ligne 1
    except Exception:
        return pd.DataFrame(columns=["CodeDT", "CodeDT_num", "Libelle", "Categorie", "ColorHEX"])

    df = _clean_columns(df)
    col_mapping = {}
    for col in df.columns:
        cl = col.lower()
        if "downtime name" in cl:
            col_mapping[col] = "CodeDT"
        elif "downtime reason" in cl:
            col_mapping[col] = "Libelle"
        elif "downtime code" in cl:
            col_mapping[col] = "Categorie"
        elif "colorhex" in cl or "color" in cl:
            col_mapping[col] = "ColorHEX"
    df = df.rename(columns=col_mapping)

    keep = [c for c in ["CodeDT", "Libelle", "Categorie", "ColorHEX"] if c in df.columns]
    if not keep:
        return pd.DataFrame(columns=["CodeDT", "CodeDT_num", "Libelle", "Categorie", "ColorHEX"])
    df = df[keep].dropna(subset=["CodeDT"])
    df["CodeDT"] = df["CodeDT"].astype(str).str.strip()

    # Clé de fusion normalisée : convertit '0'-'9' et 'A'-'Z' en entier (0-35),
    # ce qui permet de la faire correspondre au code numérique de DATA 1
    # (ex. lettre 'L' dans DT_LEGEND == 21 dans DATA 1).
    df["CodeDT_num"] = df["CodeDT"].apply(_code_to_int)

    if "ColorHEX" in df.columns:
        df["ColorHEX"] = df["ColorHEX"].apply(
            lambda x: f"#{x}" if isinstance(x, str) and x and not x.startswith("#") else x
        )
    else:
        df["ColorHEX"] = "#9E9E9E"
    return df


# ----------------------------------------------------------------------------
# Stockage persistant (Supabase Storage) - inchangé
# ----------------------------------------------------------------------------
SUPABASE_BUCKET = "oee-files"


def _sanitize_filename(name: str) -> str:
    """Nettoie un nom de fichier pour qu'il soit accepté par Supabase Storage :
    enlève les accents (é -> e), remplace les espaces/caractères spéciaux par '_'.
    """
    base, ext = os.path.splitext(name)
    base = unicodedata.normalize("NFKD", base).encode("ascii", "ignore").decode("ascii")
    base = re.sub(r"[^A-Za-z0-9_-]+", "_", base).strip("_")
    ext = re.sub(r"[^A-Za-z0-9.]+", "", ext).lower()
    return f"{base}{ext}" if base else f"fichier{ext}"


@st.cache_resource
def _get_supabase_client():
    try:
        url = st.secrets["supabase"]["url"]
        key = st.secrets["supabase"]["key"]
    except Exception:
        st.error(
            "Configuration Supabase manquante. Ajoute [supabase] url / key dans "
            "les Secrets de l'app (Settings → Secrets)."
        )
        st.stop()
    return create_client(url, key)


def _list_stored_files() -> list:
    """Liste les fichiers .xlsx présents dans le bucket Supabase (persistants)."""
    client = _get_supabase_client()
    try:
        objects = client.storage.from_(SUPABASE_BUCKET).list()
    except Exception as e:
        st.error(f"Impossible de lister les fichiers Supabase : {e}")
        return []
    names = [o.get("name", "") for o in objects]
    return sorted(n for n in names if n.lower().endswith(".xlsx") and not n.startswith("~$"))


@st.cache_data(show_spinner=False)
def _download_file_bytes(name: str) -> bytes:
    """Télécharge le contenu binaire d'un fichier depuis le bucket Supabase (mis en cache)."""
    client = _get_supabase_client()
    return client.storage.from_(SUPABASE_BUCKET).download(name)


@st.cache_data(show_spinner=True)
def load_all_data(file_names: list):
    """file_names : liste de noms de fichiers stockés dans le bucket Supabase."""
    file_names = [n for n in file_names if not n.startswith("~$")]

    raw_logs, summaries, errors = [], [], []
    dt_legend = pd.DataFrame()

    for name in file_names:
        f = io.BytesIO(_download_file_bytes(name))
        f.name = name

        r = pd.DataFrame()
        try:
            r = read_raw_log_from_file(f)
            if not r.empty:
                raw_logs.append(r)
        except Exception as e:
            errors.append(f"{name} (journal) : {e}")
        try:
            s = read_summary_from_file(f)
            if not s.empty:
                # Si le résumé n'a pas de colonne Date (fichier sans cette info), on
                # déduit la date depuis le journal du même fichier (1 fichier = 1 jour).
                if s["Date"].isna().all() and not r.empty and "Date" in r.columns and not r["Date"].isna().all():
                    inferred_date = r["Date"].mode().iloc[0]
                    s["Date"] = inferred_date
                summaries.append(s)
        except Exception as e:
            errors.append(f"{name} (résumé) : {e}")
        if dt_legend.empty:
            dt_legend = read_dt_legend_from_file(f)

    raw_log = pd.concat(raw_logs, ignore_index=True) if raw_logs else pd.DataFrame()
    summary = pd.concat(summaries, ignore_index=True) if summaries else pd.DataFrame()
    return raw_log, summary, dt_legend, file_names, errors


# ----------------------------------------------------------------------------
# Mesures OEE (inchangé)
# ----------------------------------------------------------------------------
def compute_kpis(raw_log: pd.DataFrame, summary: pd.DataFrame) -> dict:
    temps_fonctionnement = raw_log["run [h]"].sum() if not raw_log.empty else 0
    temps_arret = raw_log["confessed [h]"].sum() if not raw_log.empty else 0
    denom_dispo = temps_fonctionnement + temps_arret
    availability = (temps_fonctionnement / denom_dispo) if denom_dispo > 0 else None

    if not summary.empty and summary["Yield"].sum() > 0 and "Efficiency" in summary.columns:
        performance = (summary["Yield"] * summary["Efficiency"]).sum() / summary["Yield"].sum()
    else:
        performance = None

    total_yield = summary["Yield"].sum() if not summary.empty else 0
    total_scrap = summary["Scrap"].sum() if not summary.empty else 0
    denom_qual = total_yield + total_scrap
    quality = (total_yield / denom_qual) if denom_qual > 0 else None

    oee = (availability * performance * quality) if None not in (availability, performance, quality) else None

    return {
        "Availability": availability, "Performance": performance,
        "Quality": quality, "OEE": oee,
        "Temps de Fonctionnement (h)": temps_fonctionnement,
        "Temps Arret Confesse (h)": temps_arret,
        "Total Yield": total_yield, "Total Scrap": total_scrap,
    }


def fmt_pct(x):
    return f"{x * 100:.1f} %" if x is not None else "N/A"


def _hex_to_rgba(hex_color: str, alpha: float = 0.35) -> str:
    """Convertit '#RRGGBB' en chaîne 'rgba(r,g,b,a)' compatible avec Plotly
    (Plotly n'accepte pas le format hexadécimal à 8 chiffres '#RRGGBBAA')."""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"


def status_color(value, good=0.85, mid=0.60):
    """Vert / orange / rouge selon des seuils standards OEE."""
    if value is None:
        return TEXT_MUTED
    if value >= good:
        return GREEN
    if value >= mid:
        return ORANGE
    return RED


def status_label(value, good=0.85, mid=0.60):
    if value is None:
        return "N/A"
    if value >= good:
        return "OK"
    if value >= mid:
        return "Attention"
    return "Critique"


# ----------------------------------------------------------------------------
# Style visuel commun pour tous les graphiques Plotly (thème sombre)
# ----------------------------------------------------------------------------
def style_fig(fig, height=None, show_legend=None):
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#e7ebff", size=12),
        legend=dict(font=dict(color="#c8cfe8", size=10)),
        margin=dict(t=36, b=10, l=10, r=10),
    )
    if height:
        fig.update_layout(height=height)
    if show_legend is not None:
        fig.update_layout(showlegend=show_legend)
    fig.update_xaxes(gridcolor=TRACK, zerolinecolor=TRACK)
    fig.update_yaxes(gridcolor=TRACK, zerolinecolor=TRACK)
    return fig


def kpi_hero_card(title, value_str, subtitle="", badge_text=None, badge_color=GREEN):
    badge_html = (
        f'<div class="kpi-badge" style="background:{badge_color}22;border:1px solid {badge_color};color:{badge_color};">{badge_text}</div>'
        if badge_text else ""
    )
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">{title}</div>
            <div class="kpi-row">
                <div class="kpi-value">{value_str}</div>
                {badge_html}
            </div>
            <div class="kpi-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def donut_metric(value, title, color, center_label=""):
    """Anneau complet façon 'CSAT' : valeur au centre, reste du cercle en gris ardoise."""
    v = 0.0 if value is None else max(0.0, min(1.0, value)) * 100
    remainder = 100 - v
    fig = go.Figure(
        go.Pie(
            values=[v, remainder], hole=0.66, rotation=90, direction="clockwise", sort=False,
            marker=dict(colors=[color, TRACK], line=dict(color=CARD_BG, width=2)),
            textinfo="none", hoverinfo="skip", showlegend=False,
        )
    )
    fig.update_layout(
        height=190,
        title=dict(text=title, x=0.5, y=0.97, font=dict(size=12, color="#c8cfe8")),
        annotations=[
            dict(text=f"<b>{v:.1f}%</b>", x=0.5, y=0.53, showarrow=False, font=dict(size=20, color="#ffffff")),
            dict(text=center_label, x=0.5, y=0.40, showarrow=False, font=dict(size=10, color=TEXT_MUTED)),
        ],
    )
    return style_fig(fig)


def gauge(value, title, color):
    """Jauge circulaire style 'compteur' pour un KPI en %."""
    v = value * 100 if value is not None else 0
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=v,
        number={"suffix": " %", "font": {"size": 22, "color": "#ffffff"}},
        title={"text": title, "font": {"size": 13, "color": "#c8cfe8"}},
        gauge={
            "axis": {"range": [0, 100], "tickwidth": 1, "tickcolor": TEXT_MUTED_2, "tickfont": {"color": TEXT_MUTED_2, "size": 9}},
            "bar": {"color": color, "thickness": 0.3},
            "bgcolor": "rgba(0,0,0,0)",
            "borderwidth": 0,
            "steps": [
                {"range": [0, 60], "color": "#3a1f2a"},
                {"range": [60, 85], "color": "#3a3320"},
                {"range": [85, 100], "color": "#1f3a28"},
            ],
        },
    ))
    fig.update_layout(height=190, margin=dict(t=34, b=6, l=18, r=18))
    return style_fig(fig)


def render_health_block(raw_log_f: pd.DataFrame):
    """Répartition des machines par niveau de disponibilité (façon Promoteurs/Passifs/Détracteurs)."""
    st.markdown(
        f"""<div class="section-title" style="margin-bottom:4px;">Santé des machines</div>""",
        unsafe_allow_html=True,
    )
    if raw_log_f.empty:
        st.markdown(
            f'<div class="health-card"><span style="color:{TEXT_MUTED};font-size:13px;">Pas de données pour la sélection.</span></div>',
            unsafe_allow_html=True,
        )
        return

    g = raw_log_f.groupby("Machine").agg(run=("run [h]", "sum"), stop=("confessed [h]", "sum"))
    denom = g["run"] + g["stop"]
    g = g[denom > 0].copy()
    g["avail"] = g["run"] / (g["run"] + g["stop"])

    total = len(g)
    if total == 0:
        st.markdown(
            f'<div class="health-card"><span style="color:{TEXT_MUTED};font-size:13px;">Pas de données pour la sélection.</span></div>',
            unsafe_allow_html=True,
        )
        return

    good = int((g["avail"] >= 0.85).sum())
    mid = int(((g["avail"] >= 0.60) & (g["avail"] < 0.85)).sum())
    low = int((g["avail"] < 0.60).sum())

    rows = [
        ("🟢", GREEN, good, "Performantes (≥ 85%)"),
        ("🟡", ORANGE, mid, "Moyennes (60-85%)"),
        ("🔴", RED, low, "Critiques (< 60%)"),
    ]
    html_rows = ""
    for icon, color, count, label in rows:
        pct = (count / total * 100) if total else 0
        # IMPORTANT : tout sur une seule ligne, sans indentation. Si on indente
        # ces lignes avec 4+ espaces, Markdown les interprète comme un bloc de
        # code (<pre>) au lieu de les rendre comme du HTML.
        html_rows += (
            f'<div class="health-row">'
            f'<div class="health-dot" style="background:{color}22;">{icon}</div>'
            f'<div>'
            f'<div class="health-value">{pct:.1f}<span style="font-size:13px;">%</span></div>'
            f'<div class="health-label">{label} · {count}/{total} machines</div>'
            f'</div>'
            f'</div>'
        )
    st.markdown(f'<div class="health-card">{html_rows}</div>', unsafe_allow_html=True)


def daily_trend(raw_log_f: pd.DataFrame, summary_f: pd.DataFrame) -> pd.DataFrame:
    """Calcule Availability / Performance / Quality / OEE jour par jour pour les graphiques de tendance."""
    dates = set()
    if not raw_log_f.empty:
        dates |= set(raw_log_f["Date"])
    if not summary_f.empty:
        dates |= set(summary_f["Date"].dropna())
    dates = sorted(d for d in dates if pd.notna(d))
    rows = []
    for d in dates:
        rl = raw_log_f[raw_log_f["Date"] == d] if not raw_log_f.empty else pd.DataFrame()
        sm = summary_f[summary_f["Date"] == d] if not summary_f.empty else pd.DataFrame()
        k = compute_kpis(rl, sm)
        rows.append({
            "Date": d, "Availability": k["Availability"], "Performance": k["Performance"],
            "Quality": k["Quality"], "OEE": k["OEE"],
        })
    return pd.DataFrame(rows)


def _get_admins() -> dict:
    """Récupère les comptes admin (nom -> mot de passe) depuis les secrets Streamlit."""
    try:
        return dict(st.secrets["admins"])
    except Exception:
        return {}


# ----------------------------------------------------------------------------
# Interface
# ----------------------------------------------------------------------------
st.markdown('<div class="dash-title">📊 Dashboard OEE — Département Production</div>', unsafe_allow_html=True)

if "is_admin" not in st.session_state:
    st.session_state.is_admin = False
    st.session_state.admin_user = None

with st.sidebar:
    st.header("🔐 Espace administrateur")

    if not st.session_state.is_admin:
        with st.form("login_form"):
            username = st.text_input("Nom d'utilisateur")
            password = st.text_input("Mot de passe", type="password")
            submitted = st.form_submit_button("Se connecter")
        if submitted:
            admins = _get_admins()
            if username in admins and password == admins[username]:
                st.session_state.is_admin = True
                st.session_state.admin_user = username
                st.rerun()
            else:
                st.error("Identifiants incorrects.")
    else:
        st.success(f"Connecté : {st.session_state.admin_user}")
        if st.button("Se déconnecter"):
            st.session_state.is_admin = False
            st.session_state.admin_user = None
            st.rerun()

        st.divider()
        st.subheader("📤 Ajouter des fichiers")
        new_files = st.file_uploader(
            "Fichiers Excel journaliers (.xlsx)",
            type=["xlsx"], accept_multiple_files=True, key="admin_uploader",
        )
        if new_files and st.button("Enregistrer dans le stock", use_container_width=True):
            client = _get_supabase_client()
            try:
                for f in new_files:
                    safe_name = _sanitize_filename(f.name)
                    client.storage.from_(SUPABASE_BUCKET).upload(
                        path=safe_name,
                        file=f.getvalue(),
                        file_options={"upsert": "true"},
                    )
                st.cache_data.clear()
                st.success(f"{len(new_files)} fichier(s) ajouté(s) de façon permanente.")
                st.rerun()
            except Exception as e:
                st.error(f"Échec de l'envoi vers Supabase : {e}")

        st.divider()
        st.subheader("🗑️ Fichiers stockés")
        stored = _list_stored_files()
        if stored:
            to_delete = st.selectbox("Supprimer un fichier", ["-"] + stored)
            if to_delete != "-" and st.button("Confirmer la suppression"):
                try:
                    _get_supabase_client().storage.from_(SUPABASE_BUCKET).remove([to_delete])
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"Échec de la suppression : {e}")
        else:
            st.caption("Aucun fichier stocké pour le moment.")

file_names_on_disk = _list_stored_files()

if not file_names_on_disk:
    st.info("Aucune donnée disponible pour le moment. Un administrateur doit se connecter pour ajouter des fichiers.")
    st.stop()

raw_log, summary, dt_legend, files, errors = load_all_data(file_names_on_disk)

if errors:
    with st.expander("⚠️ Fichiers non lus correctement"):
        for e in errors:
            st.write("-", e)

if raw_log.empty and summary.empty:
    st.warning("Aucune donnée exploitable trouvée pour le moment.")
    st.stop()

# ---- Filtres ----
with st.sidebar:
    st.header("🔎 Filtres")
    all_dates = sorted(set(raw_log.get("Date", pd.Series(dtype=object))) |
                       set(summary.get("Date", pd.Series(dtype=object))))
    if not all_dates:
        st.warning("Aucune date trouvée.")
        st.stop()
    selected_dates = st.multiselect("Date(s)", options=all_dates, default=all_dates)
    selected_shifts = st.multiselect("Shift(s)", options=["A", "B", "C"], default=["A", "B", "C"])
    all_cc = sorted(summary["CC"].dropna().unique()) if "CC" in summary.columns and not summary.empty else []
    selected_cc = st.multiselect("CC (cellule)", options=all_cc, default=all_cc)

summary_f = summary[
    summary["Date"].isin(selected_dates) & (summary["CC"].isin(selected_cc) if selected_cc else True)
].copy() if not summary.empty else pd.DataFrame()

machines_du_cc = set(summary_f["Machine"].unique()) if (selected_cc and not summary_f.empty) else (
    set(summary["Machine"].unique()) if not summary.empty else set()
)

raw_log_f = raw_log[
    raw_log["Date"].isin(selected_dates) &
    raw_log["Shift"].isin(selected_shifts) &
    (raw_log["Machine"].isin(machines_du_cc) if machines_du_cc else True)
].copy() if not raw_log.empty else pd.DataFrame()

if raw_log_f.empty and summary_f.empty:
    st.warning("Aucune donnée ne correspond aux filtres sélectionnés.")
    st.stop()

kpis = compute_kpis(raw_log_f, summary_f)

date_range_txt = (
    f"{min(selected_dates)} → {max(selected_dates)}" if len(selected_dates) > 1
    else f"{selected_dates[0]}" if selected_dates else "—"
)
st.markdown(
    f'<div class="dash-subtitle">Période : {date_range_txt} · Shift(s) : {", ".join(selected_shifts) or "—"} '
    f'· CC : {", ".join(map(str, selected_cc)) if selected_cc else "Tous"}</div>',
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------
# Rangée 1 : cartes KPI (Availability / Performance / Quality / OEE)
# ----------------------------------------------------------------------------
k1, k2, k3, k4 = st.columns(4)
with k1:
    v = kpis["Availability"]
    kpi_hero_card("Availability", fmt_pct(v),
                   f"Fonctionnement : {kpis['Temps de Fonctionnement (h)']:.1f} h",
                   status_label(v), status_color(v))
with k2:
    v = kpis["Performance"]
    kpi_hero_card("Performance", fmt_pct(v),
                   "Rendement vs cadence cible",
                   status_label(v), status_color(v))
with k3:
    v = kpis["Quality"]
    kpi_hero_card("Quality", fmt_pct(v),
                   f"Yield : {kpis['Total Yield']:.0f} · Scrap : {kpis['Total Scrap']:.0f}",
                   status_label(v), status_color(v))
with k4:
    v = kpis["OEE"]
    kpi_hero_card("OEE global", fmt_pct(v),
                   "Objectif usine : 85 %",
                   status_label(v), status_color(v))

st.write("")

# ----------------------------------------------------------------------------
# Rangée 2 : jauges circulaires (Availability / Performance / Quality) + santé machines
# ----------------------------------------------------------------------------
g1, g2, g3, g4 = st.columns([1, 1, 1, 1.25])
g1.plotly_chart(donut_metric(kpis["OEE"], "OEE", ORANGE, "Objectif 85%"), use_container_width=True)
g2.plotly_chart(gauge(kpis["Availability"], "Availability", APTIV_BLUE), use_container_width=True)
g3.plotly_chart(gauge(kpis["Quality"], "Quality", GREEN), use_container_width=True)
with g4:
    render_health_block(raw_log_f)

st.write("")

# ----------------------------------------------------------------------------
# Rangée 3 : répartition par catégorie d'arrêt + tendances OEE / Availability
# ----------------------------------------------------------------------------
c1, c2, c3 = st.columns(3)

with c1:
    st.markdown('<div class="section-title">Répartition par catégorie d\'arrêt</div>', unsafe_allow_html=True)
    if not raw_log_f.empty and not dt_legend.empty and "CodeDT_num" in dt_legend.columns:
        raw_with_cat = raw_log_f.merge(
            dt_legend[["CodeDT_num", "Categorie", "ColorHEX"]],
            left_on="Downtime reason", right_on="CodeDT_num", how="left",
        )
        df_cat = (
            raw_with_cat[raw_with_cat["Downtime reason"] != 0]
            .groupby("Categorie", dropna=False)["confessed [h]"].sum()
            .reset_index().sort_values("confessed [h]", ascending=False)
        )
        df_cat["Categorie"] = df_cat["Categorie"].fillna("Non catégorisé")
        if not df_cat.empty and df_cat["confessed [h]"].sum() > 0:
            cat_color_map = (
                raw_with_cat.dropna(subset=["Categorie"]).drop_duplicates(subset=["Categorie"])
                .set_index("Categorie")["ColorHEX"].to_dict()
            )
            colors = [cat_color_map.get(cat, "#9E9E9E") for cat in df_cat["Categorie"]]
            fig_cat = go.Figure(go.Pie(
                labels=df_cat["Categorie"], values=df_cat["confessed [h]"], hole=0.5,
                marker=dict(colors=colors, line=dict(color=CARD_BG, width=2)),
                textinfo="percent", hovertemplate="<b>%{label}</b><br>%{value:.2f} h<extra></extra>",
            ))
            st.plotly_chart(style_fig(fig_cat, height=260), use_container_width=True)
        else:
            st.info("Aucun arrêt enregistré pour la sélection actuelle.")
    else:
        st.info("Légende DT_LEGEND indisponible pour catégoriser les arrêts.")

trend = daily_trend(raw_log_f, summary_f)

with c2:
    st.markdown('<div class="section-title">Évolution de l\'OEE</div>', unsafe_allow_html=True)
    if not trend.empty:
        fig_line = go.Figure()
        fig_line.add_trace(go.Scatter(
            x=trend["Date"], y=trend["OEE"] * 100, mode="lines+markers",
            line=dict(color=ORANGE, width=3), marker=dict(size=6),
            name="OEE", hovertemplate="%{x}<br>OEE : %{y:.1f}%<extra></extra>",
        ))
        fig_line.add_hline(y=85, line_dash="dot", line_color=TEXT_MUTED_2, annotation_text="Objectif 85%",
                            annotation_font_color=TEXT_MUTED_2, annotation_font_size=10)
        fig_line.update_layout(yaxis=dict(range=[0, 105], ticksuffix="%"), showlegend=False)
        st.plotly_chart(style_fig(fig_line, height=260), use_container_width=True)
    else:
        st.info("Pas assez de données pour tracer une tendance.")

with c3:
    st.markdown('<div class="section-title">Évolution de l\'Availability</div>', unsafe_allow_html=True)
    if not trend.empty:
        fig_area = go.Figure()
        fig_area.add_trace(go.Scatter(
            x=trend["Date"], y=trend["Availability"] * 100, mode="lines", fill="tozeroy",
            line=dict(color=APTIV_BLUE, width=2.5), fillcolor=_hex_to_rgba(APTIV_BLUE, 0.33),
            name="Availability", hovertemplate="%{x}<br>Availability : %{y:.1f}%<extra></extra>",
        ))
        fig_area.update_layout(yaxis=dict(range=[0, 105], ticksuffix="%"), showlegend=False)
        st.plotly_chart(style_fig(fig_area, height=260), use_container_width=True)
    else:
        st.info("Pas assez de données pour tracer une tendance.")

st.write("")

# ----------------------------------------------------------------------------
# Rangée 4 : répartition des arrêts par machine (stacked bar)
# ----------------------------------------------------------------------------
st.markdown('<div class="section-title">Code d\'arrêt (catégorie) par machine</div>', unsafe_allow_html=True)
if not raw_log_f.empty and not dt_legend.empty and "CodeDT_num" in dt_legend.columns:
    raw_with_cat_m = raw_log_f.merge(
        dt_legend[["CodeDT_num", "Categorie", "ColorHEX"]],
        left_on="Downtime reason", right_on="CodeDT_num", how="left",
    )
    raw_with_cat_m["Categorie"] = raw_with_cat_m["Categorie"].fillna("Non catégorisé")
    only_dt = raw_with_cat_m[raw_with_cat_m["Downtime reason"] != 0]

    df_mc = only_dt.groupby(["Machine", "Categorie"])["confessed [h]"].sum().reset_index()
    if not df_mc.empty and df_mc["confessed [h]"].sum() > 0:
        cat_color_map_m = (
            raw_with_cat_m.dropna(subset=["Categorie"]).drop_duplicates(subset=["Categorie"])
            .set_index("Categorie")["ColorHEX"].to_dict()
        )
        machine_order = df_mc.groupby("Machine")["confessed [h]"].sum().sort_values(ascending=False).index.tolist()
        fig_m = px.bar(
            df_mc, x="Machine", y="confessed [h]", color="Categorie",
            category_orders={"Machine": machine_order},
            color_discrete_map=cat_color_map_m,
            custom_data=["Categorie"],
        )
        fig_m.update_layout(barmode="stack", xaxis_title="", yaxis_title="Heures d'arrêt")
        event_m = st.plotly_chart(
            style_fig(fig_m, height=280), use_container_width=True,
            on_select="rerun", selection_mode="points", key="machine_cat_chart",
        )

        points = (event_m or {}).get("selection", {}).get("points", [])
        if points:
            p = points[0]
            machine_sel = p.get("x")
            cat_sel = p["customdata"][0] if p.get("customdata") else None
            detail = only_dt[
                (only_dt["Machine"] == machine_sel) & (only_dt["Categorie"] == cat_sel)
            ][["Machine", "Downtime name", "Commentaire", "confessed [h]"]]
            st.markdown(f"**Détail — Machine `{machine_sel}` / catégorie `{cat_sel}`**")
            if not detail.empty:
                st.dataframe(detail.sort_values("confessed [h]", ascending=False),
                             use_container_width=True, hide_index=True)
            else:
                st.caption("Aucun détail trouvé pour cette sélection.")
        else:
            st.caption("👆 Clique sur une barre du graphique pour afficher le détail des arrêts correspondants.")
    else:
        st.info("Aucun arrêt enregistré pour la sélection actuelle.")
else:
    st.info("Légende DT_LEGEND indisponible pour catégoriser les arrêts.")

# ----------------------------------------------------------------------------
# Analyse détaillée (dépliable) : Pareto des causes + cycle time hors tolérance
# ----------------------------------------------------------------------------
with st.expander("🔍 Analyse détaillée des arrêts et du cycle time", expanded=False):
    col_pie, col_bar = st.columns(2)

    # NOTE IMPORTANTE :
    # - Dans "DATA 1", "Downtime reason" contient le CODE (0, 1, 4, ... 31) et
    #   "Downtime name" contient le TEXTE de la cause.
    # - Dans "DT_LEGEND", c'est l'inverse : "CodeDT" (colonne "Downtime name" du fichier)
    #   contient le code, et "Libelle" (colonne "Downtime reason" du fichier) contient le texte.
    # On fusionne donc sur le CODE numérique normalisé (CodeDT_num vs Downtime reason).

    with col_pie:
        st.markdown("**Répartition par cause**")
        if not raw_log_f.empty:
            df_dt = (
                raw_log_f[raw_log_f["Downtime reason"] != 0]
                .groupby("Downtime reason")
                .agg(**{"confessed [h]": ("confessed [h]", "sum"), "Downtime name": ("Downtime name", "first")})
                .reset_index().sort_values("confessed [h]", ascending=False)
            )
            if not df_dt.empty and df_dt["confessed [h]"].sum() > 0:
                if not dt_legend.empty and "CodeDT_num" in dt_legend.columns:
                    df_dt = df_dt.merge(
                        dt_legend[["CodeDT_num", "Libelle", "ColorHEX"]],
                        left_on="Downtime reason", right_on="CodeDT_num", how="left",
                    )
                    df_dt["Label"] = df_dt["Libelle"].fillna(df_dt["Downtime name"])
                    colors = ["#9E9E9E" if pd.isna(c) else c for c in df_dt["ColorHEX"]]
                else:
                    df_dt["Label"] = df_dt["Downtime name"]
                    colors = None
                fig = go.Figure(go.Pie(
                    labels=df_dt["Label"], values=df_dt["confessed [h]"], hole=0.45,
                    marker=dict(colors=colors, line=dict(color=CARD_BG, width=2)) if colors else {},
                    textinfo="percent", hovertemplate="<b>%{label}</b><br>%{value:.2f} h<extra></extra>",
                ))
                st.plotly_chart(style_fig(fig, height=280), use_container_width=True)
            else:
                st.info("Aucun arrêt enregistré pour la sélection actuelle.")
        else:
            st.info("Pas de données disponibles.")

    with col_bar:
        st.markdown("**Top 10 des causes (heures)**")
        if not raw_log_f.empty:
            df_c = (
                raw_log_f[raw_log_f["Downtime reason"] != 0]
                .groupby("Downtime reason")
                .agg(**{"confessed [h]": ("confessed [h]", "sum"), "Downtime name": ("Downtime name", "first")})
                .reset_index().sort_values("confessed [h]", ascending=True).tail(10)
            )
            if not df_c.empty and df_c["confessed [h]"].sum() > 0:
                if not dt_legend.empty and "CodeDT_num" in dt_legend.columns:
                    df_c = df_c.merge(
                        dt_legend[["CodeDT_num", "Libelle", "ColorHEX"]],
                        left_on="Downtime reason", right_on="CodeDT_num", how="left",
                    )
                    df_c["Label"] = df_c["Libelle"].fillna(df_c["Downtime name"])
                    colors_map = {row["Label"]: (row["ColorHEX"] if pd.notna(row["ColorHEX"]) else "#9E9E9E")
                                  for _, row in df_c.iterrows()}
                else:
                    df_c["Label"] = df_c["Downtime name"]
                    colors_map = None
                fig = px.bar(
                    df_c, x="confessed [h]", y="Label", orientation="h",
                    color="Label" if colors_map else None,
                    color_discrete_map=colors_map if colors_map else None,
                    color_discrete_sequence=[RED] if not colors_map else None,
                )
                fig.update_layout(showlegend=False, xaxis_title="Heures d'arrêt", yaxis_title="")
                fig.update_traces(texttemplate="%{x:.2f} h", textposition="outside")
                st.plotly_chart(style_fig(fig, height=280), use_container_width=True)
            else:
                st.info("Aucun arrêt enregistré pour la sélection actuelle.")
        else:
            st.info("Pas de données disponibles.")

    st.markdown("---")
    st.markdown("**⏱️ Cycle time hors tolérance**")
    if not summary_f.empty and "Average cycle (s)" in summary_f.columns:
        cyc = summary_f.dropna(subset=["Average cycle (s)", "Target cycle (s)"]).copy()
        cyc["Tolerance cycle (s)"] = cyc["Tolerance cycle (s)"].fillna(0)
        cyc["Borne sup (s)"] = cyc["Target cycle (s)"] + cyc["Tolerance cycle (s)"]
        cyc["Dépassement (s)"] = cyc["Average cycle (s)"] - cyc["Borne sup (s)"]
        cyc_over = cyc[cyc["Dépassement (s)"] > 0].sort_values("Dépassement (s)", ascending=True)

        if not cyc_over.empty:
            fig = px.bar(
                cyc_over, x="Dépassement (s)", y="Machine", orientation="h",
                color_discrete_sequence=[RED],
                custom_data=["Article name", "Target cycle (s)", "Tolerance cycle (s)", "Average cycle (s)", "Dépassement (s)"],
            )
            fig.update_traces(
                hovertemplate=(
                    "<b>%{y}</b><br>Article : %{customdata[0]}<br>"
                    "Cible : %{customdata[1]:.2f} s ± %{customdata[2]:.2f} s<br>"
                    "Moyenne mesurée : %{customdata[3]:.2f} s<br>"
                    "Dépassement : %{x:.2f} s<extra></extra>"
                ),
                texttemplate="+%{x:.2f} s", textposition="outside",
            )
            fig.update_layout(xaxis_title="Dépassement au-delà de la tolérance (s)", yaxis_title="")

            cyc_chart_col, cyc_detail_col = st.columns([2, 1])
            with cyc_chart_col:
                event_cyc = st.plotly_chart(
                    style_fig(fig, height=max(220, 26 * len(cyc_over))), use_container_width=True,
                    on_select="rerun", selection_mode="points", key="cycle_time_chart",
                )
            with cyc_detail_col:
                points_cyc = (event_cyc or {}).get("selection", {}).get("points", [])
                if points_cyc:
                    p = points_cyc[0]
                    machine_sel_cyc = p.get("y")
                    row_sel = cyc_over[cyc_over["Machine"] == machine_sel_cyc]
                    if not row_sel.empty:
                        row_sel = row_sel.iloc[0]
                        st.markdown(f"**Machine `{machine_sel_cyc}`**")
                        st.write(f"Article : {row_sel.get('Article name', 'N/A')}")
                        st.write(f"Cible : {row_sel['Target cycle (s)']:.2f} s ± {row_sel['Tolerance cycle (s)']:.2f} s")
                        st.write(f"Moyenne mesurée : {row_sel['Average cycle (s)']:.2f} s")
                        st.write(f"Dépassement : +{row_sel['Dépassement (s)']:.2f} s")
                    else:
                        st.caption("Aucun détail trouvé pour cette sélection.")
                else:
                    st.caption("👆 Clique sur une barre pour voir le détail ici.")
        else:
            st.info("Aucune machine ne dépasse son intervalle de cycle time cible pour la sélection actuelle.")
    else:
        st.info("Colonnes de temps de cycle non disponibles dans les fichiers chargés.")

with st.expander("📄 Voir les données détaillées filtrées"):
    if not raw_log_f.empty:
        st.write("**RAW_LOG**")
        st.dataframe(raw_log_f, use_container_width=True)
    if not summary_f.empty:
        st.write("**SUMMARY**")
        st.dataframe(summary_f, use_container_width=True)
    if not dt_legend.empty:
        st.write("**DT_LEGEND**")
        st.dataframe(dt_legend, use_container_width=True)
