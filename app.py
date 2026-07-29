"""
Dashboard OEE - Aptiv (version design amélioré)
=================================================
Lit automatiquement tous les fichiers Excel journaliers déposés dans un dossier
(chacun avec les onglets "DATA 1"/"DATA 2"/"DT" -- OU "SUMMARY"/"DT_LEGEND",
les deux noms sont acceptés), calcule Availability, Performance, Quality et OEE,
et affiche un dashboard interactif filtrable par Date, Shift (A/B/C) et CC.

Lancer avec :  streamlit run app.py
"""

import glob
import os
import re

import pandas as pd  # type: ignore
import plotly.express as px  # type: ignore
import plotly.graph_objects as go  # type: ignore
import streamlit as st  # type: ignore

# ----------------------------------------------------------------------------
# Configuration générale de la page
# ----------------------------------------------------------------------------
st.set_page_config(page_title="Dashboard OEE - Aptiv", layout="wide", page_icon="📊")

APTIV_BLUE = "#0033A0"
APTIV_LIGHT = "#4A90D9"
GREEN = "#2E7D32"
ORANGE = "#F5A623"
RED = "#D32F2F"

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.5rem;}
    div[data-testid="stMetric"] {
        background-color: #F5F7FA;
        border: 1px solid #E0E4EA;
        border-radius: 10px;
        padding: 12px 16px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ----------------------------------------------------------------------------
# Fonctions de lecture / nettoyage des données
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
    """Nom du fichier, qu'il s'agisse d'un chemin (str) ou d'un fichier uploadé (UploadedFile)."""
    return getattr(f, "name", None) or os.path.basename(str(f))


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

    col_machine = _find_col(df, "machine")
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

    out = pd.DataFrame()
    out["Machine"] = df[col_machine]
    out["Yield"] = pd.to_numeric(df[col_yield], errors="coerce").fillna(0)
    out["Scrap"] = pd.to_numeric(df[col_scrap], errors="coerce").fillna(0) if col_scrap else 0
    out["Efficiency"] = pd.to_numeric(df[col_efficiency], errors="coerce") if col_efficiency else float("nan")
    out["CC"] = df[col_cc] if col_cc else "N/A"
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


@st.cache_data(show_spinner=True)
def load_all_data(files: list):
    """files : liste de chemins (mode dossier local) OU de fichiers uploadés (st.file_uploader)."""
    files = [f for f in files if not _name_of(f).startswith("~$")]

    raw_logs, summaries, errors = [], [], []
    dt_legend = pd.DataFrame()

    for f in files:
        r = pd.DataFrame()
        try:
            r = read_raw_log_from_file(f)
            if not r.empty:
                raw_logs.append(r)
        except Exception as e:
            errors.append(f"{_name_of(f)} (journal) : {e}")
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
            errors.append(f"{_name_of(f)} (résumé) : {e}")
        if dt_legend.empty:
            dt_legend = read_dt_legend_from_file(f)

    raw_log = pd.concat(raw_logs, ignore_index=True) if raw_logs else pd.DataFrame()
    summary = pd.concat(summaries, ignore_index=True) if summaries else pd.DataFrame()
    return raw_log, summary, dt_legend, files, errors


# ----------------------------------------------------------------------------
# Mesures OEE
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


def gauge(value, title, color):
    """Jauge circulaire style 'compteur' pour un KPI en %."""
    v = value * 100 if value is not None else 0
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=v,
        number={"suffix": " %", "font": {"size": 30}},
        title={"text": title, "font": {"size": 16}},
        gauge={
            "axis": {"range": [0, 100], "tickwidth": 1},
            "bar": {"color": color, "thickness": 0.3},
            "bgcolor": "white",
            "steps": [
                {"range": [0, 60], "color": "#FBE4E4"},
                {"range": [60, 85], "color": "#FEF3D9"},
                {"range": [85, 100], "color": "#E4F3E5"},
            ],
        },
    ))
    fig.update_layout(height=220, margin=dict(t=40, b=10, l=20, r=20))
    return fig


DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)


def _get_admins() -> dict:
    """Récupère les comptes admin (nom -> mot de passe) depuis les secrets Streamlit."""
    try:
        return dict(st.secrets["admins"])
    except Exception:
        return {}


def _list_stored_files() -> list:
    files = sorted(glob.glob(os.path.join(DATA_DIR, "*.xlsx")))
    return [f for f in files if not os.path.basename(f).startswith("~$")]


# ----------------------------------------------------------------------------
# Interface
# ----------------------------------------------------------------------------
st.title("📊 Dashboard OEE — Département Production")

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
            for f in new_files:
                with open(os.path.join(DATA_DIR, f.name), "wb") as out:
                    out.write(f.getbuffer())
            st.cache_data.clear()
            st.success(f"{len(new_files)} fichier(s) ajouté(s).")
            st.rerun()

        st.divider()
        st.subheader("🗑️ Fichiers stockés")
        stored = [os.path.basename(f) for f in _list_stored_files()]
        if stored:
            to_delete = st.selectbox("Supprimer un fichier", ["-"] + stored)
            if to_delete != "-" and st.button(f"Confirmer la suppression"):
                os.remove(os.path.join(DATA_DIR, to_delete))
                st.cache_data.clear()
                st.rerun()
        else:
            st.caption("Aucun fichier stocké pour le moment.")

files_on_disk = _list_stored_files()

if not files_on_disk:
    st.info("Aucune donnée disponible pour le moment. Un administrateur doit se connecter pour ajouter des fichiers.")
    st.stop()

raw_log, summary, dt_legend, files, errors = load_all_data(files_on_disk)

if errors:
    with st.expander("⚠️ Fichiers non lus correctement"):
        for e in errors:
            st.write("-", e)

if raw_log.empty and summary.empty:
    st.warning("Aucune donnée exploitable trouvée pour le moment.")
    st.stop()

st.caption(f"📁 {len(files)} fichier(s) chargé(s) : " + ", ".join(_name_of(f) for f in files))

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

# ----------------------------------------------------------------------------
# Rangée 1 : jauges des 4 KPI
# ----------------------------------------------------------------------------
g1, g2, g3, g4 = st.columns(4)
g1.plotly_chart(gauge(kpis["Availability"], "Availability", APTIV_BLUE), use_container_width=True)
g2.plotly_chart(gauge(kpis["Performance"], "Performance", APTIV_LIGHT), use_container_width=True)
g3.plotly_chart(gauge(kpis["Quality"], "Quality", GREEN), use_container_width=True)
g4.plotly_chart(gauge(kpis["OEE"], "OEE", ORANGE), use_container_width=True)

st.divider()

# ----------------------------------------------------------------------------
# Rangée 2 : OEE par jour / OEE par CC
# ----------------------------------------------------------------------------
c1, c2 = st.columns(2)

with c1:
    st.subheader("Évolution de l'OEE par jour")
    rows = []
    for d in sorted(selected_dates):
        rl_d = raw_log_f[raw_log_f["Date"] == d] if not raw_log_f.empty else pd.DataFrame()
        sm_d = summary_f[summary_f["Date"] == d] if not summary_f.empty else pd.DataFrame()
        k = compute_kpis(rl_d, sm_d)
        if k["OEE"] is not None:
            rows.append({"Date": d, "OEE %": k["OEE"] * 100})
    if rows:
        fig = px.line(pd.DataFrame(rows), x="Date", y="OEE %", markers=True,
                      color_discrete_sequence=[APTIV_BLUE])
        fig.update_yaxes(rangemode="tozero")
        fig.update_layout(height=350, margin=dict(t=20))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Pas assez de données pour ce graphique.")

with c2:
    st.subheader("OEE par CC")
    rows = []
    if not summary_f.empty and "CC" in summary_f.columns:
        for cc in sorted(summary_f["CC"].dropna().unique()):
            sm_cc = summary_f[summary_f["CC"] == cc]
            machines_cc = set(sm_cc["Machine"].unique())
            rl_cc = raw_log_f[raw_log_f["Machine"].isin(machines_cc)] if not raw_log_f.empty else pd.DataFrame()
            k = compute_kpis(rl_cc, sm_cc)
            if k["OEE"] is not None:
                rows.append({"CC": cc, "OEE %": k["OEE"] * 100})
    if rows:
        fig = px.bar(pd.DataFrame(rows).sort_values("OEE %", ascending=False), x="CC", y="OEE %",
                     color_discrete_sequence=[APTIV_LIGHT])
        fig.update_yaxes(rangemode="tozero")
        fig.update_layout(height=350, margin=dict(t=20))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Pas assez de données pour ce graphique.")

st.divider()

# ----------------------------------------------------------------------------
# Rangée 2bis : Machines dont le cycle time dépasse l'intervalle cible
# ----------------------------------------------------------------------------
st.subheader("⏱️ Cycle time hors tolérance")
st.caption(
    "Machines dont le temps de cycle moyen dépasse la borne supérieure de "
    "l'intervalle cible (Target cycle time ± tolérance). Les machines dans "
    "l'intervalle, ou plus rapides que la cible, ne sont pas affichées."
)
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
            custom_data=["Article name", "Target cycle (s)", "Tolerance cycle (s)", "Average cycle (s)"],
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
        fig.update_layout(height=max(320, 40 * len(cyc_over)), margin=dict(t=10),
                           xaxis_title="Dépassement au-delà de la tolérance (s)", yaxis_title="")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Aucune machine ne dépasse son intervalle de cycle time cible pour la sélection actuelle.")
else:
    st.info("Colonnes de temps de cycle non disponibles dans les fichiers chargés.")

st.divider()

# ----------------------------------------------------------------------------
# Rangée 3 : répartition des arrêts (donut par cause) + top causes (barres)
# ----------------------------------------------------------------------------
st.subheader("📉 Analyse des temps d'arrêt")

# --- Nouveau donut : répartition par CODE d'arrêt (catégorie : Prod/Mnt/MG/TR/MPC/Qlt/ME/HR...) ---
st.markdown("**Répartition par code d'arrêt (catégorie)**")
if not raw_log_f.empty and not dt_legend.empty and "CodeDT_num" in dt_legend.columns:
    raw_with_cat = raw_log_f.merge(
        dt_legend[["CodeDT_num", "Categorie", "ColorHEX"]],
        left_on="Downtime reason", right_on="CodeDT_num", how="left",
    )
    df_cat = (
        raw_with_cat[raw_with_cat["Downtime reason"] != 0]
        .groupby("Categorie", dropna=False)["confessed [h]"].sum()
        .reset_index()
        .sort_values("confessed [h]", ascending=False)
    )
    df_cat["Categorie"] = df_cat["Categorie"].fillna("Non catégorisé")
    if not df_cat.empty and df_cat["confessed [h]"].sum() > 0:
        # une couleur représentative par catégorie (même couleur pour tous les codes d'une catégorie)
        cat_color_map = (
            raw_with_cat.dropna(subset=["Categorie"])
            .drop_duplicates(subset=["Categorie"])
            .set_index("Categorie")["ColorHEX"].to_dict()
        )
        colors = [cat_color_map.get(c, "#9E9E9E") for c in df_cat["Categorie"]]
        fig_cat = go.Figure(go.Pie(
            labels=df_cat["Categorie"], values=df_cat["confessed [h]"], hole=0.45,
            marker=dict(colors=colors),
            textinfo="percent", hovertemplate="<b>%{label}</b><br>%{value:.2f} h<extra></extra>",
        ))
        fig_cat.update_layout(height=380, margin=dict(t=10, b=10), legend=dict(font=dict(size=10)))
        st.plotly_chart(fig_cat, use_container_width=True)
    else:
        st.info("Aucun arrêt enregistré pour la sélection actuelle.")
else:
    st.info("Légende DT_LEGEND indisponible pour catégoriser les arrêts.")

st.divider()
col_pie, col_bar = st.columns(2)

# NOTE IMPORTANTE :
# - Dans "DATA 1", "Downtime reason" contient le CODE (0, 1, 4, ... 31) et
#   "Downtime name" contient le TEXTE de la cause.
# - Dans "DT_LEGEND", c'est l'inverse : "CodeDT" (colonne "Downtime name" du fichier)
#   contient le code, et "Libelle" (colonne "Downtime reason" du fichier) contient le texte.
# On fusionne donc désormais sur le CODE numérique normalisé (CodeDT_num vs Downtime reason),
# et non plus sur le texte, pour que ColorHEX se rattache correctement à chaque cause.

with col_pie:
    st.markdown("**Répartition par cause**")
    if not raw_log_f.empty:
        df_dt = (
            raw_log_f[raw_log_f["Downtime reason"] != 0]
            .groupby("Downtime reason")
            .agg(**{
                "confessed [h]": ("confessed [h]", "sum"),
                "Downtime name": ("Downtime name", "first"),
            })
            .reset_index()
            .sort_values("confessed [h]", ascending=False)
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
                marker=dict(colors=colors) if colors else {},
                textinfo="percent", hovertemplate="<b>%{label}</b><br>%{value:.2f} h<extra></extra>",
            ))
            fig.update_layout(height=400, margin=dict(t=10, b=10),
                               legend=dict(font=dict(size=10)))
            st.plotly_chart(fig, use_container_width=True)
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
            .agg(**{
                "confessed [h]": ("confessed [h]", "sum"),
                "Downtime name": ("Downtime name", "first"),
            })
            .reset_index()
            .sort_values("confessed [h]", ascending=True)
            .tail(10)
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
            fig.update_layout(height=400, margin=dict(t=10, b=10), showlegend=False,
                              xaxis_title="Heures d'arrêt", yaxis_title="")
            fig.update_traces(texttemplate="%{x:.2f} h", textposition="outside")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Aucun arrêt enregistré pour la sélection actuelle.")
    else:
        st.info("Pas de données disponibles.")

st.subheader("Code d'arrêt (catégorie) par machine")
st.caption("Clique sur une barre pour voir le détail des événements correspondants (nom de l'arrêt et commentaire).")
if not raw_log_f.empty and not dt_legend.empty and "CodeDT_num" in dt_legend.columns:
    raw_with_cat_m = raw_log_f.merge(
        dt_legend[["CodeDT_num", "Categorie", "ColorHEX"]],
        left_on="Downtime reason", right_on="CodeDT_num", how="left",
    )
    raw_with_cat_m["Categorie"] = raw_with_cat_m["Categorie"].fillna("Non catégorisé")
    only_dt = raw_with_cat_m[raw_with_cat_m["Downtime reason"] != 0]

    df_mc = (
        only_dt.groupby(["Machine", "Categorie"])["confessed [h]"].sum()
        .reset_index()
    )
    if not df_mc.empty and df_mc["confessed [h]"].sum() > 0:
        cat_color_map_m = (
            raw_with_cat_m.dropna(subset=["Categorie"])
            .drop_duplicates(subset=["Categorie"])
            .set_index("Categorie")["ColorHEX"].to_dict()
        )
        machine_order = (
            df_mc.groupby("Machine")["confessed [h]"].sum()
            .sort_values(ascending=False).index.tolist()
        )
        fig_m = px.bar(
            df_mc, x="Machine", y="confessed [h]", color="Categorie",
            category_orders={"Machine": machine_order},
            color_discrete_map=cat_color_map_m,
            custom_data=["Categorie"],
        )
        fig_m.update_layout(height=380, margin=dict(t=10), barmode="stack",
                             xaxis_title="", yaxis_title="Heures d'arrêt")
        event_m = st.plotly_chart(
            fig_m, use_container_width=True,
            on_select="rerun", selection_mode="points", key="machine_cat_chart",
        )

        points = (event_m or {}).get("selection", {}).get("points", [])
        if points:
            p = points[0]
            machine_sel = p.get("x")
            cat_sel = None
            if "customdata" in p and p["customdata"]:
                cat_sel = p["customdata"][0]
            detail = only_dt[
                (only_dt["Machine"] == machine_sel) & (only_dt["Categorie"] == cat_sel)
            ][["Machine", "Downtime name", "Commentaire", "confessed [h]", "Start", "End"]]
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

with st.expander("🔍 Voir les données détaillées filtrées"):
    if not raw_log_f.empty:
        st.write("**RAW_LOG**")
        st.dataframe(raw_log_f, use_container_width=True)
    if not summary_f.empty:
        st.write("**SUMMARY**")
        st.dataframe(summary_f, use_container_width=True)
    if not dt_legend.empty:
        st.write("**DT_LEGEND**")
        st.dataframe(dt_legend, use_container_width=True)
