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


def _first_matching_sheet(filepath: str, candidates: list) -> str | None:
    """Retourne le 1er nom d'onglet existant parmi une liste de noms possibles."""
    try:
        xls = pd.ExcelFile(filepath)
        for name in candidates:
            if name in xls.sheet_names:
                return name
    except Exception:
        pass
    return None


def read_raw_log_from_file(filepath: str) -> pd.DataFrame:
    """Lit le journal d'événements (onglet 'DATA 1', accepte aussi 'RAW_LOG')."""
    sheet = _first_matching_sheet(filepath, ["DATA 1", "RAW_LOG"])
    if sheet is None:
        st.warning(f"Aucun onglet journal (DATA 1 / RAW_LOG) trouvé dans {os.path.basename(filepath)}")
        return pd.DataFrame()

    df = pd.read_excel(filepath, sheet_name=sheet)
    df = _clean_columns(df)

    required = ["Machine", "Start", "End", "Downtime reason",
                "Downtime name", "run [h]", "confessed [h]"]
    df = df[[c for c in required if c in df.columns]].copy()

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
    df["Fichier_Source"] = os.path.basename(filepath)
    return df


def read_summary_from_file(filepath: str) -> pd.DataFrame:
    """Lit le résumé quotidien (onglet 'DATA 2', accepte aussi 'SUMMARY')."""
    sheet = _first_matching_sheet(filepath, ["DATA 2", "SUMMARY"])
    if sheet is None:
        st.warning(f"Aucun onglet résumé (DATA 2 / SUMMARY) trouvé dans {os.path.basename(filepath)}")
        return pd.DataFrame()

    df = pd.read_excel(filepath, sheet_name=sheet)
    df = _clean_columns(df)
    df = df.rename(columns={
        "machine": "Machine",
        "Efficiency factor regarding cycle": "Efficiency",
    })

    required = ["Date", "Machine", "Yield", "Scrap", "Efficiency", "CC"]
    existing = [c for c in required if c in df.columns]
    if not existing:
        st.warning(f"Aucune colonne requise trouvée dans {os.path.basename(filepath)}")
        return pd.DataFrame()
    df = df[existing].copy()

    df["Yield"] = pd.to_numeric(df["Yield"], errors="coerce").fillna(0)
    df["Scrap"] = pd.to_numeric(df["Scrap"], errors="coerce").fillna(0)
    df["Efficiency"] = pd.to_numeric(df.get("Efficiency"), errors="coerce")

    df = df.dropna(subset=["Machine"])
    df = df[df["Machine"].astype(str).str.strip() != ""]

    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date
    df["Fichier_Source"] = os.path.basename(filepath)
    return df


def read_dt_legend_from_file(filepath: str) -> pd.DataFrame:
    """Lit la légende des codes d'arrêt (onglet 'DT', accepte aussi 'DT_LEGEND')."""
    sheet = _first_matching_sheet(filepath, ["DT", "DT_LEGEND"])
    if sheet is None:
        return pd.DataFrame(columns=["CodeDT", "CodeDT_num", "Libelle", "Categorie", "ColorHEX"])

    # L'onglet "DT" du logiciel machine a son en-tête en ligne 4 (pas en ligne 1)
    try:
        df = pd.read_excel(filepath, sheet_name=sheet, header=3)
        if not {"Downtime name", "Downtime reason"}.issubset({c.strip() if isinstance(c, str) else c for c in df.columns}):
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
def load_all_data(folder: str):
    files = sorted(glob.glob(os.path.join(folder, "*.xlsx")))
    files = [f for f in files if not os.path.basename(f).startswith("~$")]

    raw_logs, summaries, errors = [], [], []
    dt_legend = pd.DataFrame()

    for f in files:
        try:
            r = read_raw_log_from_file(f)
            if not r.empty:
                raw_logs.append(r)
        except Exception as e:
            errors.append(f"{os.path.basename(f)} (journal) : {e}")
        try:
            s = read_summary_from_file(f)
            if not s.empty:
                summaries.append(s)
        except Exception as e:
            errors.append(f"{os.path.basename(f)} (résumé) : {e}")
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


# ----------------------------------------------------------------------------
# Interface
# ----------------------------------------------------------------------------
st.title("📊 Dashboard OEE — Département Production")

with st.sidebar:
    st.header("⚙️ Source des données")
    folder = st.text_input(
        "Dossier des exports quotidiens",
        value=r"C:\OEE\Exports_Quotidiens",
        help="Dossier contenant vos fichiers Excel journaliers.",
    )
    if st.button("🔄 Actualiser les données", use_container_width=True):
        st.cache_data.clear()

if not os.path.isdir(folder):
    st.error(f"Le dossier indiqué n'existe pas : `{folder}`.")
    st.stop()

raw_log, summary, dt_legend, files, errors = load_all_data(folder)

if errors:
    with st.expander("⚠️ Fichiers non lus correctement"):
        for e in errors:
            st.write("-", e)

if raw_log.empty and summary.empty:
    st.warning("Aucune donnée exploitable trouvée dans ce dossier pour le moment.")
    st.stop()

st.caption(f"📁 {len(files)} fichier(s) chargé(s) depuis `{folder}`.")

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
# Rangée 3 : répartition des arrêts (donut) + top causes (barres)
# ----------------------------------------------------------------------------
st.subheader("📉 Analyse des temps d'arrêt")
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

st.subheader("Temps d'arrêt confessé par machine")
if not raw_log_f.empty:
    df_m = (raw_log_f.groupby("Machine")["confessed [h]"].sum().reset_index()
            .sort_values("confessed [h]", ascending=False))
    fig = px.bar(df_m, x="Machine", y="confessed [h]", color_discrete_sequence=[APTIV_BLUE])
    fig.update_layout(height=320, margin=dict(t=10))
    st.plotly_chart(fig, use_container_width=True)

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
