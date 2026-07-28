import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import datetime
import io
import requests

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(page_title="Tableau de bord Commercial", layout="wide", page_icon="📊")

st.markdown("""
    <style>
    .stMetric {
        background-color: #ffffff;
        padding: 15px;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }

    /* ── KPI cards ── */
    .kpi-card {
        border-radius: 10px;
        padding: 14px 16px 12px 16px;
        border: 1.5px solid #e2e8f0;
        box-shadow: 0 2px 6px rgba(0,0,0,0.06);
        height: 100%;
        display: flex;
        flex-direction: column;
        gap: 4px;
    }
    .kpi-card.green  { background: #f0fdf4; border-color: #86efac; }
    .kpi-card.red    { background: #fff1f2; border-color: #fca5a5; }
    .kpi-card.neutral{ background: #f8fafc; border-color: #cbd5e1; }
    .kpi-card.accent { background: #fffbeb; border-color: #fcd34d; }

    .kpi-label {
        font-size: 0.75rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: #64748b;
        margin: 0;
    }
    .kpi-value {
        font-size: 1.65rem;
        font-weight: 700;
        color: #1e293b;
        line-height: 1.1;
        margin: 2px 0 6px 0;
    }
    .kpi-delta {
        font-size: 0.78rem;
        font-weight: 500;
        line-height: 1.55;
        color: #64748b;
    }
    .kpi-delta .up   { color: #16a34a; }
    .kpi-delta .down { color: #dc2626; }
    .kpi-delta .flat { color: #64748b; }

    /* operator symbols between cards */
    .formula-op {
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.6rem;
        font-weight: 700;
        color: #94a3b8;
        padding: 0 4px;
        padding-top: 30px;
    }
    </style>
""", unsafe_allow_html=True)

liste_mois_noms = [
    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"
]

# --- CHARGEMENT DES DONNÉES DEPUIS GITHUB (LBFI OU PONCEBLANC) ---
@st.cache_data
def load_data(bu_name):
    url_github = "https://raw.githubusercontent.com/maximeponceblanc-a11y/LBFICommercial/main/Query_tableau_devis.xlsx"

    try:
        response = requests.get(url_github)
        response.raise_for_status()
        # On charge la feuille sélectionnée (LBFI ou PONCEBLANC)
        df = pd.read_excel(io.BytesIO(response.content), sheet_name=bu_name)
    except Exception as e:
        st.error(f"Erreur lors de la récupération du fichier sur GitHub pour {bu_name} : {e}")
        st.stop()

    df.columns = df.columns.str.strip()

    # --- HARMONISATION DES COLONNES PONCEBLANC TO LBFI ---
    if bu_name == "PONCEBLANC":
        df = df.rename(columns={
            'Dates': 'Date devis',
            'Taux Marge': 'Taux de marge',
            'Date Ouverture': 'Date ouverture dossier fab'
        })

    # --- NETTOYAGE ET PRÉPARATION DES DONNÉES ---
    if 'Date devis' in df.columns:
        s_numeric = pd.to_numeric(df['Date devis'], errors='coerce')
        df['Dates_Propres'] = pd.to_datetime(s_numeric, unit='D', origin='1899-12-30', errors='coerce')
        df['Dates_Propres'] = df['Dates_Propres'].fillna(pd.to_datetime(df['Date devis'], errors='coerce'))
        df = df.dropna(subset=['Dates_Propres'])

    if 'Dates_Propres' in df.columns:
        df['Mois_Num'] = df['Dates_Propres'].dt.month.fillna(1).astype(int)
        df['Année Devis'] = df['Dates_Propres'].dt.year.fillna(2026).astype(int)
    else:
        df['Mois_Num'] = 1
        df['Année Devis'] = 2026

    mois_fr = {
        1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril", 5: "Mai", 6: "Juin",
        7: "Juillet", 8: "Août", 9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre"
    }
    df['Mois_Nom'] = df['Mois_Num'].map(mois_fr).fillna("Janvier")

    df['Prix total'] = pd.to_numeric(df['Prix total'], errors='coerce')

    if 'Taux de marge' in df.columns:
        df['Taux de marge'] = pd.to_numeric(df['Taux de marge'], errors='coerce')
        df['Taux de marge'] = df['Taux de marge'].where(df['Taux de marge'].between(0, 100))

    if 'Nb exemplaires' in df.columns:
        df['Nb exemplaires'] = pd.to_numeric(df['Nb exemplaires'], errors='coerce')
        df['Prix unitaire'] = df['Prix total'] / df['Nb exemplaires'].replace(0, pd.NA)
        df['Prix unitaire'] = df['Prix unitaire'].where(df['Prix unitaire'] > 0)

    df['Prix total'] = df['Prix total'].fillna(0)

    if 'Signé ?' in df.columns:
        df = df.rename(columns={'Signé ?': 'Signé?'})
    if 'Signé?' in df.columns:
        df['Signé?'] = df['Signé?'].astype(str).str.strip().str.upper()
        df['Signé?'] = df['Signé?'].map({'VRAI': 'O', 'TRUE': 'O', 'O': 'O', 'FAUX': 'N', 'FALSE': 'N', 'N': 'N'}).fillna('N')

    if 'Date ouverture dossier fab' in df.columns and 'Dates_Propres' in df.columns:
        date_ouv_num = pd.to_numeric(df['Date ouverture dossier fab'], errors='coerce')
        date_ouv = pd.to_datetime(date_ouv_num, unit='D', origin='1899-12-30', errors='coerce')
        date_ouv = date_ouv.fillna(pd.to_datetime(df['Date ouverture dossier fab'], errors='coerce'))
        df['Délai devis ouverture'] = (date_ouv - df['Dates_Propres']).dt.days
        df['Délai devis ouverture'] = df['Délai devis ouverture'].where(df['Délai devis ouverture'] >= 0)

    if 'Numéro de devis' in df.columns:
        df = df.rename(columns={'Numéro de devis': 'DEVIS N°'})

    return df

# ========================================================
# LOGIQUE DE CLASSIFICATION ABC PAR ANNÉE CALENDAIRE
# ========================================================
def calculer_classification_abc_par_annee(dataframe):
    if dataframe.empty or 'Année Devis' not in dataframe.columns:
        dataframe['Catégorie Client ABC'] = "NOUVEAU CLIENT"
        return dataframe

    mask_signe_global = dataframe["Signé?"].astype(str).str.upper().str.strip() == "O"
    annees_triees = sorted(dataframe['Année Devis'].unique())

    clients_connus_avant = {}
    clients_cumul = set()
    for annee in annees_triees:
        clients_connus_avant[annee] = frozenset(clients_cumul)
        nouveaux = set(
            dataframe.loc[
                (dataframe['Année Devis'] == annee) & mask_signe_global,
                'Nom Client'
            ].dropna().unique()
        )
        clients_cumul |= nouveaux

    categories = pd.Series("NOUVEAU CLIENT", index=dataframe.index)

    for annee in annees_triees:
        mask_annee = dataframe['Année Devis'] == annee
        df_annee = dataframe[mask_annee]
        connus = clients_connus_avant[annee]

        df_annee_signe = df_annee[mask_signe_global.reindex(df_annee.index, fill_value=False)]
        ca_par_client = (
            df_annee_signe.groupby('Nom Client')['Prix total']
            .sum()
            .sort_values(ascending=False)
            .reset_index()
        )
        ca_par_client_connus = ca_par_client[ca_par_client['Nom Client'].isin(connus)]

        total_ca = ca_par_client_connus['Prix total'].sum()
        dict_abc = {}

        if total_ca > 0:
            ca_par_client_connus = ca_par_client_connus.copy()
            ca_par_client_connus['CA_Cumule_Pct'] = (
                ca_par_client_connus['Prix total'].cumsum() / total_ca * 100
            )
            for _, row in ca_par_client_connus.iterrows():
                client = row['Nom Client']
                pct = row['CA_Cumule_Pct']
                if pct <= 80.001:
                    dict_abc[client] = "GRAND COMPTE"
                elif pct <= 90.001:
                    dict_abc[client] = "CLIENT INTERMEDIAIRE"
                else:
                    dict_abc[client] = "PETITS CLIENTS"

        for idx, row in df_annee.iterrows():
            client = row['Nom Client']
            if client not in connus:
                categories.at[idx] = "NOUVEAU CLIENT"
            elif client in dict_abc:
                categories.at[idx] = dict_abc[client]
            else:
                categories.at[idx] = "PETITS CLIENTS"

    dataframe = dataframe.copy()
    dataframe['Catégorie Client ABC'] = categories
    return dataframe

# ========================================================
# BARRE LATÉRALE : SYSTÈME DE FILTRES GLOBAUX
# ========================================================
st.sidebar.header("🎛️ Filtres & Paramètres")

# Bouton à bascule pour changer de Business Unit (PONCEBLANC par défaut)
bu_toggle = st.sidebar.toggle("📈 Basculer vers LBFI (Par défaut : PONCEBLANC)", value=False)
bu_name = "LBFI" if bu_toggle else "PONCEBLANC"

# Chargement dynamique des données de la BU sélectionnée
try:
    df = load_data(bu_name)
except Exception as e:
    st.stop()

# Calcul de la classification ABC
df = calculer_classification_abc_par_annee(df)

mask_metier = pd.Series(True, index=df.index)

# ========================================================
# ENTÊTE DU TABLEAU DE BORD
# ========================================================
st.title(f"📊 Tableau de bord commercial — {bu_name}")

st.sidebar.markdown("### 📅 Filtres Temporels Globaux")

vue_mode = st.sidebar.radio(
    "Maille d'analyse :",
    options=["📅 Mensuelle", "📆 Annuelle"],
    horizontal=False,
    help="Mensuelle : évolution mois par mois sur une année. Annuelle : comparaison inter-années."
)
vue_annuelle = (vue_mode == "📆 Annuelle")

if 'Année Devis' in df.columns:
    annees_possibles = sorted([int(a) for a in df['Année Devis'].unique() if a > 2000], reverse=True)
else:
    annees_possibles = [2026, 2025]

annees_selectionnees = st.sidebar.multiselect(
    "Années d'analyse :",
    options=annees_possibles,
    default=annees_possibles
)

if not vue_annuelle:
    start_month, end_month = st.sidebar.select_slider(
        "Intervalle des mois inclus :",
        options=liste_mois_noms,
        value=("Janvier", "Décembre")
    )
else:
    st.sidebar.markdown("<span style='color:#64748b;font-size:0.9em;'>En vue annuelle, tous les mois sont inclus.</span>", unsafe_allow_html=True)
    start_month, end_month = "Janvier", "Décembre"

mois_debut_num = liste_mois_noms.index(start_month) + 1
mois_fin_num = liste_mois_noms.index(end_month) + 1
mois_selectionnes_bornes = liste_mois_noms[mois_debut_num - 1:mois_fin_num]

mask_temporel = (df['Année Devis'].isin(annees_selectionnees)) & (df['Mois_Num'] >= mois_debut_num) & (df['Mois_Num'] <= mois_fin_num)

st.sidebar.markdown("### 📁 Dossier")

# --- NOUVEAU FILTRE : CHAMPS NUMÉRIQUES POUR MONTANT ---
st.sidebar.markdown("### 💰 Filtre Montant (€)")

if 'Prix total' in df.columns and not df.empty:
    min_val_reel = float(df['Prix total'].min())
    max_val_reel = float(df['Prix total'].max())
    
    # Initialisation des colonnes de saisie
    col1, col2 = st.sidebar.columns(2)
    
    with col1:
        montant_min = st.number_input(
            "Min (€)", 
            min_value=0.0, 
            max_value=max_val_reel, 
            value=0.0, 
            step=100.0
        )
    with col2:
        montant_max = st.number_input(
            "Max (€)", 
            min_value=0.0, 
            max_value=max_val_reel, 
            value=max_val_reel, 
            step=100.0
        )

    # Application du masque de filtrage
    mask_metier &= df["Prix total"].between(montant_min, montant_max)
else:
    montant_min, montant_max = 0.0, 0.0

selected_abc = st.sidebar.multiselect(
    "Catégorie Client (ABC)",
    options=["GRAND COMPTE", "CLIENT INTERMEDIAIRE", "PETITS CLIENTS", "NOUVEAU CLIENT"],
    default=["GRAND COMPTE", "CLIENT INTERMEDIAIRE", "PETITS CLIENTS", "NOUVEAU CLIENT"]
)

search_client = st.sidebar.text_input("Rechercher un Nom Client")
search_devis = st.sidebar.text_input("Rechercher un N° de Devis")

if selected_abc and "Catégorie Client ABC" in df.columns:
    mask_metier &= df["Catégorie Client ABC"].isin(selected_abc)
if search_client and "Nom Client" in df.columns:
    mask_metier &= df["Nom Client"].astype(str).str.contains(search_client, case=False, na=False)
if search_devis and "DEVIS N°" in df.columns:
    mask_metier &= df["DEVIS N°"].astype(str).str.contains(search_devis, case=False, na=False)

for col_df, label in [("Type de produit", "Type de produit")]:
    if col_df in df.columns:
        options = sorted([str(x) for x in df[col_df].dropna().unique()])
        selected = st.sidebar.multiselect(label, options=options)
        if selected:
            mask_metier &= df[col_df].astype(str).isin(selected)

# --- FILTRE COMMERCIAL (visible uniquement lorsque le toggle est sur PONCEBLANC) ---
selected_commercial = []
if bu_name == "PONCEBLANC" and "Commercial" in df.columns:
    st.sidebar.markdown("### 🧑‍💼 Filtre Commercial")
    options_commercial = sorted([str(x) for x in df["Commercial"].dropna().unique()])
    selected_commercial = st.sidebar.multiselect("Commercial", options=options_commercial)
    if selected_commercial:
        mask_metier &= df["Commercial"].astype(str).isin(selected_commercial)

st.sidebar.markdown("### 📈 Analyse avancée")
show_trends = st.sidebar.checkbox("Afficher les courbes de tendance", value=False, help="Affiche les tendances linéaires et le R² pour analyser la dynamique sur les graphiques de nuage de points.")

df_filtered_base = df[mask_metier & mask_temporel].copy()
st.sidebar.success(f"Données filtrées : {len(df_filtered_base)} lignes.")

def extraire_kpis_annee(dataframe, annee_cible):
    df_cible = dataframe[dataframe['Année Devis'] == annee_cible]
    col_id = "DEVIS N°" if "DEVIS N°" in dataframe.columns else dataframe.columns[0]

    if df_cible.empty:
        return None

    mask_signe = df_cible["Signé?"].astype(str).str.upper().str.strip() == "O"

    ca_devis = df_cible["Prix total"].sum()
    ca_signe = df_cible[mask_signe]["Prix total"].sum()
    tx_succes_ca = (ca_signe / ca_devis * 100) if ca_devis > 0 else 0

    vol_devis = df_cible[col_id].nunique()
    vol_signe = df_cible[mask_signe][col_id].nunique()
    tx_succes_vol = (vol_signe / vol_devis * 100) if vol_devis > 0 else 0

    cmd_moy_tous = (ca_devis / vol_devis) if vol_devis > 0 else 0
    cmd_moy_signe = (ca_signe / vol_signe) if vol_signe > 0 else 0

    return {
        "ca_devis": ca_devis, "ca_signe": ca_signe, "tx_ca": tx_succes_ca,
        "vol_devis": vol_devis, "vol_signe": vol_signe, "tx_vol": tx_succes_vol,
        "cmd_moy_tous": cmd_moy_tous, "cmd_moy_signe": cmd_moy_signe
    }


# ========================================================
# HELPER : CARTE KPI « FORMULE MAGIQUE »
# ========================================================
def kpi_card(label, value, delta_prev=None, delta_next=None,
             annee_prev=None, annee_next=None,
             unite="€", is_pct=False, force_color=None):
    """
    Renders a styled KPI card using HTML.
    force_color: 'green' | 'red' | 'accent' | 'neutral' | None (auto from delta_prev)
    """

    def _delta_html(val_curr, kpis_ref, label_annee):
        if kpis_ref is None:
            return f"<span class='flat'>vs {label_annee} : —</span>"
        # kpis_ref is already the numeric diff value passed directly
        diff = val_curr
        if diff > 0.001:
            cls = "up"; arrow = "↗"; sign = "+"
        elif diff < -0.001:
            cls = "down"; arrow = "↘"; sign = ""
        else:
            cls = "flat"; arrow = "→"; sign = ""
        if is_pct:
            txt = f"{sign}{diff:.2f} pts".replace('.', ',')
        else:
            txt = f"{sign}{diff:,.0f} {unite}".replace(',', ' ')
        return f"<span class='{cls}'>{arrow} vs {label_annee} : {txt}</span>"

    # Determine card color
    if force_color:
        color_cls = force_color
    elif delta_prev is not None:
        color_cls = "green" if delta_prev > 0.001 else ("red" if delta_prev < -0.001 else "neutral")
    else:
        color_cls = "neutral"

    # Build delta lines
    delta_html = ""
    if delta_prev is not None and annee_prev is not None:
        delta_html += _delta_html(delta_prev, True, annee_prev) + "<br>"
    if delta_next is not None and annee_next is not None:
        delta_html += _delta_html(delta_next, True, annee_next)

    delta_block = f"<div class='kpi-delta'>{delta_html}</div>" if delta_html else ""

    html = f"""
    <div class="kpi-card {color_cls}">
        <p class="kpi-label">{label}</p>
        <p class="kpi-value">{value}</p>
        {delta_block}
    </div>
    """
    st.html(html)


def compute_delta(val_curr, kpis_ref, cle):
    """Returns numeric diff or None if no reference."""
    if kpis_ref is None:
        return None
    return val_curr - kpis_ref[cle]


def fmt_val(v, unite="€", is_pct=False, decimals=0):
    if is_pct:
        return f"{v:.2f} %".replace('.', ',')
    if decimals:
        return f"{v:,.{decimals}f} {unite}".replace(',', ' ')
    return f"{int(round(v)):,} {unite}".replace(',', ' ') if unite else f"{int(round(v)):,}"


def render_kpi_section(kpis_current, kpis_prev, kpis_next,
                       annee_selectionnee, annee_prev, annee_next):
    if not kpis_current:
        st.info("Aucun indicateur disponible pour cette période.")
        return

    kp = kpis_current

    def d_prev(cle):
        return compute_delta(kp[cle], kpis_prev, cle)

    def d_next(cle):
        return compute_delta(kp[cle], kpis_next, cle)

    tx_ca_color = None
    if kpis_prev:
        tx_ca_color = "green" if kp['tx_ca'] >= kpis_prev['tx_ca'] else "red"

    cmd_moy_signe_color = None
    if kpis_prev:
        cmd_moy_signe_color = "green" if kp['cmd_moy_signe'] >= kpis_prev['cmd_moy_signe'] else "red"

    # ── ROW 1 : FORMULE VOLUME ──
    st.markdown("<div style='margin-bottom:6px; font-size:0.8rem; color:#94a3b8; font-weight:600; letter-spacing:0.05em;'>📦 FORMULE VOLUME : Nb devis × % Succès (vol) = Nb commandes</div>", unsafe_allow_html=True)
    cols1 = st.columns([4, 0.5, 4, 0.5, 4, 0.5, 4])
    with cols1[0]:
        kpi_card(
            "Nombre de devis émis",
            fmt_val(kp['vol_devis'], unite="devis"),
            delta_prev=d_prev('vol_devis'), delta_next=d_next('vol_devis'),
            annee_prev=annee_prev, annee_next=annee_next,
            unite="devis",
        )
    with cols1[1]:
        st.markdown("<div class='formula-op'>×</div>", unsafe_allow_html=True)
    with cols1[2]:
        kpi_card(
            "% Succès (Volume)",
            fmt_val(kp['tx_vol'], is_pct=True),
            delta_prev=d_prev('tx_vol'), delta_next=d_next('tx_vol'),
            annee_prev=annee_prev, annee_next=annee_next,
            is_pct=True,
        )
    with cols1[3]:
        st.markdown("<div class='formula-op'>=</div>", unsafe_allow_html=True)
    with cols1[4]:
        kpi_card(
            "Nombre de commandes",
            fmt_val(kp['vol_signe'], unite="signés"),
            delta_prev=d_prev('vol_signe'), delta_next=d_next('vol_signe'),
            annee_prev=annee_prev, annee_next=annee_next,
            unite="signés",
        )
    
    # ── ROW 2 : FLÈCHES 2D ET MULTIPLICATEURS VERTICAUX ──
    cols2 = st.columns([4, 0.5, 4, 0.5, 4, 0.5, 4])
    
    with cols2[0]:
        st.write("")
    with cols2[1]:
        st.write("")
    with cols2[2]:
        kpi_card(
            "🛒 Devis moyen",
            fmt_val(kp['cmd_moy_tous']),
            delta_prev=d_prev('cmd_moy_tous'), delta_next=d_next('cmd_moy_tous'),
            annee_prev=annee_prev, annee_next=annee_next,
            force_color="accent",
        )

    with cols2[4]:
        arrow_svg = """
        <div style='display:flex; justify-content:center; align-items:center; padding-top: 15px; padding-bottom: 5px;'>
            <svg width='30' height='60' viewBox='0 0 24 60' fill='none' stroke='#cbd5e1' stroke-width='4' stroke-linecap='round' stroke-linejoin='round'>
                <line x1='12' y1='5' x2='12' y2='55'></line>
                <polyline points='5 45 12 55 19 45'></polyline>
            </svg>
        </div>
        """
        st.markdown(arrow_svg, unsafe_allow_html=True)
    with cols2[5]:
        st.markdown("<div class='formula-op' style='padding-top:30px;'>×</div>", unsafe_allow_html=True)
    with cols2[6]:
        kpi_card(
            "✅ Commande moy. (signés)",
            fmt_val(kp['cmd_moy_signe']),
            delta_prev=d_prev('cmd_moy_signe'), delta_next=d_next('cmd_moy_signe'),
            annee_prev=annee_prev, annee_next=annee_next,
            force_color=cmd_moy_signe_color,
        )

    # ── ROW 3 : FORMULE CA ──
    st.markdown("<div style='margin-bottom:6px; margin-top:10px; font-size:0.8rem; color:#94a3b8; font-weight:600; letter-spacing:0.05em;'>💰 FORMULE CA : CA Devisé × % Succès (€) = CA Commandes</div>", unsafe_allow_html=True)
    cols3 = st.columns([4, 0.5, 4, 0.5, 4, 0.5, 4])
    with cols3[0]:
        kpi_card(
            f"CA Devisé {annee_selectionnee}",
            fmt_val(kp['ca_devis']),
            delta_prev=d_prev('ca_devis'), delta_next=d_next('ca_devis'),
            annee_prev=annee_prev, annee_next=annee_next,
        )
    with cols3[1]:
        st.markdown("<div class='formula-op'>×</div>", unsafe_allow_html=True)
    with cols3[2]:
        kpi_card(
            "% Succès (€)",
            fmt_val(kp['tx_ca'], is_pct=True),
            delta_prev=d_prev('tx_ca'), delta_next=d_next('tx_ca'),
            annee_prev=annee_prev, annee_next=annee_next,
            is_pct=True,
            force_color=tx_ca_color,
        )
    with cols3[3]:
        st.markdown("<div class='formula-op'>=</div>", unsafe_allow_html=True)
    with cols3[4]:
        kpi_card(
            f"CA Commandes {annee_selectionnee}",
            fmt_val(kp['ca_signe']),
            delta_prev=d_prev('ca_signe'), delta_next=d_next('ca_signe'),
            annee_prev=annee_prev, annee_next=annee_next,
        )


# ========================================================
# RENDER : VUE UNIQUE PERFORMANCE COMMERCIALE
# ========================================================
st.subheader("Performance Commerciale")

if vue_annuelle and len(annees_selectionnees) > 0:
    annees_triees = sorted(annees_selectionnees)

    rows = []
    for a in annees_triees:
        k = extraire_kpis_annee(df_filtered_base, a)
        if k:
            rows.append({"Année": str(a), **k})
    df_kpis_all = pd.DataFrame(rows)

    annee_courante = max(annees_triees)
    annee_precedente = annee_courante - 1
    annee_en_cours_label = " — Année en cours" if annee_courante == datetime.datetime.now().year else ""

    kpis_courant = extraire_kpis_annee(df_filtered_base, annee_courante)
    kpis_precedent = extraire_kpis_annee(df[mask_metier], annee_precedente)

    ref = kpis_precedent

    st.markdown(
        f"##### 📌 Indicateurs Clés **{annee_courante}**"
        f"<span style='font-size:0.85em; font-weight:400; color:#64748b;'>"
        f"{annee_en_cours_label}"
        f"{'  ·  Δ vs ' + str(annee_precedente) if ref else ''}</span>",
        unsafe_allow_html=True
    )

    if kpis_courant:
        render_kpi_section(
            kpis_current=kpis_courant,
            kpis_prev=kpis_precedent,
            kpis_next=None,
            annee_selectionnee=annee_courante,
            annee_prev=annee_precedente,
            annee_next=None,
        )
        if ref:
            st.caption(f"↕ Deltas calculés par rapport à {annee_precedente}")

    st.divider()

    col_g, col_d = st.columns([6, 4])

    with col_g:
        st.markdown("**Performance Annuelle — Financement**")
        if not df_kpis_all.empty:
            fig_yr = make_subplots(specs=[[{"secondary_y": True}]])
            fig_yr.add_trace(go.Bar(x=df_kpis_all['Année'], y=df_kpis_all['ca_devis'], name="CA Devisé", marker_color='#00a4bd'), secondary_y=False)
            fig_yr.add_trace(go.Bar(x=df_kpis_all['Année'], y=df_kpis_all['ca_signe'], name="CA Commandes", marker_color='#4ed2e6'), secondary_y=False)
            fig_yr.add_trace(go.Scatter(x=df_kpis_all['Année'], y=df_kpis_all['tx_ca'] / 100, name="Taux succès (€)", mode='lines+markers+text',
                text=[f"{v:.1f}%".replace('.', ',') for v in df_kpis_all['tx_ca']], textposition='top center',
                line=dict(color='#0b5ca3', width=2)), secondary_y=True)
            fig_yr.update_layout(barmode='group', xaxis=dict(type='category'), yaxis=dict(title="Montant (€)"),
                yaxis2=dict(title="Taux succès", tickformat=".0%", range=[0, 1.2], showgrid=False),
                legend=dict(orientation="h", y=1.1), hovermode="x unified", height=380)
            st.plotly_chart(fig_yr, use_container_width=True, key="fig_yr")

        st.write("---")
        st.markdown("**Analyse Annuelle — Volumes & Taux de Succès**")
        if not df_kpis_all.empty:
            fig_vol_yr = make_subplots(specs=[[{"secondary_y": True}]])
            fig_vol_yr.add_trace(go.Bar(x=df_kpis_all['Année'], y=df_kpis_all['vol_devis'], name="Nb devis émis", marker_color='#ff9f00'), secondary_y=False)
            fig_vol_yr.add_trace(go.Bar(x=df_kpis_all['Année'], y=df_kpis_all['vol_signe'], name="Nb commandes", marker_color='#ffcc66'), secondary_y=False)
            fig_vol_yr.add_trace(go.Scatter(x=df_kpis_all['Année'], y=df_kpis_all['tx_vol'] / 100, name="Taux succès (vol)", mode='lines+markers+text',
                text=[f"{v:.1f}%".replace('.', ',') for v in df_kpis_all['tx_vol']], textposition='top center',
                line=dict(color='#e65c00', width=2)), secondary_y=True)
            fig_vol_yr.update_layout(barmode='group', xaxis=dict(type='category'), yaxis=dict(title="Nombre de devis"),
                yaxis2=dict(title="Taux succès", tickformat=".0%", range=[0, 1.2], showgrid=False),
                legend=dict(orientation="h", y=1.1), hovermode="x unified", height=380)
            st.plotly_chart(fig_vol_yr, use_container_width=True, key="fig_vol_yr")

    with col_d:
        st.markdown("**Part du CA signé par Catégorie ABC (toutes années)**")
        colors_map_abc = {"GRAND COMPTE": "#00a4bd", "CLIENT INTERMEDIAIRE": "#4ed2e6", "PETITS CLIENTS": "#ff9f00", "NOUVEAU CLIENT": "#94a3b8"}
        df_pie_all = df_filtered_base[df_filtered_base["Signé?"].astype(str).str.upper().str.strip() == "O"].groupby('Catégorie Client ABC')['Prix total'].sum().reset_index(name='Total CA')
        if not df_pie_all.empty:
            fig_pie_all = go.Figure(data=[go.Pie(
                labels=df_pie_all['Catégorie Client ABC'], values=df_pie_all['Total CA'], hole=.4,
                marker=dict(colors=[colors_map_abc.get(x, "#94a3b8") for x in df_pie_all['Catégorie Client ABC']]),
                textinfo='percent', hovertemplate="<b>%{label}</b><br>CA : %{value:,.0f} €<br>Part : %{percent}<extra></extra>"
            )])
        else:
            fig_pie_all = go.Figure(data=[go.Pie(
                labels=["Aucune commande"], values=[1], hole=.4,
                marker=dict(colors=["#e2e8f0"]), textinfo='none', hoverinfo='skip', showlegend=False,
            )])
            fig_pie_all.update_layout(annotations=[dict(text="Aucune<br>commande", x=0.5, y=0.5, font_size=13, showarrow=False, font_color="#94a3b8")])
        fig_pie_all.update_layout(legend=dict(orientation="h", y=-0.1, x=0), margin=dict(t=10, b=10, l=10, r=10), height=320)
        st.plotly_chart(fig_pie_all, use_container_width=True, key="fig_pie_all")

        st.markdown("**📋 Récapitulatif KPIs par Année**")
        if not df_kpis_all.empty:
            df_recap = df_kpis_all.copy()
            df_recap['CA Commandes'] = df_recap['ca_signe'].map(lambda x: f"{x:,.0f} €".replace(',', ' '))
            df_recap['CA Devisé'] = df_recap['ca_devis'].map(lambda x: f"{x:,.0f} €".replace(',', ' '))
            df_recap['Tx Succès (€)'] = df_recap['tx_ca'].map(lambda x: f"{x:.2f} %".replace('.', ','))
            df_recap['Commandes'] = df_recap['vol_signe'].astype(int)
            df_recap['Devis émis'] = df_recap['vol_devis'].astype(int)
            df_recap['Tx Succès (vol)'] = df_recap['tx_vol'].map(lambda x: f"{x:.2f} %".replace('.', ','))
            df_recap['Devis moyen'] = df_recap['cmd_moy_tous'].map(lambda x: f"{x:,.0f} €".replace(',', ' '))
            df_recap['Cmd moy. (signés)'] = df_recap['cmd_moy_signe'].map(lambda x: f"{x:,.0f} €".replace(',', ' '))
            st.dataframe(df_recap[['Année', 'CA Commandes', 'CA Devisé', 'Tx Succès (€)', 'Commandes', 'Devis émis', 'Tx Succès (vol)', 'Devis moyen', 'Cmd moy. (signés)']], use_container_width=True, hide_index=True)

elif not vue_annuelle and len(annees_selectionnees) > 0:
    tabs_val = st.tabs([f"Année {annee}" for annee in annees_selectionnees])

    for idx, annee_selectionnee in enumerate(annees_selectionnees):
        with tabs_val[idx]:
            st.markdown(f"##### 📌 Indicateurs Clés — {annee_selectionnee} ({start_month} à {end_month})")

            annee_prev = annee_selectionnee - 1
            annee_next = annee_selectionnee + 1

            kpis_current = extraire_kpis_annee(df_filtered_base, annee_selectionnee)
            kpis_prev = extraire_kpis_annee(df_filtered_base, annee_prev)
            kpis_next = extraire_kpis_annee(df_filtered_base, annee_next)

            if kpis_current:
                render_kpi_section(
                    kpis_current=kpis_current,
                    kpis_prev=kpis_prev,
                    kpis_next=kpis_next,
                    annee_selectionnee=annee_selectionnee,
                    annee_prev=annee_prev,
                    annee_next=annee_next,
                )

            st.divider()

            df_annee = df_filtered_base[df_filtered_base['Année Devis'] == annee_selectionnee]

            if not df_annee.empty:
                col_gauche, col_droite = st.columns([6.5, 3.5])

                with col_gauche:
                    st.markdown(f"**Performance Mensuelle CA ({annee_selectionnee})**")

                    base_ca = df_annee.groupby('Mois_Nom')['Prix total'].sum().reindex(mois_selectionnes_bornes, fill_value=0).reset_index(name="Chiffre d'affaire")
                    df_annee_signe = df_annee[df_annee["Signé?"].astype(str).str.upper().str.strip() == "O"]
                    signe_ca = df_annee_signe.groupby('Mois_Nom')['Prix total'].sum().reindex(mois_selectionnes_bornes, fill_value=0).reset_index(name='CA commandes')

                    df_graph = pd.merge(base_ca, signe_ca, on='Mois_Nom', how='left').fillna(0)
                    df_graph['Taux succès (€)'] = df_graph.apply(lambda row: (row['CA commandes'] / row["Chiffre d'affaire"]) if row["Chiffre d'affaire"] > 0 else 0, axis=1)

                    fig = make_subplots(specs=[[{"secondary_y": True}]])
                    fig.add_trace(go.Bar(x=df_graph['Mois_Nom'], y=df_graph["Chiffre d'affaire"], name="Chiffre d'affaire", marker_color='#00a4bd'), secondary_y=False)
                    fig.add_trace(go.Bar(x=df_graph['Mois_Nom'], y=df_graph['CA commandes'], name="CA commandes", marker_color='#4ed2e6'), secondary_y=False)
                    fig.add_trace(go.Scatter(x=df_graph['Mois_Nom'], y=df_graph['Taux succès (€)'], name="Taux succès (€)", mode='lines+markers', line=dict(color='#0b5ca3', width=2)), secondary_y=True)

                    fig.update_layout(barmode='group', xaxis=dict(type='category'), yaxis=dict(title="Montant (€)", showgrid=True),
                        yaxis2=dict(title="Taux succès (€)", tickformat=".2%", range=[0, 1], showgrid=False),
                        legend=dict(orientation="h", y=1.1, x=0), hovermode="x unified", height=380)
                    st.plotly_chart(fig, use_container_width=True, key=f"fig_ca_{annee_selectionnee}")

                    st.write("---")
                    st.markdown(f"**Analyse Mensuelle des Volumes & Taux de Succès ({annee_selectionnee})**")

                    col_id_devis = "DEVIS N°" if "DEVIS N°" in df_annee.columns else df_annee.columns[0]
                    base_vol = df_annee.groupby('Mois_Nom')[col_id_devis].nunique().reindex(mois_selectionnes_bornes, fill_value=0).reset_index(name='Nb devis')
                    signe_vol = df_annee_signe.groupby('Mois_Nom')[col_id_devis].nunique().reindex(mois_selectionnes_bornes, fill_value=0).reset_index(name='Nb commandes')

                    df_graph_vol = pd.merge(base_vol, signe_vol, on='Mois_Nom', how='left').fillna(0)
                    df_graph_vol['Taux succès (vol)'] = df_graph_vol.apply(lambda row: (row['Nb commandes'] / row['Nb devis']) if row['Nb devis'] > 0 else 0, axis=1)

                    fig_vol = make_subplots(specs=[[{"secondary_y": True}]])
                    fig_vol.add_trace(go.Bar(x=df_graph_vol['Mois_Nom'], y=df_graph_vol['Nb devis'], name="Nb devis", marker_color='#ff9f00'), secondary_y=False)
                    fig_vol.add_trace(go.Bar(x=df_graph_vol['Mois_Nom'], y=df_graph_vol['Nb commandes'], name="Nb commandes", marker_color='#ffcc66'), secondary_y=False)
                    fig_vol.add_trace(go.Scatter(x=df_graph_vol['Mois_Nom'], y=df_graph_vol['Taux succès (vol)'], name="Taux succès (vol)", mode='lines+markers', line=dict(color='#e65c00', width=2)), secondary_y=True)

                    fig_vol.update_layout(barmode='group', xaxis=dict(type='category'), yaxis=dict(title="Nombre de devis", showgrid=True),
                        yaxis2=dict(title="Taux succès (vol)", tickformat=".2%", range=[0, 1], showgrid=False),
                        legend=dict(orientation="h", y=1.1, x=0), hovermode="x unified", height=380)
                    st.plotly_chart(fig_vol, use_container_width=True, key=f"fig_vol_{annee_selectionnee}")

                with col_droite:
                    st.markdown(f"**CA par commande client ({annee_selectionnee})**")
                    df_annee_signe_client = df_annee[df_annee["Signé?"].astype(str).str.upper().str.strip() == "O"]
                    if "Nom Client" in df_annee_signe_client.columns and not df_annee_signe_client.empty:
                        df_c_annee = df_annee_signe_client.groupby("Nom Client")["Prix total"].sum().sort_values(ascending=False).reset_index()
                        df_c_annee.columns = ["Label", "Valeur"]
                        n_annee = len(df_c_annee)
                        palette_base_annee = px.colors.qualitative.Set2 + px.colors.qualitative.Pastel + px.colors.qualitative.Safe
                        couleurs_annee = [palette_base_annee[i % len(palette_base_annee)] for i in range(n_annee)]
                        fig_c_annee = go.Figure(data=[go.Pie(
                            labels=df_c_annee["Label"],
                            values=df_c_annee["Valeur"],
                            hole=0.4,
                            textinfo="percent",
                            hovertemplate="<b>%{label}</b><br>CA : %{value:,.0f} €<br>Part : %{percent}<extra></extra>",
                            marker=dict(colors=couleurs_annee),
                        )])
                        fig_c_annee.update_layout(
                            legend=dict(orientation="h", y=-0.05, x=0.5, xanchor="center", font=dict(size=10), itemwidth=30),
                            margin=dict(t=10, b=20, l=10, r=10),
                            height=320,
                        )
                        st.plotly_chart(fig_c_annee, use_container_width=True, key=f"fig_c_annee_{annee_selectionnee}")
                    else:
                        st.info(f"Aucune commande client pour {annee_selectionnee}.")

                    st.markdown(f"**📋 Liste des dossiers de devis ({annee_selectionnee})**")
                    df_table = pd.DataFrame()
                    df_table["Nom Client"] = df_annee["Nom Client"].values
                    df_table["DEVIS N°"] = df_annee["DEVIS N°"].values if "DEVIS N°" in df_annee.columns else "N/A"

                    df_table["Statut Devis"] = "Refusé"
                    mask_o = df_annee["Signé?"].astype(str).str.upper().str.strip() == "O"
                    df_table.loc[mask_o.values, "Statut Devis"] = "Commande"

                    df_table["CA"] = df_annee["Prix total"].values
                    df_table = df_table.sort_values(by="CA", ascending=False)

                    df_table_formatted = df_table.copy()
                    df_table_formatted["CA"] = df_table_formatted["CA"].map(lambda x: f"{x:,.0f} €".replace(',', ' '))

                    st.dataframe(df_table_formatted, use_container_width=True, hide_index=True, height=440)
            else:
                st.info(f"Aucune ligne de détail pour l'année {annee_selectionnee}.")

if len(annees_selectionnees) == 0:
    st.info("💡 Sélectionnez au moins une année dans les filtres de la barre latérale.")


# ========================================================
# SECTION : CAMEMBERTS COMPLÉMENTAIRES (CA SIGNÉ)
# ========================================================
st.divider()
st.subheader("Répartition du CA Commandes")

mask_signe_global = df_filtered_base["Signé?"].astype(str).str.upper().str.strip() == "O"
df_signe = df_filtered_base[mask_signe_global].copy()

for col_ni in ["Type de produit"]:
    if col_ni in df_signe.columns:
        df_signe[col_ni] = df_signe[col_ni].fillna("Non indiqué").replace("", "Non indiqué")

col_p1, col_p2 = st.columns(2)

with col_p1:
    if not df_signe.empty and 'Catégorie Client ABC' in df_signe.columns:
        colors_map_global = {"GRAND COMPTE": "#00a4bd", "CLIENT INTERMEDIAIRE": "#4ed2e6", "PETITS CLIENTS": "#ff9f00", "NOUVEAU CLIENT": "#94a3b8"}
        df_pie_global = df_signe.groupby('Catégorie Client ABC')['Prix total'].sum().reset_index(name='Total CA')
        if not df_pie_global.empty:
            fig_pie_global = go.Figure(data=[go.Pie(
                labels=df_pie_global['Catégorie Client ABC'], values=df_pie_global['Total CA'], hole=.4,
                marker=dict(colors=[colors_map_global.get(x, "#94a3b8") for x in df_pie_global['Catégorie Client ABC']]),
                textinfo='percent', hovertemplate="<b>%{label}</b><br>CA : %{value:,.0f} €<br>Part : %{percent}<extra></extra>"
            )])
        else:
            fig_pie_global = go.Figure(data=[go.Pie(
                labels=["Aucune commande"], values=[1], hole=.4,
                marker=dict(colors=["#e2e8f0"]), textinfo='none', hoverinfo='skip', showlegend=True,
            )])
            fig_pie_global.update_layout(annotations=[dict(text="Aucune<br>commande", x=0.5, y=0.5, font_size=13, showarrow=False, font_color="#94a3b8")])
        fig_pie_global.update_layout(
            title=dict(text="Part du CA signé par Catégorie ABC", font=dict(size=13), x=0),
            legend=dict(orientation="h", y=-0.1, x=0),
            margin=dict(t=40, b=10, l=10, r=10),
            height=460,
        )
        st.plotly_chart(fig_pie_global, use_container_width=True, key="fig_pie_global")
    else:
        st.info("Aucune donnée de catégorie ABC disponible.")

with col_p2:
    if "Type de produit" in df_signe.columns and not df_signe.empty:
        df_p2 = df_signe.groupby("Type de produit")["Prix total"].sum().sort_values(ascending=False).reset_index()
        df_p2.columns = ["Label", "Valeur"]
        fig_c2 = go.Figure(data=[go.Pie(
            labels=df_p2["Label"],
            values=df_p2["Valeur"],
            hole=0.4,
            textinfo="percent",
            hovertemplate="<b>%{label}</b><br>CA : %{value:,.0f} €<br>Part : %{percent}<extra></extra>",
        )])
        fig_c2.update_layout(
            title=dict(text="CA commandes par Type de produit", font=dict(size=13), x=0),
            legend=dict(orientation="h", y=-0.15, x=0),
            margin=dict(t=40, b=10, l=10, r=10),
            height=340,
        )
        st.plotly_chart(fig_c2, use_container_width=True, key="fig_c2")
    else:
        st.info("Colonne 'Type de produit' non disponible.")


# ========================================================
# SECTION : DÉLAI MOYEN D'OUVERTURE PAR MAILLE TEMPORELLE
# ========================================================
st.divider()
st.subheader("⏱️ Délai Moyen d'Ouverture des Devis")

if 'Délai devis ouverture' in df_filtered_base.columns:
    df_delai = df_filtered_base.copy()
    df_delai['Délai devis ouverture'] = pd.to_numeric(df_delai['Délai devis ouverture'], errors='coerce')
    df_delai = df_delai.dropna(subset=['Délai devis ouverture'])

    if not df_delai.empty:
        if vue_annuelle:
            df_delai_groupe = (
                df_delai.groupby('Année Devis')['Délai devis ouverture']
                .mean()
                .reset_index()
            )
            df_delai_groupe.columns = ['Période', 'Délai moyen']
            df_delai_groupe['Période'] = df_delai_groupe['Période'].astype(str)
            x_label = "Année"
        else:
            df_delai_groupe = (
                df_delai.groupby('Mois_Nom')['Délai devis ouverture']
                .mean()
                .reindex(mois_selectionnes_bornes, fill_value=None)
                .reset_index()
            )
            df_delai_groupe.columns = ['Période', 'Délai moyen']
            x_label = "Mois"

        col_delai_graph, col_delai_table = st.columns([6, 4])

        with col_delai_graph:
            fig_delai = go.Figure()
            fig_delai.add_trace(go.Bar(
                x=df_delai_groupe['Période'],
                y=df_delai_groupe['Délai moyen'].round(1),
                marker_color='#6366f1',
                text=df_delai_groupe['Délai moyen'].round(1).apply(
                    lambda v: f"{v:.1f} j".replace('.', ',') if pd.notna(v) else ""
                ),
                textposition='outside',
                hovertemplate="<b>%{x}</b><br>Délai moyen : %{y:.1f} jours<extra></extra>",
            ))
            fig_delai.update_layout(
                xaxis=dict(title=x_label, type='category'),
                yaxis=dict(title="Jours", showgrid=True),
                hovermode="x unified",
                height=360,
                margin=dict(t=20, b=40),
            )
            st.plotly_chart(fig_delai, use_container_width=True, key="fig_delai")

        with col_delai_table:
            st.markdown("**📋 Détail des délais d'ouverture**")
            cols_detail = ["DEVIS N°", "Nom Client", "Année Devis", "Mois_Nom", "Délai devis ouverture"]
            cols_present = [c for c in cols_detail if c in df_delai.columns]
            df_detail_delai = df_delai[cols_present].copy()
            df_detail_delai = df_detail_delai.rename(columns={
                "DEVIS N°": "N° Devis",
                "Année Devis": "Année",
                "Mois_Nom": "Mois",
                "Délai devis ouverture": "Délai (j)"
            })
            df_detail_delai["Délai (j)"] = df_detail_delai["Délai (j)"].round(1)
            df_detail_delai = df_detail_delai.sort_values("Délai (j)", ascending=False).reset_index(drop=True)
            st.dataframe(df_detail_delai, use_container_width=True, hide_index=True, height=350)
    else:
        st.info("Aucune donnée de délai disponible pour les filtres sélectionnés.")
else:
    st.info("Aucun délai calculable (colonne 'Date ouverture dossier fab' absente ou vide).")


# ========================================================
# SECTION : ANALYSE DES TAUX DE MARGE
# ========================================================
st.divider()
st.subheader("📈 Analyse des Taux de Marge")

def add_trendlines_to_fig(fig, df_source, x_col, y_col, is_log=False):
    df_sub = df_source.copy()

    def calc_trend(df_t):
        df_tmp = df_t[[x_col, y_col]].replace([np.inf, -np.inf], np.nan).dropna()
        if len(df_tmp) < 2: return None
        x = df_tmp[x_col]
        y = df_tmp[y_col]
        
        if is_log:
            valid_log = y > 0
            x = x[valid_log]
            y = y[valid_log]

        if pd.api.types.is_datetime64_any_dtype(x):
            x_num = x.map(lambda d: d.toordinal())
        else:
            x_num = pd.to_numeric(x, errors='coerce')
            valid = x_num.notna()
            x = x[valid]
            y = y[valid]
            x_num = x_num[valid]

        if len(x_num) < 2: return None

        if is_log:
            y_fit = np.log10(y)
        else:
            y_fit = y
            
        m, b = np.polyfit(x_num, y_fit, 1)
        y_pred_fit = m * x_num + b
        
        if is_log:
            y_pred = 10 ** y_pred_fit
        else:
            y_pred = y_pred_fit
            
        ss_res = np.sum((y_fit - y_pred_fit)**2)
        ss_tot = np.sum((y_fit - np.mean(y_fit))**2)
        r2 = 1 - (ss_res / ss_tot) if ss_tot != 0 else 0

        sort_idx = np.argsort(x_num.values)
        return x.iloc[sort_idx], np.array(y_pred)[sort_idx], m, b, r2

    def format_eq(m, b, r2):
        m_s = f"{m:.2e}" if abs(m) < 0.001 else f"{m:.4f}"
        b_s = f"{b:.2e}" if abs(b) < 0.001 else f"{b:.2f}"
        sign = "+" if b >= 0 else ""
        if is_log:
            return f"log10(y) = {m_s}x {sign} {b_s} | R² = {r2:.2f}"
        else:
            return f"y = {m_s}x {sign} {b_s} | R² = {r2:.2f}"

    res_g = calc_trend(df_sub)
    if res_g:
        xg, yg, m, b, r2 = res_g
        fig.add_trace(go.Scatter(x=xg, y=yg, mode='lines', name=f"Tendance Globale | {format_eq(m,b,r2)}",
                                 line=dict(color='#64748b', width=2, dash='dash'), hoverinfo='skip'))

    res_s = calc_trend(df_sub[df_sub['Statut'] == "Signé ✅"])
    if res_s:
        xs, ys, m, b, r2 = res_s
        fig.add_trace(go.Scatter(x=xs, y=ys, mode='lines', name=f"Tendance Signé | {format_eq(m,b,r2)}",
                                 line=dict(color='#15803d', width=2, dash='dot'), hoverinfo='skip'))

    res_r = calc_trend(df_sub[df_sub['Statut'] == "Non signé ❌"])
    if res_r:
        xr, yr, m, b, r2 = res_r
        fig.add_trace(go.Scatter(x=xr, y=yr, mode='lines', name=f"Tendance Refusé | {format_eq(m,b,r2)}",
                                 line=dict(color='#b91c1c', width=2, dash='dot'), hoverinfo='skip'))


if 'Taux de marge' in df_filtered_base.columns:
    df_marge = df_filtered_base.dropna(subset=['Taux de marge']).copy()
    
    mask_signe_marge = df_marge["Signé?"].astype(str).str.upper().str.strip() == "O"
    mask_refuse_marge = df_marge["Signé?"].astype(str).str.upper().str.strip() != "O"

    if not df_marge.empty:

        # -------------------------------------------------------
        # GRAPHIQUE 1 : DIAGRAMME BÂTONS — TAUX MOYEN PAR PÉRIODE
        # -------------------------------------------------------
        st.markdown("#### 🔵 Taux de marge moyen par maille de temps")
        st.caption("Comparaison du taux moyen global, des devis signés, et des devis refusés.")

        if vue_annuelle:
            taux_global_yr = (
                df_marge.groupby('Année Devis')['Taux de marge']
                .mean()
                .reset_index()
            )
            taux_global_yr.columns = ['Période', 'Taux moyen global']
            taux_global_yr['Période'] = taux_global_yr['Période'].astype(str)

            taux_signe_yr = (
                df_marge[mask_signe_marge].groupby('Année Devis')['Taux de marge']
                .mean()
                .reset_index()
            )
            taux_signe_yr.columns = ['Période', 'Taux moyen signé']
            taux_signe_yr['Période'] = taux_signe_yr['Période'].astype(str)

            taux_refuse_yr = (
                df_marge[mask_refuse_marge].groupby('Année Devis')['Taux de marge']
                .mean()
                .reset_index()
            )
            taux_refuse_yr.columns = ['Période', 'Taux moyen refusé']
            taux_refuse_yr['Période'] = taux_refuse_yr['Période'].astype(str)

            df_taux = pd.merge(taux_global_yr, taux_signe_yr, on='Période', how='outer').fillna(0)
            df_taux = pd.merge(df_taux, taux_refuse_yr, on='Période', how='outer').fillna(0)
            x_label_marge = "Année"

        else:
            taux_global_ms = []
            taux_signe_ms = []
            taux_refuse_ms = []
            for annee_sel in annees_selectionnees:
                df_marge_annee = df_marge[df_marge['Année Devis'] == annee_sel]
                mask_s = df_marge_annee["Signé?"].astype(str).str.upper().str.strip() == "O"
                mask_r = df_marge_annee["Signé?"].astype(str).str.upper().str.strip() != "O"

                g = df_marge_annee.groupby('Mois_Nom')['Taux de marge'].mean().reset_index()
                g.columns = ['Période', 'Taux moyen global']
                g['Année'] = str(annee_sel)

                s = df_marge_annee[mask_s].groupby('Mois_Nom')['Taux de marge'].mean().reset_index()
                s.columns = ['Période', 'Taux moyen signé']
                s['Année'] = str(annee_sel)

                r = df_marge_annee[mask_r].groupby('Mois_Nom')['Taux de marge'].mean().reset_index()
                r.columns = ['Période', 'Taux moyen refusé']
                r['Année'] = str(annee_sel)

                taux_global_ms.append(g)
                taux_signe_ms.append(s)
                taux_refuse_ms.append(r)

            if taux_global_ms:
                df_g_all = pd.concat(taux_global_ms, ignore_index=True)
                df_s_all = pd.concat(taux_signe_ms, ignore_index=True)
                df_r_all = pd.concat(taux_refuse_ms, ignore_index=True)
                
                df_taux = pd.merge(df_g_all, df_s_all, on=['Période', 'Année'], how='outer')
                df_taux = pd.merge(df_taux, df_r_all, on=['Période', 'Année'], how='outer')
                df_taux = df_taux[df_taux['Période'].isin(mois_selectionnes_bornes)]
                
                ordre_mois = {m: i for i, m in enumerate(liste_mois_noms)}
                df_taux['_ordre'] = df_taux['Période'].map(ordre_mois)
                df_taux = df_taux.sort_values(['Année', '_ordre']).drop(columns='_ordre').reset_index(drop=True).fillna(0)
                if len(annees_selectionnees) > 1:
                    df_taux['Période_affichage'] = df_taux['Période'] + ' ' + df_taux['Année']
                else:
                    df_taux['Période_affichage'] = df_taux['Période']
            else:
                df_taux = pd.DataFrame()
            x_label_marge = "Mois"

        if not df_taux.empty:
            col_bar_marge, _ = st.columns([8, 2])
            with col_bar_marge:
                fig_marge_bar = go.Figure()

                periode_col = 'Période_affichage' if 'Période_affichage' in df_taux.columns else 'Période'

                fig_marge_bar.add_trace(go.Bar(
                    x=df_taux[periode_col],
                    y=df_taux['Taux moyen global'].round(2),
                    name="Taux moyen global",
                    marker_color='#94a3b8',
                    text=df_taux['Taux moyen global'].round(2).apply(
                        lambda v: f"{v:.2f}%".replace('.', ',') if pd.notna(v) and v > 0 else ""
                    ),
                    textposition='outside',
                    hovertemplate="<b>%{x}</b><br>Taux moyen global : %{y:.2f}%<extra></extra>",
                ))

                fig_marge_bar.add_trace(go.Bar(
                    x=df_taux[periode_col],
                    y=df_taux['Taux moyen signé'].round(2),
                    name="Taux moyen signé",
                    marker_color='#22c55e',
                    text=df_taux['Taux moyen signé'].round(2).apply(
                        lambda v: f"{v:.2f}%".replace('.', ',') if pd.notna(v) and v > 0 else ""
                    ),
                    textposition='outside',
                    hovertemplate="<b>%{x}</b><br>Taux moyen signé : %{y:.2f}%<extra></extra>",
                ))

                fig_marge_bar.add_trace(go.Bar(
                    x=df_taux[periode_col],
                    y=df_taux['Taux moyen refusé'].round(2),
                    name="Taux moyen refusé",
                    marker_color='#ef4444',
                    text=df_taux['Taux moyen refusé'].round(2).apply(
                        lambda v: f"{v:.2f}%".replace('.', ',') if pd.notna(v) and v > 0 else ""
                    ),
                    textposition='outside',
                    hovertemplate="<b>%{x}</b><br>Taux moyen refusé : %{y:.2f}%<extra></extra>",
                ))

                fig_marge_bar.update_layout(
                    barmode='group',
                    xaxis=dict(title=x_label_marge, type='category'),
                    yaxis=dict(title="Taux de marge (%)", showgrid=True, range=[0, max(
                        df_taux[['Taux moyen global', 'Taux moyen signé', 'Taux moyen refusé']].max(skipna=True).max() * 1.25, 10
                    )]),
                    legend=dict(orientation="h", y=1.08, x=0),
                    hovermode="x unified",
                    height=380,
                    margin=dict(t=30, b=40),
                )
                st.plotly_chart(fig_marge_bar, use_container_width=True, key="fig_marge_bar")
        else:
            st.info("Aucune donnée de taux de marge disponible pour les filtres sélectionnés.")

        st.write("---")

        # -------------------------------------------------------
        # GRAPHIQUE NOUVEAU : NUAGE DE POINTS — TAUX DE MARGE × TEMPS
        # -------------------------------------------------------
        st.markdown("#### 🔵 Détail des taux de marge dans le temps")
        st.caption("Axe Y : Taux de marge (%) · Axe X : date · Taille : prix total du devis (€) · Couleur : statut signé/non signé")

        df_marge_scatter = df_filtered_base.dropna(subset=['Taux de marge', 'Dates_Propres']).copy()

        if not df_marge_scatter.empty:
            df_marge_scatter['Statut'] = df_marge_scatter["Signé?"].apply(
                lambda x: "Signé ✅" if str(x).upper().strip() == "O" else "Non signé ❌"
            )
            df_marge_scatter['Prix_plot'] = df_marge_scatter['Prix total'].fillna(0).clip(lower=1)

            if vue_annuelle:
                df_marge_scatter['X_val'] = df_marge_scatter['Année Devis'].astype(str)
                x_axis_type_marge = 'category'
                x_title_marge = "Année"
            else:
                df_marge_scatter['X_val'] = df_marge_scatter['Dates_Propres']
                x_axis_type_marge = 'date'
                x_title_marge = "Date"

            color_map_marge = {"Signé ✅": "#22c55e", "Non signé ❌": "#ef4444"}

            fig_scatter_marge = px.scatter(
                df_marge_scatter,
                x='X_val',
                y='Taux de marge',
                size='Prix_plot',
                color='Statut',
                color_discrete_map=color_map_marge,
                hover_data={
                    'Nom Client': True,
                    'DEVIS N°': True if 'DEVIS N°' in df_marge_scatter.columns else False,
                    'Prix_plot': False,
                    'Prix total': ':.2f',
                    'Taux de marge': ':.2f',
                    'X_val': False,
                },
                size_max=40,
                opacity=0.7,
                labels={
                    'X_val': x_title_marge,
                    'Taux de marge': 'Taux de marge (%)',
                    'Statut': 'Statut',
                    'Prix total': 'Prix total (€)',
                },
            )

            if show_trends:
                add_trendlines_to_fig(fig_scatter_marge, df_marge_scatter, 'X_val', 'Taux de marge', is_log=False)

            fig_scatter_marge.update_layout(
                xaxis=dict(title=x_title_marge, type=x_axis_type_marge),
                yaxis=dict(
                    title="Taux de marge (%)",
                    showgrid=True,
                ),
                legend=dict(orientation="h", y=1.18, x=0, title_text=""),
                hovermode="closest",
                height=520,
                margin=dict(t=30, b=40),
            )
            st.plotly_chart(fig_scatter_marge, use_container_width=True, key="fig_scatter_marge_temps")
        else:
            st.info("Aucune donnée de date/marge valide pour ce graphique.")

        st.write("---")

        # -------------------------------------------------------
        # GRAPHIQUE 2 : NUAGE DE POINTS — PRIX UNITAIRE × TEMPS
        # -------------------------------------------------------
        st.markdown("#### 🔵 Prix unitaire des devis dans le temps")
        st.caption("Axe Y : prix unitaire (Échelle Logarithmique) · Axe X : date · Taille : prix total du devis (€) · Couleur : statut signé/non signé")

        if 'Prix unitaire' in df_filtered_base.columns:
            df_scatter = df_filtered_base.dropna(subset=['Prix unitaire', 'Dates_Propres']).copy()
            df_scatter = df_scatter[df_scatter['Prix unitaire'] > 0]

            if not df_scatter.empty:
                df_scatter['Statut'] = df_scatter["Signé?"].apply(
                    lambda x: "Signé ✅" if str(x).upper().strip() == "O" else "Non signé ❌"
                )
                df_scatter['Prix_plot'] = df_scatter['Prix total'].fillna(0).clip(lower=1)

                if vue_annuelle:
                    df_scatter['X_val'] = df_scatter['Année Devis'].astype(str)
                    x_axis_type = 'category'
                    x_title = "Année"
                else:
                    df_scatter['X_val'] = df_scatter['Dates_Propres']
                    x_axis_type = 'date'
                    x_title = "Date"

                color_map = {"Signé ✅": "#22c55e", "Non signé ❌": "#ef4444"}

                fig_scatter = px.scatter(
                    df_scatter,
                    x='X_val',
                    y='Prix unitaire',
                    size='Prix_plot',
                    color='Statut',
                    color_discrete_map=color_map,
                    hover_data={
                        'Nom Client': True,
                        'DEVIS N°': True if 'DEVIS N°' in df_scatter.columns else False,
                        'Prix_plot': False,
                        'Prix total': ':.2f',
                        'Prix unitaire': ':.2f',
                        'X_val': False,
                    },
                    size_max=40,
                    opacity=0.7,
                    labels={
                        'X_val': x_title,
                        'Prix unitaire': 'Prix unitaire (€)',
                        'Statut': 'Statut',
                        'Prix total': 'Prix total (€)',
                    },
                )

                if show_trends:
                    add_trendlines_to_fig(fig_scatter, df_scatter, 'X_val', 'Prix unitaire', is_log=True)

                fig_scatter.update_layout(
                    xaxis=dict(title=x_title, type=x_axis_type),
                    yaxis=dict(
                        title="Prix unitaire (€) — Échelle Log",
                        type='log',
                        showgrid=True,
                    ),
                    legend=dict(orientation="h", y=1.18, x=0, title_text=""),
                    hovermode="closest",
                    height=520,
                    margin=dict(t=30, b=40),
                )
                st.plotly_chart(fig_scatter, use_container_width=True, key="fig_scatter_prix")
            else:
                st.info("Aucune donnée de prix unitaire disponible pour les filtres sélectionnés.")
        else:
            st.info("La colonne 'Prix unitaire' n'a pas pu être calculée (vérifiez 'Prix total' et 'Nb exemplaires').")

    else:
        st.info("Aucune donnée de taux de marge disponible pour les filtres sélectionnés.")
else:
    st.info("Colonne 'Taux de marge' absente des données.")


# ========================================================
# PIED DE PAGE : RÉSUMÉ DES FILTRES ACTIFS
# ========================================================
def construire_resume_filtres():
    parties = []
    maille = "Annuelle" if vue_annuelle else "Mensuelle"
    parties.append(f"BU : {bu_name}")
    parties.append(f"Maille : {maille}")
    if annees_selectionnees:
        parties.append(f"Années : {', '.join(str(a) for a in sorted(annees_selectionnees))}")
    if not vue_annuelle:
        if start_month == end_month:
            parties.append(f"Mois : {start_month}")
        else:
            parties.append(f"Mois : {start_month} → {end_month}")
            
    # Ajout du filtre de montant dans le récapitulatif
    parties.append(f"Montant : {montant_min:,.0f}€ à {montant_max:,.0f}€".replace(',', ' '))
    
    if selected_abc and len(selected_abc) < 4:
        parties.append(f"Catégorie ABC : {', '.join(selected_abc)}")
    else:
        parties.append("Catégorie ABC : Toutes")
    if selected_commercial:
        parties.append(f"Commercial : {', '.join(selected_commercial)}")
    if search_client:
        parties.append(f"Client : « {search_client} »")
    if search_devis:
        parties.append(f"Devis : « {search_devis} »")
    parties.append(f"{len(df_filtered_base)} ligne(s) affichée(s)")
    now = datetime.datetime.now().strftime("%d/%m/%Y à %H:%M")
    parties.append(f"Édité le {now}")
    return "  ·  ".join(parties)

st.markdown("---")
st.markdown(
    f"<p style='font-size:0.75em; color:#94a3b8; text-align:center; margin-top:4px;'>"
    f"🖨️ {construire_resume_filtres()}"
    f"</p>",
    unsafe_allow_html=True
)
