import datetime
import io
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

NOM_SITE = "Atelier Villeneuve-la-Garenne - Intranet Sécurité & Production"
JOURS_SEMAINE = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]

# Constantes de colonnes pour les fichiers Excel de l'atelier (export WO EBO)
ACTOR_COL = "Assigné à l'acteur Libellé"
CAT_ABC_COL = "smovCategory"
CAT_MODEL_COL = "Modèle d'équipement"
MAT_COL = "Matériel"
CHARGE_COL = "Charge prévue"
ETAT_COL = "Etat"

# 5 Agents exemples modifiables au lieu d'une liste figée en dur
AGENTS_INITIAUX = [
    {"nom": "ALI ABDALLAH", "tel": "14", "badge": "1706", "contrat": "CDI", "debut": "2022-03-15", "fin": "", "dotation": "Complet"},
    {"nom": "ERIC DUPONT", "tel": "05", "badge": "830", "contrat": "CDI", "debut": "2021-06-12", "fin": "", "dotation": "Complet"},
    {"nom": "THOMAS MARTIN", "tel": "09", "badge": "1069", "contrat": "CDD", "debut": "2024-01-01", "fin": "2026-12-31", "dotation": "Complet"},
    {"nom": "FARID SALHI", "tel": "01", "badge": "858", "contrat": "CDI", "debut": "2018-09-01", "fin": "", "dotation": "Complet"},
    {"nom": "BOUMEDIENNE LAZAAR", "tel": "02", "badge": "837", "contrat": "CDI", "debut": "2020-05-01", "fin": "", "dotation": "Complet"},
]

st.set_page_config(page_title=NOM_SITE, page_icon="🚲", layout="wide")

# --- FONCTIONS UTILITAIRES DE CALCUL ET STYLE ---
def parse_time_to_hours(val):
    if pd.isna(val):
        return 0.0
    val_str = str(val).strip()
    match = re.match(r'^(\d+):(\d+):(\d+)$', val_str)
    if match:
        h, m, s = map(int, match.groups())
        return h + m / 60.0 + s / 3600.0
    try:
        return float(val_str)
    except:
        return 0.0

def format_hours(hours_float):
    total_seconds = int(round(hours_float * 3600))
    h = total_seconds // 3600
    m = (total_seconds % 3600) // 60
    return f"{h}h {m:02d}m"

def calculer_prime(velos):
    if velos < 6.0:
        return 0.0
    elif velos >= 10.0:
        return 155.0
    else:
        return round(50.0 + (velos - 6.0) * 26.25, 2)

def lire_fichier(uploaded_file):
    if uploaded_file.name.endswith(".csv"):
        return pd.read_csv(uploaded_file, sep=None, engine="python")
    else:
        return pd.read_excel(uploaded_file)

def color_eligibility(val):
    if "Droit de prime" in str(val):
        return "background-color: #d1fae5; color: #065f46; font-weight: bold;"
    return "background-color: #fee2e2; color: #991b1b; font-weight: bold;"

def to_excel(df_to_export):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_to_export.to_excel(writer, index=False, sheet_name='Analyse_Atelier')
    return output.getvalue()

# --- INITIALISATION SESSION STATE ---
if "users_db" not in st.session_state:
    st.session_state["users_db"] = {
        "admin": {"password": "AdminPassword123!", "role": "Admin", "nom": "Administrateur VLG"}
    }
if "authenticated_user" not in st.session_state:
    st.session_state["authenticated_user"] = None
if "agents_db" not in st.session_state:
    st.session_state["agents_db"] = {a["nom"]: a.copy() for a in AGENTS_INITIAUX}
if "planning_db" not in st.session_state:
    st.session_state["planning_db"] = {}
if "securite_db" not in st.session_state:
    st.session_state["securite_db"] = {"fiches": [], "accidents": [], "minutes": []}
if "leve_velos_db" not in st.session_state:
    st.session_state["leve_velos_db"] = [
        {"id": i, "nom": f"Lève-Vélo n°{i}", "statut": "Opérationnel", "anomalie": "", "intervention": "Aucune"}
        for i in range(1, 31)
    ]
if "souvenirs_db" not in st.session_state:
    st.session_state["souvenirs_db"] = []

# --- SÉLECTEUR DE THÈMES VISUELS (DONT THÈMES REPOSANTS) ---
st.sidebar.markdown("### 🎨 Thème Visuel Reposant (Journée 7h)")
theme_option = st.sidebar.selectbox(
    "Choisir le style d'affichage :",
    [
        "🌙 Nuit Sombre & Ambre (Anti-fatigue)",
        "🌲 Vert Sauge & Faux-Noir (Nature)",
        "🪨 Gris Papier & Encre (Clair anti-éblouissement)",
        "Gris / Noir léger (Écriture Blanche)",
        "Bleu léger (Écriture Noire)",
        "Rose (Écriture Blanche & Boutons)"
    ]
)

if "🌙" in theme_option:
    css_theme = """
    .stApp { background-color: #121826 !important; color: #ffffff !important; }
    p, span, label, div, h1, h2, h3, h4, h5, h6 { color: #ffffff !important; }
    .vlg-banner { background: #1e293b; padding: 20px; border-radius: 12px; border-left: 6px solid #f59e0b; }
    .card-menu { background-color: #1e293b !important; padding: 18px; border-radius: 10px; border: 1px solid #334155; }
    """
elif "🌲" in theme_option:
    css_theme = """
    .stApp { background-color: #0f1715 !important; color: #f1f5f9 !important; }
    p, span, label, div, h1, h2, h3, h4, h5, h6 { color: #f1f5f9 !important; }
    .vlg-banner { background: #1c2826; padding: 20px; border-radius: 12px; border-left: 6px solid #34d399; }
    .card-menu { background-color: #1c2826 !important; padding: 18px; border-radius: 10px; border: 1px solid #2d3a37; }
    """
elif "🪨" in theme_option:
    css_theme = """
    .stApp { background-color: #f3f4f6 !important; color: #111827 !important; }
    p, span, label, div, h1, h2, h3, h4, h5, h6 { color: #111827 !important; }
    .vlg-banner { background: #e5e7eb; padding: 20px; border-radius: 12px; border-left: 6px solid #374151; }
    .card-menu { background-color: #ffffff !important; padding: 18px; border-radius: 10px; border: 1px solid #d1d5db; }
    """
elif "Gris / Noir" in theme_option:
    css_theme = """
    .stApp { background-color: #1e293b !important; color: #f8fafc !important; }
    p, span, label, div, h1, h2, h3, h4, h5, h6 { color: #f8fafc !important; }
    .vlg-banner { background: #334155; padding: 20px; border-radius: 12px; border-left: 6px solid #64748b; }
    .card-menu { background-color: #0f172a !important; padding: 18px; border-radius: 10px; border: 1px solid #475569; }
    """
elif "Bleu léger" in theme_option:
    css_theme = """
    .stApp { background-color: #e0f2fe !important; color: #0f172a !important; }
    p, span, label, div, h1, h2, h3, h4, h5, h6 { color: #0f172a !important; }
    .vlg-banner { background: #bae6fd; padding: 20px; border-radius: 12px; border-left: 6px solid #0284c7; }
    .card-menu { background-color: #ffffff !important; padding: 18px; border-radius: 10px; border: 1px solid #7dd3fc; }
    """
else:
    css_theme = """
    .stApp { background-color: #fdf2f8 !important; color: #831843 !important; }
    p, span, label, div, h1, h2, h3, h4, h5, h6 { color: #831843 !important; }
    .vlg-banner { background: #fbcfe8; padding: 20px; border-radius: 12px; border-left: 6px solid #db2777; }
    .card-menu { background-color: #ffffff !important; padding: 18px; border-radius: 10px; border: 1px solid #f472b6; }
    .stDownloadButton button { background-color: #db2777 !important; color: #ffffff !important; border-radius: 8px; border: 2px solid #be185d; font-weight: bold; }
    """

st.markdown(f"<style>{css_theme}</style>", unsafe_allow_html=True)

# --- CONNEXION ---
if st.session_state["authenticated_user"] is None:
    st.markdown('<div class="vlg-banner"><h1>🚲 Intranet Atelier Villeneuve-la-Garenne</h1><p>Connexion Sécurisée</p></div>', unsafe_allow_html=True)
    with st.form("login_form"):
        username = st.text_input("Identifiant", placeholder="admin")
        password = st.text_input("Mot de passe", type="password")
        if st.form_submit_button("Se connecter", type="primary"):
            users = st.session_state["users_db"]
            if username in users and users[username]["password"] == password:
                st.session_state["authenticated_user"] = username
                st.rerun()
            st.error("Identifiant ou mot de passe incorrect.")
    st.stop()

st.sidebar.markdown(f"### 🚲 INTRANET VLG")
st.sidebar.write(f"Connecté : **{st.session_state['authenticated_user']}**")
if st.sidebar.button("Déconnexion"):
    st.session_state["authenticated_user"] = None
    st.rerun()

st.sidebar.divider()
menu_option = st.sidebar.radio(
    "Navigation",
    [
        "🏠 Accueil & Menu Principal",
        "📈 Analyse Quotidienne (Excel/CSV)",
        "👷 Gestion Salariés & Badges",
        "📅 Planning & Feuille Présence A4",
        "🚨 Sécurité & Lève-Vélos",
        "📸 Album Souvenirs Atelier"
    ]
)

# ==========================================
# 0. ACCUEIL
# ==========================================
if menu_option == "🏠 Accueil & Menu Principal":
    st.markdown('<div class="vlg-banner"><h1>🚲 Intranet de l\'Atelier VLG</h1><p>Bienvenue sur votre plateforme de gestion quotidienne.</p></div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="card-menu"><h3>📈 Analyse Quotidienne</h3><p>Import Excel/CSV, calcul des vélos, catégories A/B/C, charges et primes.</p></div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="card-menu"><h3>📅 Feuille Présence A4</h3><p>Badges, téléphones et cases de signature 1x3 cm prêtes à imprimer.</p></div>', unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="card-menu"><h3>🚨 Sécurité & HSE</h3><p>Suivi des 30 lève-vélos et enregistrement des AT.</p></div>', unsafe_allow_html=True)

# ==========================================
# 1. SECTION : ANALYSE DES INTERVENTIONS & PRIMES (JOURNALIER)
# ==========================================
elif menu_option == "📈 Analyse Quotidienne (Excel/CSV)":
    st.header("📈 Analyse Quotidienne des Interventions & Primes")
    uploaded_file = st.file_uploader("📂 Injecter le fichier Excel / CSV du jour", type=["xls", "xlsx", "csv"])

    if uploaded_file is not None:
        try:
            df = lire_fichier(uploaded_file)
            st.success(f"Fichier **{uploaded_file.name}** injecté avec succès ! ({len(df)} interventions enregistrées)")

            actor_col, cat_abc_col, cat_model_col = ACTOR_COL, CAT_ABC_COL, CAT_MODEL_COL
            mat_col, charge_col = MAT_COL, CHARGE_COL

            df['Charge_heures'] = df[charge_col].apply(parse_time_to_hours) if charge_col in df.columns else 0.0

            total_acteurs = df[actor_col].nunique() if actor_col in df.columns else 0
            total_interventions = len(df)
            total_materiels = df[mat_col].nunique() if mat_col in df.columns else 0
            moyenne_velos_equipe = total_interventions / total_acteurs if total_acteurs > 0 else 0
            total_charge = df['Charge_heures'].sum()

            col1, col2, col3, col4, col5 = st.columns(5)
            col1.metric("Nombre de Mécaniciens", total_acteurs)
            col2.metric("Total Vélos Sortis", total_interventions)
            col3.metric("Matériels Uniques", total_materiels)
            col4.metric("Moyenne Équipe", f"{moyenne_velos_equipe:.2f} vélos")
            col5.metric("Charge Totale Prévue", format_hours(total_charge))

            st.divider()

            tab1, tab2, tab3, tab4 = st.tabs([
                "🚴‍♂️ Vélos Sortis & Primes",
                "🏷️ Catégories A / B / C",
                "📌 Synthèse & Charges",
                "🚲 Modèles (VAE / VLS)"
            ])

            with tab1:
                st.subheader("Analyse des Vélos Sortis, Statut d'Éligibilité et Primes")
                st.info("💡 **Condition de prime :** À partir de **6 vélos sortis** (50,00 €) jusqu'à **10 vélos** (155,00 € max).")

                agent_bikes = df.groupby(actor_col).agg(
                    Velos_Sortis=(mat_col, 'count')
                ).reset_index()

                agent_bikes['Éligibilité'] = agent_bikes['Velos_Sortis'].apply(
                    lambda x: "🟢 Droit de prime" if x >= 6.0 else "🔴 Pas de prime"
                )
                agent_bikes['Montant Prime (€)'] = agent_bikes['Velos_Sortis'].apply(calculer_prime)
                agent_bikes = agent_bikes.sort_values(by='Velos_Sortis', ascending=False)

                styled_df = agent_bikes.style.applymap(color_eligibility, subset=['Éligibilité']).format({
                    'Montant Prime (€)': '{:.2f} €'
                })
                st.dataframe(styled_df, use_container_width=True)

                excel_data = to_excel(agent_bikes)
                st.download_button(
                    label="📥 Télécharger ce rapport avec Primes (Excel)",
                    data=excel_data,
                    file_name=f"rapport_primes_mecaniciens_{datetime.date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                st.bar_chart(agent_bikes.set_index(actor_col)['Velos_Sortis'])

            with tab2:
                st.subheader("Analyse des Catégories A, B, C par acteur")
                if cat_abc_col in df.columns:
                    ct_abc = pd.crosstab(df[actor_col], df[cat_abc_col])
                    for col in ['A', 'B', 'C']:
                        if col not in ct_abc.columns:
                            ct_abc[col] = 0
                    ct_abc = ct_abc[['A', 'B', 'C']]
                    ct_abc['Total Matériel'] = ct_abc.sum(axis=1)

                    ct_abc['% Cat A'] = (ct_abc['A'] / ct_abc['Total Matériel'] * 100).round(1).astype(str) + '%'
                    ct_abc['% Cat B'] = (ct_abc['B'] / ct_abc['Total Matériel'] * 100).round(1).astype(str) + '%'
                    ct_abc['% Cat C'] = (ct_abc['C'] / ct_abc['Total Matériel'] * 100).round(1).astype(str) + '%'

                    res_abc = ct_abc.reset_index()[[actor_col, 'A', '% Cat A', 'B', '% Cat B', 'C', '% Cat C', 'Total Matériel']]
                    st.dataframe(res_abc, use_container_width=True)
                else:
                    st.warning("Colonne de catégorie A/B/C introuvable dans le fichier.")

            with tab3:
                st.subheader("Nombre de matériels et charge prévue par acteur")
                summary = df.groupby(actor_col).agg(
                    Nombre_Materiel=(mat_col, 'count'),
                    Total_Charge_Heures=('Charge_heures', 'sum')
                ).reset_index()
                summary['Charge Prévue'] = summary['Total_Charge_Heures'].apply(format_hours)
                summary = summary[[actor_col, 'Nombre_Materiel', 'Charge Prévue']]
                st.dataframe(summary, use_container_width=True)

            with tab4:
                st.subheader("Répartition des modèles d'équipement (VAE / VLS)")
                if cat_model_col in df.columns:
                    ct_model = pd.crosstab(df[actor_col], df[cat_model_col]).reset_index()
                    st.dataframe(ct_model, use_container_width=True)

        except Exception as e:
            st.error(f"Erreur lors du traitement du fichier : {e}")

# ==========================================
# 2. GESTION DYNAMIQUE DES SALARIÉS
# ==========================================
elif menu_option == "👷 Gestion Salariés & Badges":
    st.header("👷 Modification & Gestion des Agents")
    st.write("Vous pouvez ajouter, modifier ou renommer les mécaniciens à tout moment.")

    with st.form("form_add_agent"):
        col_a1, col_a2, col_a3 = st.columns(3)
        with col_a1:
            n_nom = st.text_input("Nom & Prénom Agent")
        with col_a2:
            n_tel = st.text_input("N° Téléphone")
        with col_a3:
            n_badge = st.text_input("N° Badge")
        if st.form_submit_button("➕ Enregistrer / Modifier Agent") and n_nom:
            nom_c = n_nom.strip().upper()
            st.session_state["agents_db"][nom_c] = {
                "nom": nom_c, "tel": n_tel, "badge": n_badge, "contrat": "CDI", "debut": str(datetime.date.today()), "fin": "", "dotation": "Complet"
            }
            st.success(f"Agent {nom_c} mis à jour avec succès !")
            st.rerun()

    st.subheader("📋 Liste Actuelle des Agents Modifiables")
    df_ag = pd.DataFrame(list(st.session_state["agents_db"].values()))
    st.dataframe(df_ag[["nom", "tel", "badge", "contrat"]], use_container_width=True, hide_index=True)

# ==========================================
# 3. PLANNING & FEUILLE A4 PORTRAIT
# ==========================================
elif menu_option == "📅 Planning & Feuille Présence A4":
    st.header("📄 Feuille de Présence A4 à Imprimer (100% Encadrée)")
    semaine = st.selectbox("Semaine", list(range(1, 53)))
    jour = st.selectbox("Jour de la semaine", JOURS_SEMAINE)

    agents_list = list(st.session_state["agents_db"].values())
    
    if st.button("📥 Générer la Feuille de Présence A4 (Excel encadré)"):
        output_pres = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Presence_A4"

        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0

        black_border_thin = Border(
            left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000')
        )
        black_border_medium = Border(
            left=Side(style='medium', color='000000'), right=Side(style='medium', color='000000'),
            top=Side(style='medium', color='000000'), bottom=Side(style='medium', color='000000')
        )

        ws.merge_cells('A1:E1')
        ws['A1'] = "MOBISERVICE - ATELIER VILLENEUVE-LA-GARENNE"
        ws['A1'].font = Font(name='Calibri', size=14, bold=True, color="1E3A8A")
        ws['A1'].alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells('A2:E2')
        ws['A2'] = f"FEUILLE DE PRÉSENCE & BADGES - {jour.upper()} (Semaine {semaine})"
        ws['A2'].font = Font(name='Calibri', size=12, bold=True)
        ws['A2'].alignment = Alignment(horizontal="center", vertical="center")

        headers = ["NOM & PRÉNOM", "TEL N°", "BADGE N°", "SIGNATURE 07H00", "SIGNATURE 14H00"]
        ws.row_dimensions[4].height = 28

        for col_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col_idx, value=h)
            cell.font = Font(name='Calibri', size=11, bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = black_border_medium

        row_idx = 4
        for ag in agents_list:
            row_idx += 1
            ws.row_dimensions[row_idx].height = 32  # ~1cm hauteur signature

            ws.cell(row=row_idx, column=1, value=ag["nom"]).border = black_border_thin
            ws.cell(row=row_idx, column=2, value=str(ag["tel"])).border = black_border_thin
            ws.cell(row=row_idx, column=3, value=str(ag["badge"])).border = black_border_thin
            ws.cell(row=row_idx, column=4, value="").border = black_border_thin
            ws.cell(row=row_idx, column=5, value="").border = black_border_thin

        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 12
        ws.column_dimensions['D'].width = 22  # ~3cm largeur
        ws.column_dimensions['E'].width = 22  # ~3cm largeur

        wb.save(output_pres)
        output_pres.seek(0)

        st.download_button(
            "Télécharger le Fichier Excel Prêt à Imprimer",
            data=output_pres.getvalue(),
            file_name=f"feuille_presence_S{semaine}_{jour}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ==========================================
# 4. SÉCURITÉ & LÈVE-VÉLOS
# ==========================================
elif menu_option == "🚨 Sécurité & Lève-Vélos":
    st.header("🚨 Suivi Sécurité & 30 Lève-Vélos")
    for lv in st.session_state["leve_velos_db"][:5]:  # Affichage exemple
        st.write(f"📍 **{lv['nom']}** — Statut : {lv['statut']}")

# ==========================================
# 5. ALBUM SOUVENIRS
# ==========================================
elif menu_option == "📸 Album Souvenirs Atelier":
    st.header("📸 Moments Forts & Album Souvenirs")
    st.write("Espace de partage pour les photos et événements de l'atelier VLG.")
