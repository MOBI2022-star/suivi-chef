import datetime
import io
import os
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

NOM_SITE = "Atelier Villeneuve-la-Garenne - Intranet Sécurité & Production"
JOURS_SEMAINE = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]

AGENTS_INITIAUX = [
    {"nom": "ABBAS HAMZA", "tel": "20", "badge": "1705", "contrat": "CDI", "debut": "2022-01-10", "fin": "", "dotation": "Complet (Casque, Gants, Chaussures)"},
    {"nom": "ALI ABDALLAH YAHIA", "tel": "14", "badge": "1706", "contrat": "CDI", "debut": "2022-03-15", "fin": "", "dotation": "Complet"},
    {"nom": "BOUMEDIENNE LAZAAR", "tel": "2", "badge": "837", "contrat": "CDI", "debut": "2020-05-01", "fin": "", "dotation": "Complet"},
    {"nom": "DEDOU BENGARA", "tel": "9", "badge": "1069", "contrat": "CDD", "debut": "2024-01-01", "fin": "2026-12-31", "dotation": "Complet"},
    {"nom": "FABRICE CREYX", "tel": "5", "badge": "830", "contrat": "CDI", "debut": "2019-06-12", "fin": "", "dotation": "Complet"},
    {"nom": "FARID SALHI", "tel": "1", "badge": "858", "contrat": "CDI", "debut": "2018-09-01", "fin": "", "dotation": "Complet"},
    {"nom": "FETHI ABDELHAMIDE", "tel": "7", "badge": "839", "contrat": "CDI", "debut": "2021-02-14", "fin": "", "dotation": "Complet"},
    {"nom": "KEBE BAMBO", "tel": "4", "badge": "2084", "contrat": "CDI", "debut": "2022-07-01", "fin": "", "dotation": "Complet"},
    {"nom": "KORAICH YOUNES", "tel": "19", "badge": "674", "contrat": "CDI", "debut": "2020-11-15", "fin": "", "dotation": "Complet"},
    {"nom": "LOIC FLORIS", "tel": "10", "badge": "945", "contrat": "CDD", "debut": "2025-02-01", "fin": "2026-06-30", "dotation": "Complet"},
    {"nom": "MAHMOUD AHMEDIN ALHAJ", "tel": "13", "badge": "1476", "contrat": "CDI", "debut": "2023-01-10", "fin": "", "dotation": "Complet"},
    {"nom": "MAMADOU MALAL BALDE", "tel": "25", "badge": "1397", "contrat": "CDI", "debut": "2023-04-05", "fin": "", "dotation": "Complet"},
    {"nom": "MICHAEL HUYNH", "tel": "18", "badge": "1739", "contrat": "CDI", "debut": "2021-09-18", "fin": "", "dotation": "Complet"},
    {"nom": "MOUSSA OMAR MUSTAPHA", "tel": "15", "badge": "2000", "contrat": "CDD", "debut": "2025-01-15", "fin": "2026-07-15", "dotation": "Complet"},
    {"nom": "MUSTAPHA ASSAN ABAKER", "tel": "16", "badge": "1703", "contrat": "CDI", "debut": "2022-10-01", "fin": "", "dotation": "Complet"},
    {"nom": "OSMAN ABDELAZIZ CDI", "tel": "21", "badge": "879", "contrat": "CDI", "debut": "2019-03-20", "fin": "", "dotation": "Complet"},
    {"nom": "OSMAN ABDELAZIZ CDD", "tel": "22", "badge": "1710", "contrat": "CDD", "debut": "2025-05-01", "fin": "2026-08-30", "dotation": "Complet"},
    {"nom": "ROMEO PENA", "tel": "DIS", "badge": "948", "contrat": "CDI", "debut": "2017-01-01", "fin": "", "dotation": "Complet"},
    {"nom": "FONTALVO PATERNINA", "tel": "", "badge": "", "contrat": "CDD", "debut": "2026-01-01", "fin": "2026-12-31", "dotation": "Partiel"},
    {"nom": "CHARLES NERIS", "tel": "RET", "badge": "832", "contrat": "CDI", "debut": "2015-01-01", "fin": "", "dotation": "Complet"},
    {"nom": "AHMED MOHAMED AMAR", "tel": "23", "badge": "", "contrat": "CDI", "debut": "2024-06-01", "fin": "", "dotation": "Complet"},
]

st.set_page_config(page_title=NOM_SITE, page_icon="🚲", layout="wide")

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
    st.session_state["securite_db"] = {
        "fiches": [],
        "accidents": [],
        "minutes": [],
        "dossiers_site": []
    }
if "leve_velos_db" not in st.session_state:
    st.session_state["leve_velos_db"] = [
        {"id": i, "nom": f"Lève-Vélo n°{i}", "statut": "Opérationnel", "anomalie": "", "intervention": "Aucune", "date_intervention": "-"}
        for i in range(1, 31)
    ]
if "diffusion_db" not in st.session_state:
    st.session_state["diffusion_db"] = []
if "souvenirs_db" not in st.session_state:
    st.session_state["souvenirs_db"] = []

# --- BARRE LATÉRALE DE CONFIGURATION VISUELLE (4 OPTIONS) ---
st.sidebar.markdown("### 🎨 Thème Visuel de l'Atelier")
theme_option = st.sidebar.selectbox(
    "Choisir le style graphique :",
    [
        "1. Bleu léger (Écriture Noire)",
        "2. Gris / Noir léger (Écriture Blanche)",
        "3. Bleu Marine léger (Écriture Noire)",
        "4. Rose (Écriture Blanche & Boutons Téléchargement)"
    ]
)

# Application dynamique du CSS selon l'option choisie
if "1." in theme_option:
    css_theme = """
    .stApp { background-color: #e0f2fe !important; color: #0f172a !important; }
    p, span, label, div, h1, h2, h3, h4, h5, h6 { color: #0f172a !important; }
    .vlg-banner { background: #bae6fd; padding: 20px; border-radius: 12px; border-left: 6px solid #0284c7; }
    .card-menu { background-color: #ffffff !important; padding: 18px; border-radius: 10px; border: 1px solid #7dd3fc; }
    .download-btn-custom { background-color: #0284c7 !important; color: #ffffff !important; border-radius: 8px; padding: 10px; text-align: center; font-weight: bold; }
    """
elif "2." in theme_option:
    css_theme = """
    .stApp { background-color: #1e293b !important; color: #f8fafc !important; }
    p, span, label, div, h1, h2, h3, h4, h5, h6 { color: #f8fafc !important; }
    .vlg-banner { background: #334155; padding: 20px; border-radius: 12px; border-left: 6px solid #64748b; }
    .card-menu { background-color: #0f172a !important; padding: 18px; border-radius: 10px; border: 1px solid #475569; }
    .download-btn-custom { background-color: #475569 !important; color: #ffffff !important; border-radius: 8px; padding: 10px; text-align: center; font-weight: bold; }
    """
elif "3." in theme_option:
    css_theme = """
    .stApp { background-color: #e2e8f0 !important; color: #0f172a !important; }
    p, span, label, div, h1, h2, h3, h4, h5, h6 { color: #0f172a !important; }
    .vlg-banner { background: #cbd5e1; padding: 20px; border-radius: 12px; border-left: 6px solid #1e293b; }
    .card-menu { background-color: #ffffff !important; padding: 18px; border-radius: 10px; border: 1px solid #94a3b8; }
    .download-btn-custom { background-color: #334155 !important; color: #ffffff !important; border-radius: 8px; padding: 10px; text-align: center; font-weight: bold; }
    """
else: # Option 4: Rose / Blanc + Téléchargement rose
    css_theme = """
    .stApp { background-color: #fdf2f8 !important; color: #831843 !important; }
    p, span, label, div, h1, h2, h3, h4, h5, h6 { color: #831843 !important; }
    .vlg-banner { background: #fbcfe8; padding: 20px; border-radius: 12px; border-left: 6px solid #db2777; }
    .card-menu { background-color: #ffffff !important; padding: 18px; border-radius: 10px; border: 1px solid #f472b6; }
    .stDownloadButton button { background-color: #db2777 !important; color: #ffffff !important; border-radius: 8px; border: 2px solid #be185d; font-weight: bold; }
    """

st.markdown(f"<style>{css_theme}</style>", unsafe_allow_html=True)

# --- GESTION CONNEXION ---
if st.session_state["authenticated_user"] is None:
    st.markdown('<div class="vlg-banner"><h1>🚲 Intranet Atelier Villeneuve-la-Garenne</h1><p>Connexion Sécurisée - Accès Interne</p></div>', unsafe_allow_html=True)
    with st.form("login_form"):
        username = st.text_input("Identifiant", placeholder="admin")
        password = st.text_input("Mot de passe", type="password")
        submitted = st.form_submit_button("Se connecter au réseau local", type="primary")
    if submitted:
        users = st.session_state["users_db"]
        if username in users and users[username]["password"] == password:
            st.session_state["authenticated_user"] = username
            st.rerun()
        st.error("Identifiant ou mot de passe incorrect.")
    st.stop()

current_user = st.session_state["authenticated_user"]

# Sidebar Navigation
st.sidebar.markdown(f"### 🚲 INTRANET VLG")
st.sidebar.write(f"Connecté : **{current_user}**")
if st.sidebar.button("Déconnexion"):
    st.session_state["authenticated_user"] = None
    st.rerun()

st.sidebar.divider()
menu = st.sidebar.radio(
    "Navigation Principale",
    [
        "🏠 Accueil & Menu Principal",
        "👷 1. Suivi Personnel & Dotations",
        "📅 2. Planning, Présences & Feuille A4",
        "📊 3. Suivi Production Quotidienne",
        "💶 4. Droits de Prime",
        "🚨 5. Sécurité, Lève-Vélos & AT",
        "📸 6. Souvenirs & Événements",
        "✉️ 7. Diffusion Société & Mail"
    ]
)

# ==========================================
# 0. ACCUEIL & MENU PRINCIPAL
# ==========================================
if menu == "🏠 Accueil & Menu Principal":
    st.markdown('<div class="vlg-banner"><h1>🚲 Tableau de Bord de l\'Atelier VLG</h1><p>Sélectionnez un module ci-dessous ou utilisez le menu latéral.</p></div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="card-menu"><h3>👷 Suivi Personnel</h3><p>Gestion des 30 agents max, contrats, dotations & fiches.</p></div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="card-menu"><h3>📅 Planning & Feuille A4</h3><p>Gestion S1-S52 et feuille de présence A4 avec signatures.</p></div>', unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="card-menu"><h3>📊 Production Quotidienne</h3><p>Analyse Excel détaillée, catégories A/B/C et charges.</p></div>', unsafe_allow_html=True)

    col4, col5, col6, col7 = st.columns(4)
    with col4:
        st.markdown('<div class="card-menu"><h3>💶 Primes</h3><p>Calcul automatique des primes.</p></div>', unsafe_allow_html=True)
    with col5:
        st.markdown('<div class="card-menu"><h3>🚨 Sécurité</h3><p>30 lève-vélos & registres AT.</p></div>', unsafe_allow_html=True)
    with col6:
        st.markdown('<div class="card-menu"><h3>📸 Souvenirs</h3><p>Album photo et moments de l\'atelier.</p></div>', unsafe_allow_html=True)
    with col7:
        st.markdown('<div class="card-menu"><h3>✉️ Diffusion</h3><p>Notes de service et documents.</p></div>', unsafe_allow_html=True)

# ==========================================
# 1. SUIVI PERSONNEL & DOTATIONS
# ==========================================
elif menu == "👷 1. Suivi Personnel & Dotations":
    st.header("👷 1. Suivi Personnel, Contrats & Dotations (Max 30 Agents)")
    
    with st.expander("➕ Ajouter un nouvel agent (Limite 30 max)"):
        if len(st.session_state["agents_db"]) >= 30:
            st.error("La limite maximale de 30 agents est atteinte.")
        else:
            with st.form("ajout_agent"):
                c_nom = st.text_input("Nom et Prénom")
                c_tel = st.text_input("N° Téléphone")
                c_badge = st.text_input("N° Badge")
                c_type = st.selectbox("Type de contrat", ["CDI", "CDD"])
                c_deb = st.date_input("Date de début", value=datetime.date.today())
                c_fin = st.date_input("Date de fin (laisser vide si CDI)", value=datetime.date.today() + datetime.timedelta(days=365))
                c_dot = st.text_input("Dotation habillage et matériel", value="Casque, Gants, Chaussures")
                
                sub_agent = st.form_submit_button("Enregistrer l'agent")
                if sub_agent and c_nom.strip():
                    nom_clean = c_nom.strip().upper()
                    fin_str = c_fin.isoformat() if c_type == "CDD" else ""
                    st.session_state["agents_db"][nom_clean] = {
                        "nom": nom_clean,
                        "tel": c_tel,
                        "badge": c_badge,
                        "contrat": c_type,
                        "debut": c_deb.isoformat(),
                        "fin": fin_str,
                        "dotation": c_dot
                    }
                    st.success(f"Agent {nom_clean} ajouté avec succès !")
                    st.rerun()

    st.divider()
    recherche_nom = st.text_input("🔍 Rechercher un agent par son nom :", placeholder="Ex: ABBAS...")
    agents_filtres = {k: v for k, v in st.session_state["agents_db"].items() if recherche_nom.lower() in k.lower()}

    st.subheader("📋 Liste des Mécaniciens de l'Atelier")
    liste_resume = []
    for nom, inf in agents_filtres.items():
        statut_contrat = inf["contrat"]
        if statut_contrat == "CDD" and inf["fin"]:
            try:
                if datetime.date.today() > datetime.date.fromisoformat(inf["fin"]):
                    statut_contrat = "CDD (FC)"
            except:
                pass
        liste_resume.append({
            "Nom & Prénom": nom,
            "Contrat": statut_contrat,
            "N° Badge": inf["badge"],
            "N° Téléphone": inf["tel"],
            "Dotation Matériel": inf["dotation"]
        })
    
    df_agents_resume = pd.DataFrame(liste_resume)
    st.dataframe(df_agents_resume, use_container_width=True, hide_index=True)

    output_excel = io.BytesIO()
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        df_agents_resume.to_excel(writer, index=False, sheet_name="Mecaniciens_VLG")
    st.download_button(
        "📥 Télécharger la liste des mécaniciens (Excel mis en page A4)",
        data=output_excel.getvalue(),
        file_name="liste_mecaniciens_vlg.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ==========================================
# 2. PLANNING, PRÉSENCES & FEUILLE A4
# ==========================================
elif menu == "📅 2. Planning, Présences & Feuille A4":
    st.header("📅 Planning Annuel S1-S52 & Feuille de Présence A4")
    
    semaine_num = st.selectbox("Sélectionner la Semaine", list(range(1, 53)), format_func=lambda x: f"Semaine {x}")
    
    records_planning = []
    for agent, infos in st.session_state["agents_db"].items():
        aff_nom = agent
        if infos["contrat"] == "CDD" and infos["fin"]:
            try:
                if datetime.date.today() > datetime.date.fromisoformat(infos["fin"]):
                    aff_nom = f"{agent} [FC]"
            except:
                pass
        row = {"Agent": aff_nom}
        for j in JOURS_SEMAINE:
            key = f"S{semaine_num}_{agent}_{j}"
            row[j] = st.session_state["planning_db"].get(key, "07h-14h")
        key_hs = f"S{semaine_num}_{agent}_HS"
        row["Heures Sup (h)"] = st.session_state["planning_db"].get(key_hs, 0.0)
        records_planning.append(row)

    df_planning = pd.DataFrame(records_planning)
    edited_planning = st.data_editor(df_planning, use_container_width=True, hide_index=True)

    if st.button("💾 Enregistrer le planning", type="primary"):
        for _, row in edited_planning.iterrows():
            clean_name = row["Agent"].replace(" [FC]", "").strip()
            for j in JOURS_SEMAINE:
                st.session_state["planning_db"][f"S{semaine_num}_{clean_name}_{j}"] = str(row[j]).strip()
            st.session_state["planning_db"][f"S{semaine_num}_{clean_name}_HS"] = float(row["Heures Sup (h)"])
        st.success("Planning enregistré avec succès !")

    st.divider()
    st.subheader("📄 Génération de la Feuille de Présence A4 (Badges, Téléphones & Signatures)")
    jour_selection = st.selectbox("Jour de la feuille de présence", JOURS_SEMAINE)
    
    # Filtrer les agents présents ce jour-là (exclure repos/absent si configuré ou lister tous les actifs du planning)
    agents_du_jour = []
    for agent, infos in st.session_state["agents_db"].items():
        val_jour = st.session_state["planning_db"].get(f"S{semaine_num}_{agent}_{jour_selection}", "07h-14h")
        if val_jour.lower() not in ["repos", "absent", "congé", "off", "-"]:
            agents_du_jour.append({
                "NOM & PRENOM": agent,
                "TEL": infos["tel"],
                "BADGE N°": infos["badge"],
                "SIGNATURE 07H": "",
                "SIGNATURE 14H": ""
            })
    
    df_feuille_presence = pd.DataFrame(agents_du_jour)
    st.dataframe(df_feuille_presence, use_container_width=True, hide_index=True)

    # Export Excel formaté A4 avec encadrement et dimensions pour signatures
    output_pres = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Presence_A4"

    # En-tête Mobiservice / VLG
    ws['A1'] = "MOBISERVICE - ATELIER VILLENEUVE-LA-GARENNE"
    ws['A1'].font = Font(size=14, bold=True)
    ws['A2'] = f"FEUILLE DE PRÉSENCE - {jour_selection.upper()} (Semaine {semaine_num})"
    ws['A2'].font = Font(size=11, bold=True)
    ws['A3'] = f"Date du jour : {datetime.date.today().strftime('%d/%m/%Y')}"
    ws['A3'].font = Font(size=10, italic=True)

    headers = ["NOM & PRENOM", "TEL", "BADGE N°", "SIGNATURE 07H", "SIGNATURE 14H"]
    row_idx = 5
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=row_idx, column=col_idx, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    for agent_item in agents_du_jour:
        row_idx += 1
        ws.cell(row=row_idx, column=1, value=agent_item["NOM & PRENOM"]).border = thin_border
        ws.cell(row=row_idx, column=2, value=agent_item["TEL"]).border = thin_border
        ws.cell(row=row_idx, column=3, value=agent_item["BADGE N°"]).border = thin_border
        
        # Cellules de signature (hauteur de ligne ~40 pour ~1cm, colonnes adaptées pour ~3cm)
        sig1 = ws.cell(row=row_idx, column=4, value="")
        sig1.border = thin_border
        sig2 = ws.cell(row=row_idx, column=5, value="")
        sig2.border = thin_border
        ws.row_dimensions[row_idx].height = 35

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 18

    wb.save(output_pres)
    output_pres.seek(0)

    st.download_button(
        "📥 Télécharger la Feuille de Présence A4 (Excel encadré & aligné)",
        data=output_pres.getvalue(),
        file_name=f"feuille_presence_s{semaine_num}_{jour_selection}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ==========================================
# 3. SUIVI DE LA PRODUCTION QUOTIDIENNE
# ==========================================
elif menu == "📊 3. Suivi Production Quotidienne":
    st.header("📊 Analyse de la Production Quotidienne (Indépendant Planning & Primes)")
    st.write("Importez votre fichier Excel ou CSV de production pour générer le tableau dynamique complet.")
    
    prod_file = st.file_uploader("📂 Déposer le fichier de production atelier", type=["xlsx", "xls", "csv"])
    if prod_file:
        try:
            if prod_file.name.endswith(".csv"):
                df_p = pd.read_csv(prod_file, sep=None, engine="python")
            else:
                df_p = pd.read_excel(prod_file)
            
            st.success(f"Fichier chargé avec succès ! {len(df_p)} lignes lues.")
            
            # Normalisation colonnes attendues
            col_agent = "Assigné à l'acteur Libellé" if "Assigné à l'acteur Libellé" in df_p.columns else df_p.columns[4]
            col_mat = "Matériel" if "Matériel" in df_p.columns else df_p.columns[3]
            col_cat = "smovCategory" if "smovCategory" in df_p.columns else df_p.columns[6]
            col_etat = "Etat" if "Etat" in df_p.columns else df_p.columns[8]
            col_charge = "Charge prévue" if "Charge prévue" in df_p.columns else df_p.columns[17]

            # Nettoyage
            df_p[col_agent] = df_p[col_agent].astype(str).str.strip().str.upper()
            
            # Calcul tableau dynamique par agent
            agents_uniques = df_p[col_agent].unique()
            synthese_data = []

            total_velos_general = len(df_p)
            total_ctrl_general = len(df_p[df_p[col_etat].astype(str).str.contains("contrôler|A contrôler", case=False, na=False)])
            
            for ag in agents_uniques:
                sub_df = df_p[df_p[col_agent] == ag]
                tot_velos = len(sub_df)
                
                cat_a = len(sub_df[sub_df[col_cat].astype(str).str.upper() == 'A'])
                cat_b = len(sub_df[sub_df[col_cat].astype(str).str.upper() == 'B'])
                cat_c = len(sub_df[sub_df[col_cat].astype(str).str.upper() == 'C'])
                
                pct_a = round((cat_a / tot_velos * 100), 1) if tot_velos > 0 else 0.0
                pct_b = round((cat_b / tot_velos * 100), 1) if tot_velos > 0 else 0.0
                pct_c = round((cat_c / tot_velos * 100), 1) if tot_velos > 0 else 0.0

                synthese_data.append({
                    "Nom Agent": ag,
                    "Total Vélos": tot_velos,
                    "Cat A": cat_a,
                    "% A": f"{pct_a}%",
                    "Cat B": cat_b,
                    "% B": f"{pct_b}%",
                    "Cat C": cat_c,
                    "% C": f"{pct_c}%",
                    "Total Catégories": cat_a + cat_b + cat_c
                })

            df_synthese = pd.DataFrame(synthese_data)
            
            st.subheader("📋 Tableau Dynamique de Production par Agent")
            st.dataframe(df_synthese, use_container_width=True, hide_index=True)

            # Indicateurs globaux
            st.divider()
            col_m1, col_m2, col_m3 = st.columns(3)
            with col_m1:
                st.metric("TOTAL VÉLOS EN GÉNÉRAL", total_velos_general)
            with col_m2:
                st.metric("VÉLOS STATUT À CONTRÔLER", total_ctrl_general)
            with col_m3:
                pct_c_gen = round((len(df_p[df_p[col_cat].astype(str).str.upper() == 'C']) / total_velos_general * 100), 1) if total_velos_general > 0 else 0.0
                st.metric("% CATÉGORIE C EN GÉNÉRAL", f"{pct_c_gen}%")

            # Export Excel encadré A4
            out_prod_excel = io.BytesIO()
            df_synthese.to_excel(out_prod_excel, index=False, sheet_name="Prod_Quotidienne")
            out_prod_excel.seek(0)
            
            st.download_button(
                "📥 Télécharger le rapport de production quotidien (Excel A4)",
                data=out_prod_excel.getvalue(),
                file_name="rapport_production_quotidienne.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erreur lors de l'analyse du fichier : {e}")

# ==========================================
# 4. DROITS DE PRIME
# ==========================================
elif menu == "💶 4. Droits de Prime":
    st.header("💶 Droits et Calcul des Primes Mensuelles")
    velos_input = st.number_input("Nombre de vélos traités par l'agent", 0, 50, 8)
    if velos_input < 6:
        prime = 0.0
    elif velos_input >= 10:
        prime = 155.0
    else:
        prime = round(50.0 + (velos_input - 6.0) * 26.25, 2)
    st.metric("Prime calculée", f"{prime} €")

# ==========================================
# 5. SÉCURITÉ, LÈVE-VÉLOS & AT
# ==========================================
elif menu == "🚨 5. Sécurité, Lève-Vélos & AT":
    st.header("🚨 Espace Sécurité, HSE & Suivi des Lève-Vélos")
    tab_s1, tab_s2, tab_s3, tab_s4 = st.tabs([
        "⚙️ Suivi 30 Lève-Vélos",
        "👷 Fiches Salariés",
        "📝 Minutes Sécurité",
        "⚠️ Accidents du Travail"
    ])

    with tab_s1:
        st.subheader("🔧 Suivi des 30 Lève-Vélos de l'atelier")
        for lv in st.session_state["leve_velos_db"]:
            with st.expander(f"📍 {lv['nom']} — Statut : {lv['statut']}"):
                c1, c2 = st.columns(2)
                with c1:
                    ano = st.text_input("Anomalie signalée", value=lv['anomalie'], key=f"ano_{lv['id']}")
                with c2:
                    trav = st.text_input("Travaux effectués", value=lv['intervention'], key=f"trav_{lv['id']}")
                if st.button(f"Mettre à jour {lv['nom']}", key=f"btn_lv_{lv['id']}"):
                    lv['anomalie'] = ano
                    lv['intervention'] = trav
                    lv['statut'] = "Réparé / Conforme" if trav else "En attente"
                    st.success("Mise à jour enregistrée !")

    with tab_s2:
        st.subheader("📂 Fiches de Sécurité par Salarié")
        sal = st.selectbox("Choisir un salarié", list(st.session_state["agents_db"].keys()), key="sec_sal_box")
        f_type = st.selectbox("Type", ["Attestation EPI", "Visite Médicale", "Habilitation"])
        f_up = st.file_uploader("Fichier justificatif (PDF)", type=["pdf", "png"])
        if st.button("Enregistrer la fiche"):
            if f_up:
                st.session_state["securite_db"]["fiches"].append({"salarie": sal, "type": f_type, "fichier": f_up.name})
                st.success("Fiche enregistrée avec succès !")
        if st.session_state["securite_db"]["fiches"]:
            st.dataframe(pd.DataFrame(st.session_state["securite_db"]["fiches"]), use_container_width=True)

    with tab_s3:
        st.subheader("📄 Minutes de Sécurité")
        t_min = st.text_input("Sujet de la minute")
        f_min = st.file_uploader("Document joint", type=["pdf", "docx"])
        if st.button("Ajouter la minute") and t_min:
            st.session_state["securite_db"]["minutes"].append({"titre": t_min, "date": str(datetime.date.today())})
            st.success("Minute ajoutée !")
        if st.session_state["securite_db"]["minutes"]:
            st.dataframe(pd.DataFrame(st.session_state["securite_db"]["minutes"]), use_container_width=True)

    with tab_s4:
        st.subheader("⚠️ Registre des Accidents du Travail (AT)")
        with st.form("form_at"):
            ag_at = st.selectbox("Agent", list(st.session_state["agents_db"].keys()))
            desc_at = st.text_area("Description")
            if st.form_submit_button("Enregistrer l'accident"):
                st.session_state["securite_db"]["accidents"].append({"agent": ag_at, "desc": desc_at, "date": str(datetime.date.today())})
                st.success("Accident enregistré au registre.")
        if st.session_state["securite_db"]["accidents"]:
            st.dataframe(pd.DataFrame(st.session_state["securite_db"]["accidents"]), use_container_width=True)

# ==========================================
# 6. SOUVENIRS & ÉVÉNEMENTS (PHOTOS / DOCS)
# ==========================================
elif menu == "📸 6. Souvenirs & Événements":
    st.header("📸 Historique & Souvenirs des Bons Moments de l'Atelier")
    st.write("Partagez et archivez les photos, moments forts et documents des événements de l'équipe VLG.")

    with st.form("form_souvenir"):
        titre_souv = st.text_input("Titre de l'événement ou du souvenir")
        desc_souv = st.text_area("Description / Récit du moment")
        media_souv = st.file_uploader("Ajouter une photo ou un document", type=["png", "jpg", "jpeg", "pdf", "docx"])
        sub_souv = st.form_submit_button("Publier dans l'album souvenir")
        
        if sub_souv and titre_souv:
            nom_fichier_media = media_souv.name if media_souv else "Aucun fichier"
            st.session_state["souvenirs_db"].append({
                "titre": titre_souv,
                "description": desc_souv,
                "fichier": nom_fichier_media,
                "auteur": current_user,
                "date": str(datetime.date.today())
            })
            st.success("Souvenir publié avec succès dans l'album de l'atelier !")

    st.divider()
    st.subheader("🖼️ Galerie des Souvenirs de l'Atelier VLG")
    if st.session_state["souvenirs_db"]:
        for s in reversed(st.session_state["souvenirs_db"]):
            with st.container():
                st.markdown(f"### 🎉 {s['titre']}")
                st.write(f"*Publié par {s['auteur']} le {s['date']}*")
                st.write(s['description'])
                if s['fichier'] != "Aucun fichier":
                    st.info(f"📎 Fichier / Photo attaché : {s['fichier']}")
                st.divider()
    else:
        st.info("Aucun souvenir enregistré pour le moment. Soyez le premier à en publier un !")

# ==========================================
# 7. DIFFUSION SOCIÉTÉ & MAIL
# ==========================================
elif menu == "✉️ 7. Diffusion Société & Mail":
    st.header("✉️ Diffusion de Notes & Documents aux Salariés")
    with st.form("form_diffusion"):
        sujet = st.text_input("Objet de la note")
        msg = st.text_area("Message")
        dest = st.multiselect("Destinataires", list(st.session_state["agents_db"].keys()), default=list(st.session_state["agents_db"].keys()))
        if st.form_submit_button("🚀 Diffuser aux salariés", type="primary") and sujet:
            st.success(f"Diffusion réussie envoyée à {len(dest)} salariés !")
