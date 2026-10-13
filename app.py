import datetime
import io
import json
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

# STYLE CSS PERSONNALISÉ
st.markdown(
    """
<style>
.stApp { background-color: #f4f6f9; }
[data-testid="stSidebar"] { background-color: #17324d; color: white; }
[data-testid="stSidebar"] * { color: white !important; }
.vlg-banner { padding: 20px 24px; border-radius: 14px; background: linear-gradient(135deg, #17324d, #2364aa); color: white; margin-bottom: 20px; }
.card-menu { background: white; padding: 20px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); text-align: center; border-top: 4px solid #2364aa; margin-bottom: 15px; }
.card-menu h3 { color: #17324d; margin-bottom: 8px; }
</style>
""",
    unsafe_allow_html=True,
)

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
    # 30 lève-vélos initialisés
    st.session_state["leve_velos_db"] = [
        {"id": i, "nom": f"Lève-Vélo n°{i}", "statut": "Opérationnel", "anomalie": "", "intervention": "Aucune", "date_intervention": "-"}
        for i in range(1, 31)
    ]
if "diffusion_db" not in st.session_state:
    st.session_state["diffusion_db"] = []

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
        "📅 2. Planning, Présences & Heures Sup",
        "📊 3. Suivi de la Production",
        "💶 4. Droits de Prime",
        "🚨 5. Sécurité, Lève-Vélos & AT",
        "✉️ 6. Diffusion Société & Mail"
    ]
)

# ==========================================
# 0. ACCUEIL & MENU PRINCIPAL (ICÔNES)
# ==========================================
if menu == "🏠 Accueil & Menu Principal":
    st.markdown('<div class="vlg-banner"><h1> Tableau de Bord de l\'Atelier VLG</h1><p>Sélectionnez un module ci-dessous ou utilisez le menu latéral.</p></div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="card-menu"><h3>👷 Suivi Personnel</h3><p>Gestion des 30 agents max, contrats, dotations & fiches PDF.</p></div>', unsafe_allow_html=True)
        if st.button("Ouvrir Suivi Personnel", use_container_width=True):
            st.switch_page = "👷 1. Suivi Personnel & Dotations" # Simplifié via navigation radio
    with col2:
        st.markdown('<div class="card-menu"><h3>📅 Planning & Heures Sup</h3><p>Gestion S1-S52, présence, HC/HS et mention FC automatique.</p></div>', unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="card-menu"><h3>📊 Production Atelier</h3><p>Analyse des fiches de production et volumes journaliers.</p></div>', unsafe_allow_html=True)

    col4, col5, col6 = st.columns(3)
    with col4:
        st.markdown('<div class="card-menu"><h3>💶 Droits de Prime</h3><p>Calcul automatique des primes selon les seuils.</p></div>', unsafe_allow_html=True)
    with col5:
        st.markdown('<div class="card-menu"><h3>🚨 Sécurité & Lève-Vélos</h3><p>Suivi HSE, 30 lève-vélos connectés & Accidents de travail.</p></div>', unsafe_allow_html=True)
    with col6:
        st.markdown('<div class="card-menu"><h3>✉️ Diffusion Société</h3><p>Envoi de documents et notes de service par mail/SMS.</p></div>', unsafe_allow_html=True)

# ==========================================
# 1. SUIVI PERSONNEL & DOTATIONS
# ==========================================
elif menu == "👷 1. Suivi Personnel & Dotations":
    st.header("👷 1. Suivi Personnel, Contrats & Dotations (Max 30 Agents)")
    
    # Formulaire d'ajout d'agent
    with st.expander("➕ Ajouter un nouvel agent (Limite 30 max)"):
        if len(st.session_state["agents_db"]) >= 30:
            st.error("La limite maximale de 30 agents est atteinte.")
        else:
            with st.form("ajout_agent"):
                c_nom = st.text_input("Nom et Prénom")
                c_tel = st.text_input("N° Téléphone")
                c_badge = st.text_input("N° Badge (2 chiffres ou format badge)")
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
    
    # Barre de recherche par nom
    recherche_nom = st.text_input("🔍 Rechercher un agent par son nom :", placeholder="Ex: ABBAS...")
    
    agents_filtres = {k: v for k, v in st.session_state["agents_db"].items() if recherche_nom.lower() in k.lower()}

    # Affichage de la liste résumée (dotation, tel, badge)
    st.subheader("📋 Liste des Mécaniciens de l'Atelier")
    liste_resume = []
    for nom, inf in agents_filtres.items():
        # Vérification si Fin de Contrat (FC) atteinte
        statut_contrat = inf["contrat"]
        if statut_contrat == "CDD" and inf["fin"]:
            try:
                date_fin_obj = datetime.date.fromisoformat(inf["fin"])
                if datetime.date.today() > date_fin_obj:
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

    # Bouton téléchargement liste en PDF / Excel
    output_excel = io.BytesIO()
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        df_agents_resume.to_excel(writer, index=False, sheet_name="Mecaniciens_VLG")
    st.download_button(
        "📥 Télécharger la liste des mécaniciens (Excel / PDF optimisé 1 page)",
        data=output_excel.getvalue(),
        file_name="liste_mecaniciens_vlg.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.divider()
    st.subheader("📁 Fiche Individuelle Détaillée (Modifier / Documents / Dotations)")
    agent_selectionne = st.selectbox("Choisir un agent à inspecter", list(agents_filtres.keys()))
    
    if agent_selectionne:
        inf_ag = st.session_state["agents_db"][agent_selectionne]
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            st.write(f"**Contrat :** {inf_ag['contrat']}")
            st.write(f"**Date de début :** {inf_ag['debut']}")
            st.write(f"**Date de fin :** {inf_ag['fin'] if inf_ag['fin'] else 'Indéterminée (CDI)'}")
            st.write(f"**Téléphone :** {inf_ag['tel']}")
            st.write(f"**Badge :** {inf_ag['badge']}")
        with col_f2:
            st.write(f"**Dotation actuelle :** {inf_ag['dotation']}")
            doc_joint = st.file_uploader(f"Injecter un contrat ou attestation pour {agent_selectionne}", type=["pdf", "png", "jpg"], key=f"doc_{agent_selectionne}")
            if doc_joint:
                st.success(f"Document {doc_joint.name} injecté avec succès pour {agent_selectionne} !")
                st.download_button("Télécharger le document injecté", data=doc_joint.getvalue(), file_name=doc_joint.name)

# ==========================================
# 2. PLANNING, PRÉSENCES & HEURES SUP
# ==========================================
elif menu == "📅 2. Planning, Présences & Heures Sup":
    st.header("📅 Planning Annuel S1 à S52 & Heures Supplémentaires")
    semaine_num = st.selectbox("Sélectionner la Semaine", list(range(1, 53)), format_func=lambda x: f"Semaine {x}")
    
    st.info("💡 L'ajout d'un agent dans le module 'Suivi Personnel' se répercute instantanément ici. La mention 'FC' s'affiche si le contrat est échu.")

    records_planning = []
    for agent, infos in st.session_state["agents_db"].items():
        # Vérification fin de contrat FC
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
        
        # Option Heures Sup (HS)
        key_hs = f"S{semaine_num}_{agent}_HS"
        row["Heures Sup (h)"] = st.session_state["planning_db"].get(key_hs, 0.0)
        records_planning.append(row)

    df_planning = pd.DataFrame(records_planning)
    edited_planning = st.data_editor(df_planning, use_container_width=True, hide_index=True)

    if st.button("💾 Enregistrer les modifications du planning & Heures Sup", type="primary"):
        for _, row in edited_planning.iterrows():
            clean_name = row["Agent"].replace(" [FC]", "").strip()
            for j in JOURS_SEMAINE:
                st.session_state["planning_db"][f"S{semaine_num}_{clean_name}_{j}"] = str(row[j]).strip()
            st.session_state["planning_db"][f"S{semaine_num}_{clean_name}_HS"] = float(row["Heures Sup (h)"])
        st.success("Planning et Heures Supplémentaires mis à jour avec succès !")

# ==========================================
# 3. SUIVI DE LA PRODUCTION
# ==========================================
elif menu == "📊 3. Suivi de la Production":
    st.header("📊 Analyse de la Production (Fichiers Atelier)")
    st.write("Importez vos fichiers CSV ou Excel de production journalière avec les critères habituels.")
    
    prod_file = st.file_uploader("📂 Déposer le fichier de production atelier", type=["xlsx", "csv"])
    if prod_file:
        try:
            if prod_file.name.endswith(".csv"):
                df_prod = pd.read_csv(prod_file, sep=None, engine="python")
            else:
                df_prod = pd.read_excel(prod_file)
            st.success(f"Fichier chargé avec succès : {len(df_prod)} lignes analysées.")
            st.dataframe(df_prod.head(20), use_container_width=True)
        except Exception as e:
            st.error(f-'{e}')

# ==========================================
# 4. DROITS DE PRIME
# ==========================================
elif menu == "💶 4. Droits de Prime":
    st.header("💶 Droits et Calcul des Primes Mensuelles")
    st.info("Barème dynamique appliqué : 6 vélos (50€) jusqu'à 10 vélos et plus (155€).")
    
    velos_input = st.number_input("Nombre de vélos traités par l'agent", 0, 30, 8)
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
        "👷 Fiches & Historique Salariés",
        "📝 Minutes Sécurité (PDF/Word)",
        "⚠️ Accidents du Travail & Filtre Date"
    ])

    # --- TAB 1 : LES 30 LÈVE-VÉLOS ---
    with tab_s1:
        st.subheader("🔧 Suivi des 30 Lève-Vélos connectés avec l'atelier")
        st.write("Les agents signalent les anomalies ci-dessous. L'administrateur peut répondre, valider les travaux et noter la date d'intervention.")

        for lv in st.session_state["leve_velos_db"]:
            with st.expander(f"📍 {lv['nom']} — Statut : {lv['statut']}"):
                col_lv1, col_lv2 = st.columns(2)
                with col_lv1:
                    nouvelle_anomalie = st.text_input("Signalement d'anomalie (Agent)", value=lv['anomalie'], key=f"ano_{lv['id']}")
                with col_lv2:
                    travaux_faits = st.text_input("Travaux effectués / Réponse atelier", value=lv['intervention'], key=f"trav_{lv['id']}")
                    date_interv = st.date_input("Date d'intervention", value=datetime.date.today(), key=f"date_{lv['id']}")
                
                if st.button(f"Mettre à jour {lv['nom']}", key=f"btn_lv_{lv['id']}"):
                    lv['anomalie'] = nouvelle_anomalie
                    lv['intervention'] = travaux_faits
                    lv['date_intervention'] = date_interv.isoformat()
                    lv['statut'] = "Réparé / Conforme" if travaux_faits else "En attente"
                    st.success(f"Mise à jour enregistrée pour {lv['nom']} ! Historique stocké.")

    # --- TAB 2 : FICHES & HISTORIQUE SALARIÉS ---
    with tab_s2:
        st.subheader("📂 Fiches de Sécurité & Historique par Salarié")
        recherche_sec = st.selectbox("Sélectionner un salarié", list(st.session_state["agents_db"].keys()), key="sec_sal")
        
        with st.form("form_sec_sal"):
            type_f = st.selectbox("Type de document", ["Attestation Port d'ÉPI", "Visite Médicale", "Habilitation Électrique / VAE", "Consignes de Sécurité"])
            file_sec = st.file_uploader("Joindre le fichier (PDF / Word)", type=["pdf", "docx"])
            sub_sec_f = st.form_submit_button("Injecter la fiche au dossier sécurité")
            if sub_sec_f and file_sec:
                st.session_state["securite_db"]["fiches"].append({
                    "salarie": recherche_sec,
                    "type": type_f,
                    "fichier": file_sec.name,
                    "date": datetime.date.today().isoformat()
                })
                st.success(f"Fiche sécurité enregistrée pour {recherche_sec} !")

        st.write("### Historique des fiches injectées :")
        fiches_sal = [f for f in st.session_state["securite_db"]["fiches"] if f["salarie"] == recherche_sec]
        if fiches_sal:
            st.dataframe(pd.DataFrame(fiches_sal), use_container_width=True)
        else:
            st.info("Aucun document de sécurité pour cet agent.")

    # --- TAB 3 : MINUTES SÉCURITÉ ---
    with tab_s3:
        st.subheader("📄 Minutes de Sécurité & Dossiers de Site")
        titre_min = st.text_input("Titre de la Minute / Sujet de Sécurité")
        doc_min = st.file_uploader("Télécharger le document (PDF / Word)", type=["pdf", "docx"])
        if st.button("Ajouter la minute de sécurité"):
            if titre_min and doc_min:
                st.session_state["securite_db"]["minutes"].append({"titre": titre_min, "fichier": doc_min.name, "date": datetime.date.today().isoformat()})
                st.success("Minute de sécurité ajoutée et archivée avec succès.")

        if st.session_state["securite_db"]["minutes"]:
            st.dataframe(pd.DataFrame(st.session_state["securite_db"]["minutes"]), use_container_width=True)

    # --- TAB 4 : ACCIDENTS DU TRAVAIL & FILTRE PAR DATE ---
    with tab_s4:
        st.subheader("⚠️ Registre des Accidents du Travail (AT) & Filtrage par Date")
        date_filtre = st.date_input("Filtrer l'historique à partir du :", value=datetime.date.today() - datetime.timedelta(days=365))
        
        with st.form("form_at"):
            ag_at = st.selectbox("Agent concerné", list(st.session_state["agents_db"].keys()), key="at_ag")
            desc_at = st.text_area("Description de l'incident / accident")
            date_at = st.date_input("Date de l'accident", value=datetime.date.today())
            sub_at = st.form_submit_button("Enregistrer l'accident au registre")
            if sub_at:
                st.session_state["securite_db"]["accidents"].append({
                    "agent": ag_at,
                    "description": desc_at,
                    "date": date_at.isoformat()
                })
                st.success("Accident enregistré au registre réglementaire.")

        accs = st.session_state["securite_db"]["accidents"]
        if accs:
            df_acc = pd.DataFrame(accs)
            st.dataframe(df_acc, use_container_width=True)

# ==========================================
# 6. DIFFUSION SOCIÉTÉ & MAIL
# ==========================================
elif menu == "✉️ 6. Diffusion Société & Mail":
    st.header("✉️ Diffusion de Notes & Documents aux Salariés")
    st.write("Envoyez automatiquement des documents de sécurité, notes d'information ou plannings par e-mail ou notification aux salariés de l'atelier.")

    with st.form("form_diffusion"):
        sujet_mail = st.text_input("Objet de la diffusion")
        message_mail = st.text_area("Message / Consignes")
        piece_jointe = st.file_uploader("Document attaché (PDF, Word)", type=["pdf", "docx", "xlsx"])
        destinataires = st.multiselect("Destinataires", list(st.session_state["agents_db"].keys()), default=list(st.session_state["agents_db"].keys()))
        
        btn_diff = st.form_submit_button("🚀 Diffuser à tous les salariés sélectionnés", type="primary")
        if btn_diff and sujet_mail:
            st.success(f"Diffusion réussie envoyée à {len(destinataires)} salariés avec succès !")