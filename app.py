import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, date, timedelta

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(
    page_title="Analyse Chef - Gestion Flotte & Insertion",
    page_icon="🚲",
    layout="wide"
)

# --- 1. SÉCURITÉ ET AUTHENTIFICATION ---
if "connecte" not in st.session_state:
    st.session_state.connecte = False

if not st.session_state.connecte:
    st.markdown("<h2 style='text-align: center;'>🔐 Connexion - Analyse Chef</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>Veuillez vous identifier pour accéder à l'application interne.</p>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        username = st.text_input("Identifiant")
        password = st.text_input("Mot de passe", type="password")
        
        if st.button("Se connecter", use_container_width=True):
            if username == "Admin" and password == "IMAD2014@":
                st.session_state.connecte = True
                st.rerun()
            else:
                st.error("Identifiant ou mot de passe incorrect.")
    st.stop()

# --- 2. BARRE LATÉRALE & PARAMÈTRES ---
st.sidebar.title("🚲 Menu Analyse Chef")
st.sidebar.success("Connecté en tant qu'Administrateur")

# Stockage & Mémoire
option_stockage = st.sidebar.selectbox(
    "Emplacement de la mémoire",
    ["PC Local (C:\\app_velos)", "Réseau d'entreprise / NAS", "Dossier Synchronisé (Drive)"]
)

st.sidebar.markdown("---")

# Importation des fichiers principaux et de comparaison (Anti-perte)
st.sidebar.subheader("📂 Fichiers de Production")
fichier_principal = st.sidebar.file_uploader("Fichier Excel / CSV principal", type=["xlsx", "xls", "csv"])
fichier_comparaison = st.sidebar.file_uploader("Fichier 2 (Pour Réconciliation / Anti-perte)", type=["xlsx", "xls", "csv"])

st.sidebar.markdown("---")

menu = st.sidebar.radio(
    "Navigation",
    [
        "📊 Tableau de Bord & Formulaires", 
        "📈 Graphiques & Visualisations", 
        "🕒 Charge Journalière & Mensuelle", 
        "👥 Agents, Contrats & Alertes", 
        "💰 Calculateur de Primes (Vélos Soldés)",
        "📅 Planning & Codes (ANR, AJNP, CP...)",
        "🔍 Réconciliation & Anti-Perte",
        "🧰 Sécurité, Outillage & Personnel"
    ]
)

if st.sidebar.button("Déconnexion"):
    st.session_state.connecte = False
    st.rerun()

# --- CHARGEMENT DES DONNÉES ---
@st.cache_data
def charger_donnees_fictives():
    data_interventions = pd.DataFrame({
        "Code de l'intervention": ["INT-001", "INT-002", "INT-003", "INT-004", "INT-005"],
        "Date/heure de fin de l'intervention": ["2026-10-01 10:00:00", "2026-10-02 11:30:00", "2026-10-02 14:00:00", "2026-10-03 09:15:00", "2026-10-04 16:45:00"],
        "Assigné à l'acteur Libellé": ["KEBE BAMBO", "DEDOU BENGARAA", "KEBE BAMBO", "KORAICH YOUNES", "FARID SALHI"],
        "Modèle d'équipement": ["VAE", "VLS", "VAE", "VLS", "VAE"],
        "Identifiant du formulaire personnalisé": ["A", "B", "C", "A", "B"],
        "Etat": ["Soudé", "Soudé", "À contrôler", "Soudé", "Soudé"],
        "Charge prévue": ["00:30:00", "00:45:00", "01:00:00", "00:25:00", "00:40:00"]
    })
    
    data_agents = pd.DataFrame({
        "Nom": ["KEBE BAMBO", "DEDOU BENGARAA", "KORAICH YOUNES", "FARID SALHI"],
        "Badge": ["B001", "B002", "B003", "B004"],
        "Tel_Travail": ["0611223344", "0622334455", "0633445566", "0644556677"],
        "Visiteur_ID": ["V-101", "V-102", "V-103", "V-104"],
        "Date_Debut": ["2026-01-10", "2025-11-01", "2026-05-10", "2026-05-20"],
        "Date_Fin": ["2026-11-05", "2026-11-15", "2026-06-05", "2027-05-20"], # Alerte pour Kebe et Dedou
        "Jours_Presents": [18, 19, 15, 20],
        "Heures_Moyennes_Jour": [7.5, 7.2, 6.8, 8.0], # Contrôle 7h minimum
        "Jours_ANR": [1, 0, 2, 0],
        "Jours_AJNP": [0, 1, 0, 1],
        "Jours_CP": [1, 1, 0, 0],
        "Jours_AT": [0, 0, 1, 0]
    })
    return data_interventions, data_agents

# Lecture du fichier principal
if fichier_principal is not None:
    try:
        if fichier_principal.name.endswith('.csv'):
            df_interventions = pd.read_csv(fichier_principal)
        else:
            df_interventions = pd.read_excel(fichier_principal)
        st.sidebar.success("✅ Fichier principal chargé !")
        _, df_agents = charger_donnees_fictives()
    except Exception as e:
        st.sidebar.error(f"Erreur de lecture : {e}")
        df_interventions, df_agents = charger_donnees_fictives()
else:
    df_interventions, df_agents = charger_donnees_fictives()

# Normalisation des dates de fin d'intervention
if "Date/heure de fin de l'intervention" in df_interventions.columns:
    df_interventions["Date_Fin_dt"] = pd.to_datetime(df_interventions["Date/heure de fin de l'intervention"], errors='coerce').dt.date
else:
    df_interventions["Date_Fin_dt"] = date.today()

df_agents["Date_Fin_dt"] = pd.to_datetime(df_agents["Date_Fin"], errors='coerce').dt.date

# --- MODULE D'ALERTE ROUGE (7 JOURS OUVRABLES) ---
def est_dans_7_jours_ouvrables(date_fin):
    if pd.isna(date_fin):
        return False
    aujourdhui = date.today()
    if date_fin < aujourdhui:
        return True
    jours_ouvrables = 0
    curseur = aujourdhui
    while curseur < date_fin and jours_ouvrables < 8:
        curseur += timedelta(days=1)
        if curseur.weekday() < 5:
            jours_ouvrables += 1
    return jours_ouvrables <= 7

agents_en_alerte = []
for idx, row in df_agents.iterrows():
    if est_dans_7_jours_ouvrables(row["Date_Fin_dt"]):
        agents_en_alerte.append(row["Nom"])

if agents_en_alerte:
    st.markdown(
        f"""
        <div style="background-color: #ff4b4b; color: white; padding: 15px; border-radius: 8px; margin-bottom: 20px;">
            <h3>🚨 ALERTE CRITIQUE : Fin de contrat imminente (≤ 7 jours ouvrables)</h3>
            <p>Agents concernés : <b>{', '.join(agents_en_alerte)}</b>. Veuillez régulariser leur situation.</p>
        </div>
        """,
        unsafe_allow_html=True
    )

# --- 3. CONTENU DES ONGLETS ---

if menu == "📊 Tableau de Bord & Formulaires":
    st.title("📊 Tableau de Bord & Analyse de Production")
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Interventions Totales", len(df_interventions))
    col2.metric("Agents Actifs", len(df_agents))
    col3.metric("Stockage Actif", option_stockage)
    col4.metric("Fichiers chargés", "Oui" if fichier_principal else "Démo")
    
    st.markdown("---")
    st.subheader("📋 Synthèse des Formulaires (A, B, C) & États")
    if "Identifiant du formulaire personnalisé" in df_interventions.columns:
        form_summary = df_interventions.groupby("Identifiant du formulaire personnalisé").agg(
            Total=("Code de l'intervention", "count")
        ).reset_index()
        st.dataframe(form_summary, use_container_width=True)
    else:
        st.write("Colonne de formulaire non détectée dans ce fichier.")

elif menu == "📈 Graphiques & Visualisations":
    st.title("📈 Visualisations Graphiques Avancées (Barres, Camemberts, Courbes)")
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📊 Vélos traités par Mécanicien")
        if "Assigné à l'acteur Libellé" in df_interventions.columns:
            fig, ax = plt.subplots()
            acteur_counts = df_interventions["Assigné à l'acteur Libellé"].value_counts()
            acteur_counts.plot(kind="bar", ax=ax, color="#2b5c8f")
            ax.set_ylabel("Nombre d'interventions")
            st.pyplot(fig)
        else:
            st.info("Données d'acteurs indisponibles.")
            
    with col2:
        st.subheader("🥧 Répartition par Modèle d'Équipement")
        if "Modèle d'équipement" in df_interventions.columns:
            fig2, ax2 = plt.subplots()
            modele_counts = df_interventions["Modèle d'équipement"].value_counts()
            modele_counts.plot(kind="pie", ax=ax2, autopct="%1.1f%%", startangle=90)
            ax2.set_ylabel("")
            st.pyplot(fig2)
        else:
            st.info("Données de modèles indisponibles.")

elif menu == "🕒 Charge Journalière & Mensuelle":
    st.title("🕒 Suivi de la Charge de Travail (Jour & Mois)")
    
    # Sélecteur de période de travaux
    st.subheader("📅 Sélectionner la période de travail")
    col_d1, col_d2 = st.columns(2)
    date_debut_sel = col_d1.date_input("Date de début", date.today() - timedelta(days=30))
    date_fin_sel = col_d2.date_input("Date de fin", date.today())
    
    # Filtrage par date
    df_f = df_interventions[(df_interventions["Date_Fin_dt"] >= date_debut_sel) & (df_interventions["Date_Fin_dt"] <= date_fin_sel)]
    
    st.write(f"Nombre d'interventions sur la période sélectionnée : **{len(df_f)}**")
    
    if not df_f.empty and "Date_Fin_dt" in df_f.columns:
        charge_jour = df_f.groupby("Date_Fin_dt").size().reset_index(name="Nombre_Interventions")
        st.subheader("📈 Courbe d'évolution journalière")
        st.line_chart(charge_jour.set_index("Date_Fin_dt"))
        st.dataframe(charge_jour, use_container_width=True)
    else:
        st.warning("Aucune donnée pour cette période.")

elif menu == "👥 Agents, Contrats & Alertes":
    st.title("👥 Gestion des Salariés en Insertion & Fiches Personnel")
    st.write("Liste complète des agents, numéros de badge, téléphones de travail et visiteurs :")
    
    st.dataframe(df_agents[["Nom", "Badge", "Tel_Travail", "Visiteur_ID", "Date_Debut", "Date_Fin"]], use_container_width=True)
    
    st.info("💡 L'administrateur peut modifier ou ajouter des agents via le panneau de configuration interne.")

elif menu == "💰 Calculateur de Primes (Vélos Soldés)":
    st.title("💰 Module de Calcul des Primes (Basé uniquement sur les Vélos Soldés)")
    st.markdown("""
    * **Règle des paliers :** Début de prime à **7 vélos/jour** (60 €) jusqu'à **10 vélos/jour** (155 €) avec interpolation progressive.
    * **Prise en compte du planning :** Les codes **ANR** et **AJNP** sont comptés comme des jours travaillés/assimilés.
    * **Contrôle de présence :** Vérification du seuil minimum de présence (ex: 7h/jour).
    """)
    
    # Filtrer uniquement les vélos soldés ("Soudé" ou état validé)
    if "Etat" in df_interventions.columns:
        df_soldes = df_interventions[df_interventions["Etat"].str.contains("Soudé|Soldé", case=False, na=False)]
    else:
        df_soldes = df_interventions # Par défaut si la colonne n'existe pas
        
    resultats_primes = []
    for idx, row in df_agents.iterrows():
        nom_agent = row["Nom"]
        # Compter les vélos soldés de cet agent dans le fichier
        if "Assigné à l'acteur Libellé" in df_soldes.columns:
            velos_soldes = len(df_soldes[df_soldes["Assigné à l'acteur Libellé"].str.strip().str.upper() == nom_agent.strip().str.upper()])
        else:
            velos_soldes = 120 # Valeur par défaut de secours
            
        # Jours travaillés + assimilation ANR et AJNP
        jours_utiles = row["Jours_Presents"] + row["Jours_ANR"] + row["Jours_AJNP"]
        moyenne = velos_soldes / jours_utiles if jours_utiles > 0 else 0
        
        # Contrôle de présence minimum (7h/jour)
        respect_temps = row["Heures_Moyennes_Jour"] >= 7.0
        
        # Calcul de la prime progressive
        if not respect_temps or moyenne < 7.0:
            prime_th = 0.0
        elif 7.0 <= moyenne <= 10.0:
            prime_th = 60.0 + ((moyenne - 7.0) / 3.0) * (155.0 - 60.0)
        else:
            prime_th = 155.0
            
        taux_pres = min(jours_utiles / 20.0, 1.0)
        prime_fin = prime_th * taux_pres
        
        resultats_primes.append({
            "Agent": nom_agent,
            "Vélos Soldés": velos_soldes,
            "Jours Travaillés + ANR/AJNP": jours_utiles,
            "Moyenne Vélos/Jour": round(moyenne, 2),
            "Présence ≥ 7h/j": "✅ Oui" if respect_temps else "❌ Non",
            "Prime Finale (€)": round(prime_fin, 2)
        })
        
    df_primes_calc = pd.DataFrame(resultats_primes)
    st.dataframe(df_primes_calc, use_container_width=True)

elif menu == "📅 Planning & Codes (ANR, AJNP, CP...)":
    st.title("📅 Gestion du Planning & Codes d'Absence / Présence (7h00 - 16h00)")
    st.markdown("""
    **Codes pris en charge dans le système :**
    `ANR`, `AJNP`, `CP`, `AT`, `ARM`, `CPP`, `DGT`, `IAE`, `ABS`, `FOR`, `DISP`, `CONT`, `RET`.
    """)
    
    st.subheader("Tableau récapitulatif des pointages et codes du personnel")
    st.dataframe(df_agents[["Nom", "Jours_Presents", "Jours_ANR", "Jours_AJNP", "Jours_CP", "Jours_AT", "Heures_Moyennes_Jour"]], use_container_width=True)
    st.info("Plage horaire standard de suivi d'atelier : **07h00 à 16h00** (Minimum 7h requis par jour).")

elif menu == "🔍 Réconciliation & Anti-Perte":
    st.title("🔍 Réconciliation & Interventions Perdues (Comparaison de 2 Fichiers)")
    st.write("Ce module compare le fichier principal et le fichier 2 pour identifier les numéros d'intervention et matériel perdus ou non facturés.")
    
    if fichier_comparaison is not None:
        try:
            if fichier_comparaison.name.endswith('.csv'):
                df_comp = pd.read_csv(fichier_comparaison)
            else:
                df_comp = pd.read_excel(fichier_comparaison)
            
            st.success("✅ Fichier 2 chargé avec succès pour comparaison !")
            
            # Comparaison par code d'intervention si disponible
            if "Code de l'intervention" in df_interventions.columns and "Code de l'intervention" in df_comp.columns:
                manquants = set(df_interventions["Code de l'intervention"]) - set(df_comp["Code de l'intervention"])
                if manquants:
                    st.warning(f"⚠️ {len(manquants)} intervention(s) présente(s) dans le fichier 1 mais absentes du fichier 2 :")
                    st.write(list(manquants))
                else:
                    st.success("✅ Aucune divergence d'intervention constatée entre les deux fichiers.")
            else:
                st.info("Les colonnes de codes d'intervention ne permettent pas un croisement direct automatique.")
        except Exception as e:
            st.error(f"Erreur lors de la lecture du fichier de comparaison : {e}")
    else:
        st.warning("⚠️ Veuillez importer un **Fichier 2** dans la barre latérale pour activer la réconciliation anti-perte.")

elif menu == "🧰 Sécurité, Outillage & Personnel":
    st.title("🧰 Suivi Sécurité, Outillage (40 Outils) & Registre")
    
    tab1, tab2, tab3 = st.tabs(["Outillage (40 Outils)", "Suivi Levage & Sécurité", "Feuilles d'Émargement"])
    
    with tab1:
        st.subheader("Gestion des dotations et prêts d'outils (Perceuses, Pinces, Clés...)")
        outils_data = pd.DataFrame({
            "ID Outil": [f"OUT-{i:02d}" for i in range(1, 41)],
            "Nom de l'outil": ["Perceuse sans fil", "Pince coupante", "Clé dynamométrique", "Tournevis isolés", "Multimètre"] * 8,
            "État": ["Bon", "En service", "En service", "En maintenance", "Bon"] * 8,
            "Attribué à": ["KEBE BAMBO", "DEDOU BENGARAA", "KORAICH YOUNES", "FARID SALHI", "Non assigné"] * 8
        })
        st.dataframe(outils_data, use_container_width=True)
        
    with tab2:
        st.subheader("Registre de Sécurité & Levage des Vélos")
        st.write("Suivi obligatoire des postes de levage (24 postes) et dates de diagnostic :")
        securite_data = pd.DataFrame({
            "Poste de Levage": [f"Poste #{i}" for i in range(1, 25)],
            "Date de Dernier Contrôle": pd.date_range(start="2026-10-01", periods=24, freq="D").strftime("%Y-%m-%d"),
            "État de Conformité": ["Conforme", "Conforme", "À réviser", "Conforme"] * 6,
            "Sujet / Remarque": ["RAS", "Vérification vérin", "Graissage chaîne", "RAS"] * 6
        })
        st.dataframe(securite_data, use_container_width=True)
        
    with tab3:
        st.subheader("Feuilles d'Émargement & Suivi Journalier")
        st.write("Registre des signatures de présence et émargement des équipes d'insertion par date.")
        emargement_sample = pd.DataFrame({
            "Date": ["2026-10-28", "2026-10-28", "2026-10-28"],
            "Agent": ["KEBE BAMBO", "DEDOU BENGARAA", "KORAICH YOUNES"],
            "Heure Arrivée (07h00+)": ["07:02", "06:58", "07:15"],
            "Heure Départ (16h00)": ["16:05", "16:00", "16:10"],
            "Signature Émargement": ["Validé", "Validé", "Validé"]
        })
        st.dataframe(emargement_sample, use_container_width=True)
