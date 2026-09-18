import streamlit as st
import pandas as pd
import random
import re
import os
from difflib import SequenceMatcher
from gtts import gTTS
import io

# ==========================================
# CONFIGURATION ET DESIGN RESPONSIVE
# ==========================================
st.set_page_config(page_title="LingoApp", page_icon="🎓", layout="wide")

def appliquer_style_moderne():
    st.markdown("""
    <link rel="apple-touch-icon" href="https://img.icons8.com/color/512/language.png">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;800;900&display=swap');
        html, body, [class*="css"] { font-family: 'Nunito', sans-serif !important; }
        .stApp { background: linear-gradient(135deg, #a1c4fd 0%, #c2e9fb 100%); padding: 0px !important; }
        .block-container { padding-top: 1rem !important; padding-bottom: 2rem !important; padding-left: 0.8rem !important; padding-right: 0.8rem !important; max-width: 100% !important; }
        h1 { font-size: 1.8rem !important; font-weight: 900 !important; color: #1e293b !important; margin-bottom: 0.5rem !important; }
        p, span, label, li { color: #334155 !important; font-size: 1rem !important; }
        div[role="radiogroup"] { background: rgba(255, 255, 255, 0.85) !important; border-radius: 16px !important; padding: 5px !important; display: flex !important; flex-direction: row !important; justify-content: space-between !important; margin-bottom: 1rem !important; }
        div[role="radiogroup"] label { flex: 1 1 auto !important; text-align: center !important; padding: 8px 4px !important; border-radius: 12px !important; margin: 0 !important; border: none !important; background: transparent !important; }
        div[role="radiogroup"] label[data-checked="true"] { background: #ffffff !important; box-shadow: 0 2px 8px rgba(0,0,0,0.1) !important; }
        .stButton>button { border-radius: 16px !important; font-weight: 800 !important; font-size: 1.05rem !important; padding: 12px 20px !important; min-height: 52px !important; width: 100% !important; background-color: #ffffff !important; color: #1e293b !important; border: 1px solid rgba(255, 255, 255, 0.9) !important; }
        .stButton>button[kind="primary"], div[data-testid="stForm"] button { background: linear-gradient(135deg, #ff5f6d 0%, #ff2a4b 100%) !important; color: white !important; border: none !important; }
        [data-testid="stForm"], div[data-testid="stExpander"] { background-color: rgba(255, 255, 255, 0.8) !important; border-radius: 20px !important; padding: 18px !important; box-shadow: 0 8px 25px rgba(0,0,0,0.05) !important; border: 1px solid rgba(255, 255, 255, 0.8) !important; backdrop-filter: blur(12px); }
        .stTextInput>div>div>input, .stNumberInput>div>div>input, select { border-radius: 14px !important; color: #0f172a !important; background-color: #ffffff !important; border: 1px solid #cbd5e1 !important; font-size: 1.1rem !important; height: 48px !important; padding: 12px !important; }
        .stTabs [data-baseweb="tab-list"] { gap: 10px !important; }
        .stTabs [data-baseweb="tab"] { border-radius: 12px 12px 0 0 !important; background-color: rgba(255,255,255,0.5) !important; font-weight: 700 !important; padding: 10px 16px !important; }
        .stTabs [data-baseweb="tab"][aria-selected="true"] { background-color: #ffffff !important; }
        div[data-testid="stDataFrame"] { background: white !important; border-radius: 16px !important; padding: 5px !important; box-shadow: 0 4px 15px rgba(0,0,0,0.05) !important; }
    </style>
    """, unsafe_allow_html=True)

appliquer_style_moderne()

FICHIER_EXCEL = "Vocabulaire d'anglais à connaitre .xlsx"
LANGUES_COLS = ["Francais", "Anglais", "Espagnol", "Italien", "Allemand"]
LANGUES_AFFICHAGE = {"Francais": "🥖 Français", "Anglais": "☕ Anglais", "Espagnol": "💃 Espagnol", "Italien": "🍕 Italien", "Allemand": "🥨 Allemand"}
CODES_AUDIO = {"Francais": "fr", "Anglais": "en", "Espagnol": "es", "Italien": "it", "Allemand": "de"}
INVERS_AFFICHAGE = {v: k for k, v in LANGUES_AFFICHAGE.items()}

def obtenir_picto_cat(cat_nom):
    nom = str(cat_nom).strip().lower()
    if any(k in nom for k in ["fruit", "legume", "légume", "nourriture", "aliment"]): icon = "🍎"
    elif any(k in nom for k in ["securit", "sécurité", "casque", "protection"]): icon = "🪖"
    elif any(k in nom for k in ["outil", "matériel", "materiel"]): icon = "🔨"
    elif any(k in nom for k in ["entreprise", "usine", "societe", "société"]): icon = "🏭"
    elif "environne" in nom: icon = "🍃"
    elif any(k in nom for k in ["maison", "interieur", "intérieur"]): icon = "🏠"
    elif any(k in nom for k in ["soin", "sante", "santé", "medical"]): icon = "🩺"
    elif any(k in nom for k in ["voyage", "transport", "voiture"]): icon = "✈️"
    elif any(k in nom for k in ["travail", "bureau"]): icon = "💼"
    elif "animal" in nom or "animaux" in nom: icon = "🐾"
    elif any(k in nom for k in ["vetement", "vêtement", "habille"]): icon = "👕"
    elif "corps" in nom: icon = "🧠"
    elif "ville" in nom: icon = "🏙️"
    elif "nature" in nom: icon = "🌿"
    elif "autre" in nom: icon = "📦"
    elif "préposition" in nom or "preposition" in nom: icon = "🔗"
    elif "verbe" in nom: icon = "⚡"
    elif "toiec" in nom or "toeic" in nom: icon = "🎓"
    else: icon = "🏷️"
    return f"{icon} {cat_nom}"

CORRECTIONS = {
    'crouch': 'couch', 'écureil': 'écureuil', 'attérir': 'atterrir',
    'machoire': 'mâchoire', 'bell peper': 'bell pepper', 'zuchini': 'zucchini', 'chadelle': 'chandelle', 'chuchotter': 'chuchoter'
}

def obtenir_nom_fichier(utilisateur):
    nom_propre = "".join(x for x in utilisateur if x.isalnum()).lower()
    return f"base_profil_{nom_propre}.csv"

def preparer_base_globale(utilisateur):
    fichier_utilisateur = obtenir_nom_fichier(utilisateur)
    if os.path.exists(fichier_utilisateur): return pd.read_csv(fichier_utilisateur)
    if os.path.exists(FICHIER_EXCEL):
        xls = pd.ExcelFile(FICHIER_EXCEL)
        df_brut = pd.read_excel(xls, sheet_name=xls.sheet_names[0])
        colonnes = df_brut.columns.tolist()
        liste_mots = []
        for i in range(0, len(colonnes), 3):
            categorie = str(colonnes[i]).strip()
            temp_df = df_brut[[colonnes[i], colonnes[i+1]]].copy()
            temp_df.columns = ['Anglais', 'Francais']
            temp_df['Categorie'] = categorie
            temp_df = temp_df.dropna(subset=['Anglais', 'Francais']) 
            liste_mots.append(temp_df)
        df = pd.concat(liste_mots, ignore_index=True)
        df['Anglais'] = df['Anglais'].astype(str).str.strip().replace(CORRECTIONS)
        df['Francais'] = df['Francais'].astype(str).str.strip().replace(CORRECTIONS)
        for lang in ["Espagnol", "Italien", "Allemand"]: df[lang] = None
        df['Score_100'] = 0; df['Score_50'] = 0; df['Score_0'] = 0; df['Pourcentage'] = 0.0
        df = df[['Categorie'] + LANGUES_COLS + ['Score_100', 'Score_50', 'Score_0', 'Pourcentage']]
        df.to_csv(fichier_utilisateur, index=False)
        return df
    return pd.DataFrame(columns=['Categorie'] + LANGUES_COLS + ['Score_100', 'Score_50', 'Score_0', 'Pourcentage'])

def sauvegarder_base():
    fichier = obtenir_nom_fichier(st.session_state.utilisateur_connecte)
    st.session_state.vocabulaire.to_csv(fichier, index=False)

def verifier_reponse(reponse_utilisateur, reponse_attendue):
    pattern = r"^(the |a |an |to |le |la |l'|les |un |une |des )"
    user_clean = re.sub(pattern, '', str(reponse_utilisateur).lower().strip())
    attendu_clean = re.sub(pattern, '', str(reponse_attendue).lower().strip())
    if user_clean == attendu_clean: return 100
    if SequenceMatcher(None, user_clean, attendu_clean).ratio() >= 0.8: return 50
    return 0

def generer_audio(texte, code_langue):
    tts = gTTS(text=str(texte), lang=code_langue)
    fp = io.BytesIO()
    tts.write_to_fp(fp)
    fp.seek(0)
    return fp

def calculer_pourcentage(row):
    total_tentatives = row['Score_100'] + row['Score_50'] + row['Score_0']
    if total_tentatives == 0: return 0.0
    points = (row['Score_100'] * 1.0) + (row['Score_50'] * 0.5)
    return round((points / total_tentatives) * 100, 1)


# ==========================================
# GESTION DE L'IDENTIFICATION SANS MOT DE PASSE
# ==========================================
query_user = st.query_params.get("u", None)

if not query_user:
    st.markdown("<h1 style='text-align: center;'>🎓 LingoApp</h1>", unsafe_allow_html=True)
    st.write("---")
    st.write("### 👋 Bienvenue !")
    st.write("Aucun mot de passe n'est requis. Pour ne pas mélanger tes scores avec les autres, entre simplement ton prénom ci-dessous :")
    
    with st.form("form_login"):
        pseudo_input = st.text_input("👤 Ton prénom :").strip()
        if st.form_submit_button("C'est parti ! 🚀", type="primary", use_container_width=True):
            if pseudo_input:
                st.query_params["u"] = pseudo_input
                st.rerun()
            else:
                st.warning("Merci d'entrer un prénom.")
    st.stop()

# Si l'utilisateur est présent dans l'URL
if 'utilisateur_connecte' not in st.session_state or st.session_state.utilisateur_connecte != query_user:
    st.session_state.utilisateur_connecte = query_user
    st.session_state.vocabulaire = preparer_base_globale(query_user)

if 'mot_actuel' not in st.session_state: st.session_state.mot_actuel = None
if 'test_etape' not in st.session_state: st.session_state.test_etape = "config"
if 'entrainement_etat' not in st.session_state: st.session_state.entrainement_etat = "attente"

# ==========================================
# HEADER ET NAVIGATION (UTILISATEUR CONNECTÉ)
# ==========================================
col_profil, col_logout = st.columns([3, 1])
with col_profil: st.markdown(f"**👤 Profil de {st.session_state.utilisateur_connecte}**")
with col_logout:
    if st.button("🚪 Changer", use_container_width=True):
        st.query_params.clear()
        st.session_state.utilisateur_connecte = None
        st.session_state.vocabulaire = None
        st.rerun()

menu = st.radio("Navigation :", ["📝 Quiz", "⏱️ Test", "📖 Dict.", "📊 Scores", "⚙️ Gérer"], horizontal=True, label_visibility="collapsed")

# ==========================================
# PAGE 1 : ENTRAÎNEMENT
# ==========================================
if menu == "📝 Quiz":
    st.title("🎓 Entraînement")
    st.session_state.test_etape = "config" 
    
    col1, col2 = st.columns(2)
    with col1: choix_depart = st.selectbox("Question :", list(LANGUES_AFFICHAGE.values()), index=0)
    with col2: choix_arrivee = st.selectbox("Réponse :", list(LANGUES_AFFICHAGE.values()), index=1)
    lang_dep, lang_arr = INVERS_AFFICHAGE[choix_depart], INVERS_AFFICHAGE[choix_arrivee]
    
    if lang_dep == lang_arr:
        st.warning("Choisissez deux langues différentes !")
    else:
        df_quiz = st.session_state.vocabulaire.dropna(subset=[lang_dep, lang_arr])
        if len(df_quiz) > 0:
            if st.session_state.entrainement_etat == "attente":
                st.write("---")
                if st.button("🚀 Commencer l'entraînement", type="primary", use_container_width=True):
                    poids = [100 if (row['Score_100'] + row['Score_50'] + row['Score_0']) == 0 else max(1, 101 - row['Pourcentage']) for _, row in df_quiz.iterrows()]
                    st.session_state.mot_actuel = df_quiz.loc[random.choices(df_quiz.index, weights=poids, k=1)[0]]
                    st.session_state.entrainement_etat = "reponse"
                    st.rerun()
            
            elif st.session_state.mot_actuel is not None:
                mot = st.session_state.mot_actuel
                index_mot = mot.name
                pourcentage_actuel = st.session_state.vocabulaire.at[index_mot, 'Pourcentage']
                
                with st.container():
                    st.subheader(f"Traduisez : **{mot[lang_dep]}**")
                    st.caption(f"Catégorie : {obtenir_picto_cat(mot['Categorie'])} | Score : **{pourcentage_actuel}%**")
                    st.audio(generer_audio(mot[lang_dep], CODES_AUDIO[lang_dep]), format="audio/mp3")
                    st.write("")
                        
                    if st.session_state.entrainement_etat == "reponse":
                        with st.form("form_entrainement", clear_on_submit=True):
                            reponse = st.text_input("Votre réponse :", key="input_quiz")
                            if st.form_submit_button("✅ Valider", use_container_width=True):
                                if reponse:
                                    score = verifier_reponse(reponse, mot[lang_arr])
                                    if score == 100: st.session_state.vocabulaire.at[index_mot, 'Score_100'] += 1
                                    elif score == 50: st.session_state.vocabulaire.at[index_mot, 'Score_50'] += 1
                                    else: st.session_state.vocabulaire.at[index_mot, 'Score_0'] += 1
                                    
                                    st.session_state.vocabulaire.at[index_mot, 'Pourcentage'] = calculer_pourcentage(st.session_state.vocabulaire.loc[index_mot])
                                    sauvegarder_base()
                                    
                                    st.session_state.dernier_score = score
                                    st.session_state.derniere_reponse = reponse
                                    st.session_state.entrainement_etat = "correction"
                                    st.rerun()
                                else: st.warning("Vous n'avez rien écrit !")
                                
                    elif st.session_state.entrainement_etat == "correction":
                        score = st.session_state.dernier_score
                        if score == 100:
                            st.success(f"**PARFAIT !** La réponse était : {mot[lang_arr]}")
                        elif score == 50:
                            st.warning(f"**PRESQUE !** Vous avez écrit '{st.session_state.derniere_reponse}'. Exact : **{mot[lang_arr]}**")
                        else:
                            st.error(f"**FAUX.** Vous avez écrit '{st.session_state.derniere_reponse}'. Exact : **{mot[lang_arr]}**")
                            
                        st.audio(generer_audio(mot[lang_arr], CODES_AUDIO[lang_arr]), format="audio/mp3")
                        st.write("")
                        if st.button("➡️ Mot suivant", type="primary", use_container_width=True):
                            poids = [100 if (row['Score_100'] + row['Score_50'] + row['Score_0']) == 0 else max(1, 101 - row['Pourcentage']) for _, row in df_quiz.iterrows()]
                            st.session_state.mot_actuel = df_quiz.loc[random.choices(df_quiz.index, weights=poids, k=1)[0]]
                            st.session_state.entrainement_etat = "reponse"
                            st.rerun()
                
                st.write("")
                if st.button("🔄 Changer de langue / Réinitialiser", use_container_width=True):
                    st.session_state.mot_actuel = None
                    st.session_state.entrainement_etat = "attente"
                    st.rerun()

# ==========================================
# PAGE 2 : LE TEST
# ==========================================
elif menu == "⏱️ Test":
    st.title("⏱️ Évaluation")
    st.session_state.entrainement_etat = "attente"
    
    if st.session_state.test_etape == "config":
        col1, col2 = st.columns(2)
        with col1: choix_depart = st.selectbox("Départ :", list(LANGUES_AFFICHAGE.values()), index=0)
        with col2: choix_arrivee = st.selectbox("Arrivée :", list(LANGUES_AFFICHAGE.values()), index=1)
        lang_dep, lang_arr = INVERS_AFFICHAGE[choix_depart], INVERS_AFFICHAGE[choix_arrivee]

        categories_raw = sorted([str(c) for c in st.session_state.vocabulaire['Categorie'].unique() if pd.notna(c) and c != ""])
        cat_options = ["🌍 Toutes les catégories"] + [obtenir_picto_cat(c) for c in categories_raw]
        choix_cat_test = st.selectbox("📂 Catégorie à tester :", cat_options)

        critere = st.radio("Critère :", ["Score le plus faible", "Mots sous un certain %", "Au hasard"])
        
        nb_mots = st.number_input("Nombre de mots :", min_value=1, max_value=200, value=20)
        seuil_pct = 50
        if critere == "Mots sous un certain %": 
            seuil_pct = st.slider("Pourcentage max :", 0, 100, 50)
                
        if st.button("🚀 Démarrer le Test", type="primary", use_container_width=True):
            if lang_dep == lang_arr: st.error("Choisissez des langues différentes.")
            else:
                df_dispo = st.session_state.vocabulaire.dropna(subset=[lang_dep, lang_arr])
                
                if choix_cat_test != "🌍 Toutes les catégories":
                    cat_pure = choix_cat_test.split(" ", 1)[1] if " " in choix_cat_test else choix_cat_test
                    df_dispo = df_dispo[df_dispo['Categorie'] == cat_pure]

                if len(df_dispo) > 0:
                    if critere == "Mots sous un certain %":
                        df_filtre = df_dispo[df_dispo['Pourcentage'] < seuil_pct]
                        df_test = df_filtre.sample(n=min(nb_mots, len(df_filtre)))
                    elif critere == "Score le plus faible":
                        df_test = df_dispo.sort_values(by=['Pourcentage']).head(nb_mots).sample(frac=1)
                    else:
                        df_test = df_dispo.sample(n=min(nb_mots, len(df_dispo)))
                    
                    if len(df_test) > 0:
                        st.session_state.test_mots = df_test
                        st.session_state.test_lang_dep = lang_dep
                        st.session_state.test_lang_arr = lang_arr
                        st.session_state.test_etape = "en_cours"
                        st.rerun()
                    else: st.warning("Aucun mot trouvé avec ces critères.")
                else: st.error("Aucun mot disponible pour cette sélection.")
                
    elif st.session_state.test_etape == "en_cours":
        st.subheader("📝 Feuille de test")
        with st.form("formulaire_test"):
            reponses_saisies = {}
            for i, (idx, mot) in enumerate(st.session_state.test_mots.iterrows()):
                st.write(f"**{i+1}. {mot[st.session_state.test_lang_dep]}** ({obtenir_picto_cat(mot['Categorie'])})")
                reponses_saisies[idx] = st.text_input(f"Rep {i}", label_visibility="collapsed", key=f"test_rep_{idx}")
                    
            soumis = st.form_submit_button("✅ Valider mon Test", type="primary", use_container_width=True)
            if soumis:
                note = 0.0; details = []
                for idx, mot in st.session_state.test_mots.iterrows():
                    rep_user = st.session_state[f"test_rep_{idx}"]
                    attendu = mot[st.session_state.test_lang_arr]
                    score = verifier_reponse(rep_user, attendu)
                    
                    if score == 100: note += 1; st.session_state.vocabulaire.at[idx, 'Score_100'] += 1
                    elif score == 50: note += 0.5; st.session_state.vocabulaire.at[idx, 'Score_50'] += 1
                    else: st.session_state.vocabulaire.at[idx, 'Score_0'] += 1
                        
                    st.session_state.vocabulaire.at[idx, 'Pourcentage'] = calculer_pourcentage(st.session_state.vocabulaire.loc[idx])
                    details.append({'question': mot[st.session_state.test_lang_dep], 'attendu': attendu, 'reponse': rep_user, 'score': score})
                    
                sauvegarder_base()
                st.session_state.test_details = details
                st.session_state.test_note = note
                st.session_state.test_etape = "resultats"
                st.rerun()

    elif st.session_state.test_etape == "resultats":
        total_mots = len(st.session_state.test_details)
        note = st.session_state.test_note
        pct = round((note / total_mots) * 100, 1)
        
        st.success("🎉 TEST TERMINÉ !")
        st.subheader(f"Note : {note} / {total_mots} ({pct}%)")
        st.write("---")
        
        for i, res in enumerate(st.session_state.test_details):
            st.write(f"**{i+1}. {res['question']}**")
            if res['score'] == 100: st.success(f"✅ {res['attendu']}")
            elif res['score'] == 50: st.warning(f"⚠️ Écrit : '{res['reponse']}' -> Exact : **{res['attendu']}**")
            else: st.error(f"❌ Écrit : '{res['reponse']}' -> Exact : **{res['attendu']}**")
            st.audio(generer_audio(res['attendu'], CODES_AUDIO[st.session_state.test_lang_arr]), format="audio/mp3")
            st.write("---")
                
        if st.button("🔄 Refaire un test", type="primary", use_container_width=True):
            st.session_state.test_etape = "config"
            st.rerun()

# ==========================================
# PAGE 3 : DICTIONNAIRE
# ==========================================
elif menu == "📖 Dict.":
    st.title("📖 Dictionnaire")
    st.session_state.test_etape = "config"
    st.session_state.entrainement_etat = "attente"
    
    col_l1, col_l2 = st.columns(2)
    with col_l1: lang_1 = st.selectbox("Langue 1 :", list(LANGUES_AFFICHAGE.values()), index=0)
    with col_l2: lang_2 = st.selectbox("Langue 2 :", list(LANGUES_AFFICHAGE.values()), index=1)
        
    code_l1 = INVERS_AFFICHAGE[lang_1]
    code_l2 = INVERS_AFFICHAGE[lang_2]
    
    df_dic = st.session_state.vocabulaire.dropna(subset=[code_l1, code_l2]).copy()
    
    recherche = st.text_input("🔍 Rechercher un mot :", "").strip().lower()
    categories_raw = sorted([str(c) for c in df_dic['Categorie'].unique() if pd.notna(c) and c != ""])
    cat_options = ["🔍 Toutes les catégories"] + [obtenir_picto_cat(c) for c in categories_raw]
    choix_cat_affichee = st.selectbox("📂 Filtrer par catégorie :", cat_options)
    
    if recherche:
        masque = df_dic[[code_l1, code_l2]].apply(lambda col: col.str.lower().str.contains(recherche, na=False)).any(axis=1)
        df_dic = df_dic[masque]
        
    if choix_cat_affichee != "🔍 Toutes les catégories":
        cat_pure = choix_cat_affichee.split(" ", 1)[1] if " " in choix_cat_affichee else choix_cat_affichee
        df_dic = df_dic[df_dic['Categorie'] == cat_pure]
        
    st.caption(f"**{len(df_dic)} mot(s) trouvé(s)**")
    
    df_dic = df_dic.sort_values(by=code_l1)
    df_dic['Categorie'] = df_dic['Categorie'].apply(obtenir_picto_cat)
    cols_a_afficher = ['Categorie', code_l1, code_l2, 'Pourcentage']
    df_final = df_dic[cols_a_afficher].rename(columns={code_l1: lang_1, code_l2: lang_2})
    
    st.dataframe(df_final, use_container_width=True, hide_index=True)

# ==========================================
# PAGE 4 : SCORES
# ==========================================
elif menu == "📊 Scores":
    st.title("📊 Mes Scores")
    st.session_state.test_etape = "config"
    st.session_state.entrainement_etat = "attente"
    
    st.write("### 🌐 Choisir les langues à consulter")
    col_s1, col_s2 = st.columns(2)
    with col_s1: s_lang1 = st.selectbox("Langue principale :", list(LANGUES_AFFICHAGE.values()), index=0, key="score_l1")
    with col_s2: s_lang2 = st.selectbox("Deuxième langue :", list(LANGUES_AFFICHAGE.values()), index=1, key="score_l2")
        
    code_s1 = INVERS_AFFICHAGE[s_lang1]
    code_s2 = INVERS_AFFICHAGE[s_lang2]
    
    df_scores = st.session_state.vocabulaire.dropna(subset=[code_s1, code_s2]).copy()
    
    st.write("---")
    st.write("### 📂 Progression par catégorie")
    cat_stats = df_scores.groupby('Categorie')['Pourcentage'].mean().round(1).reset_index()
    cat_stats = cat_stats.sort_values(by='Pourcentage', ascending=False)
    
    with st.expander("📊 Voir le détail par catégorie", expanded=True):
        for _, row in cat_stats.iterrows():
            c_nom = row['Categorie']
            c_pct = int(row['Pourcentage'])
            st.write(f"**{obtenir_picto_cat(c_nom)}** — **{c_pct}%**")
            st.progress(c_pct / 100)
    
    st.write("---")
    st.write("### 📋 Détail des mots")
    
    df_affichage = df_scores[['Categorie', code_s1, code_s2, 'Score_100', 'Score_50', 'Score_0', 'Pourcentage']].sort_values(by='Pourcentage', ascending=False)
    df_affichage['Categorie'] = df_affichage['Categorie'].apply(obtenir_picto_cat)
    df_affichage = df_affichage.rename(columns={
        code_s1: s_lang1, code_s2: s_lang2,
        'Score_100': '✅ Réussi', 'Score_50': '⚠️ Moyen', 'Score_0': '❌ Erreurs', 'Pourcentage': 'Score (%)'
    })
    
    st.dataframe(df_affichage, use_container_width=True, hide_index=True)

# ==========================================
# PAGE 5 : GESTION DES MOTS ET MISES À JOUR
# ==========================================
elif menu == "⚙️ Gérer":
    st.title("⚙️ Gérer ma base")
    st.session_state.test_etape = "config"
    st.session_state.entrainement_etat = "attente"
    
    # Bouton magique de mise à jour depuis l'Excel
    st.write("### 🔄 Mettre à jour depuis GitHub")
    st.write("Si tu as modifié ton fichier Excel sur ton ordinateur et que tu l'as mis sur GitHub, clique ici pour charger les nouveaux mots. *(Note : cela remet tes scores à zéro pour reprendre sur une base propre !)*")
    if st.button("⚠️ Recharger le fichier Excel", type="primary", use_container_width=True):
        fichier_profil = obtenir_nom_fichier(st.session_state.utilisateur_connecte)
        if os.path.exists(fichier_profil):
            os.remove(fichier_profil) # On supprime l'ancien historique
        st.session_state.vocabulaire = preparer_base_globale(st.session_state.utilisateur_connecte) # On recrée depuis l'Excel
        st.success("Ta base a été réinitialisée avec succès avec le nouveau fichier Excel !")
        st.rerun()

    st.write("---")
    
    with st.form("form_ajout", clear_on_submit=True):
        st.write("### ➕ Ajouter un mot manuellement")
        col1, col2 = st.columns(2)
        with col1: nv_francais = st.text_input("Français*")
        with col2: nv_anglais = st.text_input("Anglais*")
        
        cats = sorted([str(c) for c in st.session_state.vocabulaire['Categorie'].unique() if pd.notna(c) and c != ""])
        choix_cat = st.selectbox("Catégorie*", cats + ["✨ NOUVELLE..."])
        nv_categorie = st.text_input("Nom de la catégorie :") if choix_cat == "✨ NOUVELLE..." else choix_cat
            
        if st.form_submit_button("✅ Ajouter le mot", use_container_width=True):
            if nv_francais and nv_anglais and nv_categorie:
                nouveau_mot = pd.DataFrame([{
                    'Categorie': nv_categorie, 'Francais': nv_francais, 'Anglais': nv_anglais, 
                    'Espagnol': None, 'Italien': None, 'Allemand': None,
                    'Score_100': 0, 'Score_50': 0, 'Score_0': 0, 'Pourcentage': 0.0
                }])
                st.session_state.vocabulaire = pd.concat([st.session_state.vocabulaire, nouveau_mot], ignore_index=True)
                sauvegarder_base()
                st.success("Mot ajouté !")
                st.rerun()
    
    st.write("---")
    st.write("### ✏️ Éditer la base")
    voir_tout = st.checkbox("🔍 Afficher toutes les langues (Espagnol, Italien, Allemand)", value=False)
    cols_edition = ['Categorie'] + LANGUES_COLS + ['Pourcentage'] if voir_tout else ['Categorie', 'Francais', 'Anglais', 'Pourcentage']
        
    df_modifie = st.data_editor(st.session_state.vocabulaire, column_order=cols_edition, num_rows="dynamic", use_container_width=True, hide_index=True)
    
    if st.button("💾 Enregistrer les modifications manuelles", use_container_width=True):
        st.session_state.vocabulaire = df_modifie
        sauvegarder_base()
        st.success("Modifications enregistrées !")