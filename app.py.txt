import streamlit as st
import pandas as pd
import re
import unicodedata

# Configuration de la page Streamlit
st.set_page_config(page_title="Check Reporting SAC vs BI", layout="wide", page_icon="📊")

pd.set_option('future.no_silent_downcasting', True)

# --- FONCTIONS DE NETTOYAGE (Version Originale et Stable) ---
def safe_float_conversion(val):
    if isinstance(val, pd.Series): val = val.iloc[0]
    s = str(val).strip()
    if s == '' or s.lower() in ['nan', 'nat', 'none', 'nc', '<na>']: return 0.0
    cleaned = re.sub(r'[^\d,.-]', '', s).replace(',', '.')
    if cleaned in ['', '.', '-']: return 0.0
    try: return float(cleaned)
    except ValueError: return 0.0

def clean_excel_text(text):
    if not isinstance(text, str): return text
    # Cette ligne est vitale : elle retire les puces ET les retours à la ligne cachés !
    return re.sub(r'[-•=●◦○▪\x00-\x1F\x7F]', '', text).strip()

def remove_accents(input_str):
    if not isinstance(input_str, str): return ""
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return u"".join([c for c in nfkd_form if not unicodedata.combining(c)])

def generate_join_key(enseigne, rayon):
    enseigne_clean = remove_accents(str(enseigne)).upper()
    enseigne_compact = re.sub(r'[^A-Z0-9]', '', enseigne_clean)

    # Anti-Collision des Enseignes
    if "JEANCOUTU" in enseigne_compact: e = "JEANCO"
    elif "REDAPPLE" in enseigne_compact: e = "REDAPP"
    elif "BRUNET" in enseigne_compact: e = "BRUNET"
    elif "GAPOUTLET" in enseigne_compact: e = "GAPOUT"
    elif "GAP" in enseigne_compact: e = "GAP"
    elif "CORNERADF" in enseigne_compact: e = "CORNADF"
    elif "CORNERLOLLIPOPS" in enseigne_compact: e = "CORNLOL"
    elif "INTERMARCHEMOA" in enseigne_compact: e = "INTMOA"
    elif "INTERMARCHE" in enseigne_compact: e = "INTERM"
    elif "GHANTYMOA" in enseigne_compact: e = "GHAMOA"
    elif "GHANTY" in enseigne_compact: e = "GHANTY"
    else:
        words = re.findall(r'[A-Z0-9]+', enseigne_clean)
        e = "UNKNOWN"
        if words:
            e = words[0]
            if len(e) <= 3 and len(words) > 1: e += words[1]
        e = e[:6]

    r = re.sub(r'[^A-Z0-9]', '', remove_accents(str(rayon)).upper())
    return f"{e}_{r}"

def find_actual_header_sac(df):
    for i in range(len(df)):
        row = df.iloc[i].astype(str).values
        if any('Rayon' in str(val) for val in row):
            new_cols = [clean_excel_text(str(val)) if str(val) != 'nan' else f'COL_{j}' for j, val in enumerate(row)]
            df.columns = new_cols
            df_clean = df.iloc[i+1:].reset_index(drop=True)
            return df_clean.loc[:, ~df_clean.columns.duplicated()].copy()
    return df

# --- MOTEUR DE RÈGLES DYNAMIQUE ---
def apply_business_rules(df, col_enseigne, report_type):
    df['Rayon'] = df['Rayon'].astype(str)

    def apply_rule(enseignes, rayons, new_rayon):
        ens_regex = '(' + '|'.join([re.escape(e.lower()) for e in enseignes]) + ')'
        ray_regex = '^(' + '|'.join([re.escape(r.lower().strip()) for r in rayons]) + ')$'
        mask = df[col_enseigne].astype(str).str.lower().str.contains(ens_regex, regex=True, na=False) & \
               df['Rayon'].str.lower().str.contains(ray_regex, regex=True, na=False)
        df.loc[mask, 'Rayon'] = new_rayon

    # ==========================================
    # 🟠 RÈGLES POUR LE REPORTING : DIFFUSION
    # ==========================================
    if report_type == "Diffusion":
        apply_rule(['El Corte'], ['ParCaiss', 'Cheveux'], 'Bijoux')
        apply_rule(['Conbipel'], ['Capsules'], 'Bijoux')
        apply_rule(['Upim'], ['Access', 'Cheveux', 'Lunettes'], 'Bijoux')
        apply_rule(['Grain de Malice'], ['Access'], 'Bijoux')
        apply_rule(['Hema'], ['Access', 'Actua', 'Mariage'], 'Lunettes')
        apply_rule(['GAP', 'GAP Outlet'], ['Access'], 'Bijoux')
        apply_rule(['Auchan'], ['Enfant'], 'Bijoux')
        apply_rule(['BZB'], ['Cheveux'], 'Bijoux')
        apply_rule(['Cache Cache'], ['Cheveux'], 'Bijoux')
        apply_rule(['Zebra'], ['Cheveux'], 'Bijoux')

        apply_rule(['Retail'], ['Beaute', 'Montres'], 'Access')
        apply_rule(['Lollipops Retail'], ['Montres'], 'Access')
        apply_rule(['El Corte'], ['Montres', 'Beaute'], 'Access')
        apply_rule(['GL', 'Galeries Lafayette'], ['Papeteri', 'Parfums'], 'Access')
        apply_rule(['Intermarché MOA', 'MOA'], ['ParCaiss'], 'Access')

        apply_rule(['GAP', 'GAP Outlet'], ['Acier', 'Deco'], 'Capsules')
        apply_rule(['GAP', 'GAP Outlet'], ['Catalog'], 'Enfant')
        apply_rule(['Corner ADF'], ['Access'], 'Deco')

        communes_map = {
            'ArGarant': 'Bijoux', 'Sacherie': 'Bijoux', 'Catalog': 'Lunettes de vue',
            'Papeteri': 'Cartes de voeux', 'Chaussur': 'Access', 'Mariage': 'Bijoux',
            'Actua': 'Bijoux', 'Montres': 'Bijoux', 'ParCaiss': 'Parcours Caisse', 'PARCAISS': 'Parcours Caisse'
        }
        for old_val, new_val in communes_map.items():
            df.loc[df['Rayon'].str.lower().str.contains(f'^{re.escape(old_val.lower())}$', regex=True, na=False), 'Rayon'] = new_val

    # ==========================================
    # 🔵 RÈGLES POUR : BROTHERS / ACCESSORIES USA / ACCESSORIES CANADA
    # ==========================================
    elif report_type in ["Brothers", "Accessories USA", "Accessories Canada"]:
        apply_rule(['GAP', 'GAP Outlet'], ['Acier', 'Deco'], 'Capsules')
        apply_rule(['GAP', 'GAP Outlet'], ['Access'], 'Bijoux')
        apply_rule(['GAP', 'GAP Outlet'], ['Catalog'], 'Enfant')

        communes_map_brothers = {
            'Sacherie': 'Bijoux', 'Catalog': 'Lunettes de vue', 'Papeteri': 'Cartes de voeux',
            'Chaussur': 'Access', 'Mariage': 'Bijoux', 'Montres': 'Bijoux',
            'ParCaiss': 'Parcours Caisse', 'PARCAISS': 'Parcours Caisse'
        }
        for old_val, new_val in communes_map_brothers.items():
            df.loc[df['Rayon'].str.lower().str.contains(f'^{re.escape(old_val.lower())}$', regex=True, na=False), 'Rayon'] = new_val

    return df

# --- CHARGEMENT POWER BI ---
def load_powerbi_data(file_buffer, target_col):
    try: df_pbi = pd.read_excel(file_buffer, header=1)
    except:
        try: df_pbi = pd.read_excel(file_buffer)
        except: return pd.DataFrame()

    if target_col not in df_pbi.columns: return pd.DataFrame()

    mask_total = df_pbi['-Rayon'].astype(str).str.strip().str.lower() == 'total'
    df_pbi.loc[mask_total, target_col] = df_pbi[target_col].shift(1)
    df_pbi[target_col] = pd.to_numeric(df_pbi[target_col], errors='coerce')
    df_pbi = df_pbi.dropna(subset=[target_col, '-Centrale', '-Rayon']).copy()
    df_pbi = df_pbi[~df_pbi['-Centrale'].astype(str).str.strip().str.lower().isin(['total', 'nan', 'total général'])]
    df_pbi = df_pbi[~df_pbi['-Rayon'].astype(str).str.strip().str.lower().isin(['nan', 'none', ''])]

    df_pbi['-Centrale'] = df_pbi['-Centrale'].apply(clean_excel_text)
    df_pbi['-Rayon'] = df_pbi['-Rayon'].astype(str).apply(clean_excel_text)

    if 'Magasin comparable' in df_pbi.columns:
        df_pbi = df_pbi[~df_pbi['Magasin comparable'].astype(str).str.contains('Dont mag comp', case=False, na=False)]
    df_pbi.loc[df_pbi['-Rayon'].str.lower() == 'total', '-Rayon'] = 'TOTAL'

    # Exclusion Giant Tiger
    df_pbi = df_pbi[~df_pbi['-Centrale'].astype(str).str.contains('Giant Tiger', case=False, na=False)]

    df_pbi['Join_Key'] = df_pbi.apply(lambda x: generate_join_key(x['-Centrale'], x['-Rayon']), axis=1).astype(str)
    df_pbi[target_col] = df_pbi[target_col].apply(safe_float_conversion)
    df_pbi = df_pbi.drop_duplicates('Join_Key', keep='first')
    return df_pbi[['Join_Key', '-Centrale', '-Rayon', target_col]]

# --- INTERFACE STREAMLIT ---
st.title("📊 Réconciliation SAC vs Power BI")

report_type = st.selectbox(
    "1️⃣ Choisir le type de reporting",
    ["Diffusion", "Brothers", "Accessories USA", "Accessories Canada"]
)

st.markdown("---")
st.write("2️⃣ Importer les fichiers correspondants")

col1, col2 = st.columns(2)
with col1:
    file_sac = st.file_uploader("Fichier SAC (Check_Reporting.xlsx)", type=['xlsx'])
    file_ann = st.file_uploader("Export BI (Annuel)", type=['xlsx', 'csv'])
with col2:
    file_mens = st.file_uploader("Export BI (Mensuel)", type=['xlsx', 'csv'])
    file_sem = st.file_uploader("Export BI (Semaines)", type=['xlsx', 'csv'])

if st.button("🚀 Lancer l'Analyse", type="primary"):
    if not file_sac:
        st.error("Veuillez importer au moins le fichier SAC (Check_Reporting.xlsx).")
    else:
        with st.spinner('Analyse des données en cours...'):
            try:
                xl_sac = pd.ExcelFile(file_sac)
                all_recaps = {}

                TESTS_MAPPING = {
                    "Annuel 2026": {"onglet_sac": "Cumul 2026", "file": file_ann, "col_valeur": "N", "mode": "Annuel"},
                    "Annuel 2025": {"onglet_sac": "Cumul 2025", "file": file_ann, "col_valeur": "N-1", "mode": "Annuel"},
                    "Mensuel 2026": {"onglet_sac": "Cumul mois 2026", "file": file_mens, "col_valeur": "N", "mode": "Mensuel"},
                    "Mensuel 2025": {"onglet_sac": "Cumul mois 2025", "file": file_mens, "col_valeur": "N-1", "mode": "Mensuel"},
                    "Semaine Der 2026": {"onglet_sac": "Cumul mois 2026", "file": file_sem, "col_valeur": "N", "mode": "Semaine_Derniere"},
                    "Semaine Der 2025": {"onglet_sac": "Cumul mois 2025", "file": file_sem, "col_valeur": "N-1", "mode": "Semaine_Derniere"},
                    "Semaine Avant-Der 2026": {"onglet_sac": "Cumul mois 2026", "file": file_sem, "col_valeur": "N.1", "mode": "Semaine_Avant_Derniere"},
                    "Semaine Avant-Der 2025": {"onglet_sac": "Cumul mois 2025", "file": file_sem, "col_valeur": "N-1.1", "mode": "Semaine_Avant_Derniere"}
                }

                for nom_test, conf in TESTS_MAPPING.items():
                    if conf['file'] is None: continue

                    df_pbi = load_powerbi_data(conf['file'], conf['col_valeur'])
                    if df_pbi.empty: continue

                    try: df_raw = xl_sac.parse(conf['onglet_sac'])
                    except ValueError: continue

                    df_sac = find_actual_header_sac(df_raw)

                    col_enseigne = None
                    for c in df_sac.columns:
                        str_c = str(c).lower()
                        if col_enseigne is None and ('enseigne' in str_c or 'centrale' in str_c or 'client' in str_c): col_enseigne = c
                    if not col_enseigne: continue

                    # Exclusion Giant Tiger (SAC)
                    df_sac = df_sac[~df_sac[col_enseigne].astype(str).str.contains('Giant Tiger', case=False, na=False)]

                    col_val = None
                    if conf['mode'] == "Mensuel":
                        cols_to_sum = df_sac.columns[3:8]
                        for c in cols_to_sum: df_sac[c] = df_sac[c].apply(safe_float_conversion)
                        col_val = 'Valeur_Cumulee_D_a_H'
                        df_sac[col_val] = df_sac[cols_to_sum].sum(axis=1)
                    elif conf['mode'] in ["Semaine_Derniere", "Semaine_Avant_Derniere"]:
                        possible_cols = df_sac.columns[3:8]
                        valid_cols = [c for c in possible_cols if df_sac[c].apply(safe_float_conversion).abs().sum() > 0]
                        if len(valid_cols) < 2: continue
                        col_val = valid_cols[-1] if conf['mode'] == "Semaine_Derniere" else valid_cols[-2]
                        df_sac[col_val] = df_sac[col_val].apply(safe_float_conversion)
                    else:
                        mots_cles = ['va net', 'ttc', 'ca ', 'chiffre', 'valeur', 'réalisé', 'realise']
                        for c in df_sac.columns:
                            if col_val is None and any(mot in str(c).lower() for mot in mots_cles): col_val = c
                        if not col_val:
                            try:
                                idx_rayon = list(df_sac.columns).index('Rayon')
                                for c in df_sac.columns[idx_rayon+1:]:
                                    if df_sac[c].apply(safe_float_conversion).abs().sum() > 0:
                                        col_val = c
                                        break
                            except: pass
                        if not col_val: continue
                        df_sac[col_val] = df_sac[col_val].apply(safe_float_conversion)

                    # DÉTECTION DU PAYS (Méthode robuste originale)
                    col_pays = None
                    max_matches = 0
                    pattern_pays_detect = 'France|Guadeloupe|Guyane|Maurice|Réunion|Reunion|Martinique|Calédonie|Caledonie|Malte|Saint[- ]Martin|Royaume[- ]Uni|Barthelemy|Barthélemy|Luxembourg|USA|Etats|Canada|CAN|US'

                    for c in df_sac.columns:
                        if str(df_sac[c].dtype) in ['float64', 'int64']: continue
                        matches = df_sac[c].astype(str).str.contains(pattern_pays_detect, case=False, na=False, regex=True).sum()
                        if matches > max_matches:
                            max_matches = matches
                            col_pays = c

                    if not col_pays or max_matches < 5:
                        col_pays = None

                    df_sac = apply_business_rules(df_sac, col_enseigne, report_type)

                    if report_type == "Diffusion":
                        df_sac.loc[df_sac[col_enseigne].astype(str).str.strip().str.upper() == 'GL', col_enseigne] = 'Galeries Lafayette'

                    # FILTRES GÉOGRAPHIQUES PROPRES
                    if col_pays:
                        if report_type == "Diffusion":
                            pattern_pays = 'France|Guadeloupe|Guyane|Maurice|Réunion|Reunion|Martinique|Calédonie|Caledonie|Malte|Saint[- ]Martin|Royaume[- ]Uni|Barthelemy|Barthélemy|Luxembourg'
                            cond_globale = df_sac[col_pays].astype(str).str.contains(pattern_pays, case=False, na=False, regex=True)

                            cond_belgique_ok = df_sac[col_enseigne].astype(str).str.contains('Besson|Paprika|Monoprix|Morgan|Pimkie|Promod', case=False, na=False) & df_sac[col_pays].astype(str).str.contains('Belgique', case=False, na=False)
                            cond_celio_interdit = df_sac[col_enseigne].astype(str).str.contains('Celio', case=False, na=False) & ~df_sac[col_pays].astype(str).str.contains('France', case=False, na=False)
                            cond_hema_interdit = df_sac[col_enseigne].astype(str).str.contains('Hema', case=False, na=False) & ~df_sac[col_pays].astype(str).str.contains('France', case=False, na=False)

                            df_sac = df_sac[(cond_globale | cond_belgique_ok) & ~cond_celio_interdit & ~cond_hema_interdit]

                            mask_maurice = df_sac[col_pays].astype(str).str.contains('Maurice', case=False, na=False)
                            if mask_maurice.sum() > 0: df_sac.loc[mask_maurice, col_val] *= 0.019

                        elif report_type in ["Brothers", "Accessories USA"]:
                            mask_usa = df_sac[col_pays].astype(str).str.contains('USA|Etats-Unis|United States|US', case=False, na=False, regex=True)
                            df_sac = df_sac[mask_usa]

                        elif report_type == "Accessories Canada":
                            mask_canada = df_sac[col_pays].astype(str).str.contains('Canada|CAN', case=False, na=False, regex=True)
                            mask_100_canada = df_sac[col_enseigne].astype(str).str.contains('Jean Coutu|Brunet|Red Apple', case=False, na=False)
                            df_sac = df_sac[mask_canada | mask_100_canada]

                    if not df_sac.empty:
                        df_sac['Join_Key'] = df_sac.apply(lambda x: generate_join_key(x[col_enseigne], x['Rayon']), axis=1).astype(str)
                        df_sac_rayons = df_sac.groupby('Join_Key', as_index=False)[col_val].sum()
                        df_sac['Centrale_Key'] = df_sac['Join_Key'].apply(lambda x: x.split('_')[0])
                        df_sac_total = df_sac.groupby('Centrale_Key', as_index=False)[col_val].sum()
                        df_sac_total['Join_Key'] = df_sac_total['Centrale_Key'] + "_TOTAL"
                        df_sac = pd.concat([df_sac_rayons, df_sac_total[['Join_Key', col_val]]], ignore_index=True)
                        df_sac = df_sac.groupby('Join_Key', as_index=False)[col_val].sum()
                    else:
                        df_sac = pd.DataFrame(columns=['Join_Key', col_val])

                    merged = df_pbi.merge(df_sac, on='Join_Key', how='left')
                    merged[conf['col_valeur']] = merged[conf['col_valeur']].fillna(0.0)
                    merged[col_val] = merged[col_val].fillna(0.0)
                    merged['Ecart'] = (merged[col_val] - merged[conf['col_valeur']]).round(2)
                    merged['Statut'] = merged['Ecart'].apply(lambda x: '❌ Anomalie' if abs(x) > 1 else '✅ OK')

                    colonnes_export = ['-Centrale', '-Rayon', col_val, conf['col_valeur'], 'Ecart', 'Statut']
                    recap_complet = merged[colonnes_export].rename(columns={
                        '-Centrale': 'Centrale', '-Rayon': 'Rayon',
                        col_val: 'Valeur SAC', conf['col_valeur']: 'Valeur Power BI',
                        'Statut': 'Statut'
                    })

                    recap_complet = recap_complet.sort_values(by='Statut', ascending=True)
                    all_recaps[nom_test] = recap_complet

                if all_recaps:
                    st.success("Analyse terminée avec succès !")
                    onglets = st.tabs(list(all_recaps.keys()))
                    for tab, (nom, df) in zip(onglets, all_recaps.items()):
                        with tab:
                            nb_anomalies = len(df[df['Statut'] == '❌ Anomalie'])
                            if nb_anomalies > 0:
                                st.error(f"{nb_anomalies} anomalie(s) détectée(s) sur cet onglet.")
                            else:
                                st.success("Tout est OK ! ✅")
                            st.dataframe(df, use_container_width=True, hide_index=True)
                else:
                    st.warning("Aucun résultat n'a pu être généré. Vérifiez les fichiers importés et les correspondances.")

            except Exception as e:
                st.error(f"Une erreur inattendue est survenue : {e}")