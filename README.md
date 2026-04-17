# 📊 Outils de Réconciliation SAC / PREV vs Power BI

Application Streamlit multi-pages pour contrôler automatiquement les écarts
entre les données SAC / Objectifs PREV et les exports Power BI.

## 🧭 Contenu

L'application comprend **deux outils** accessibles depuis une page d'accueil :

### 📊 Réconciliation SAC vs Power BI
Contrôle des écarts entre le fichier SAC et les exports Power BI :
- Annuel 2026 / 2025 (colonnes N / N-1)
- Mensuel 2026 / 2025
- Semaine Dernière 2026 / 2025
- Semaine Avant-Dernière 2026 / 2025

### 🎯 Réconciliation SAC / PREV vs Power BI
Version étendue qui **ajoute** le contrôle des objectifs PREV :
- Les 6 tests SAC classiques (N / N-1)
- 🎯 Objectif Annuel — cumul janvier → mois détecté du fichier PREV vs colonne `Objectif` du BI Annuel
- 🎯 Objectif Mensuel — colonne du mois détecté du PREV vs colonne `Objectif` du BI Mensuel

Le mois en cours est détecté automatiquement dans la **1ère ligne du fichier BI**
(ex: `au 31/03/2026` → mars).

## 🚀 Démarrage local

```bash
pip install -r requirements.txt
streamlit run Home.py
```

L'application s'ouvre sur la page d'accueil où tu peux choisir l'outil à lancer.

## ☁️ Déploiement Streamlit Cloud

1. Pousser ce dépôt sur GitHub
2. Sur [share.streamlit.io](https://share.streamlit.io) : nouvelle app
3. Sélectionner le repo + branche
4. **Main file path** : `Home.py`
5. Deploy 🚀

## 📁 Structure du dépôt

```
sac-bi-reconciliation/
├── Home.py                              ← Page d'accueil (entrée principale)
├── pages/
│   ├── 1_📊_Reconciliation_SAC.py       ← Application SAC
│   └── 2_🎯_Reconciliation_Objectifs.py ← Application SAC + Objectifs PREV
├── requirements.txt
└── README.md
```

## 📥 Fichiers attendus

### Pour l'application SAC
- `Check_Reporting.xlsx` — fichier SAC avec les onglets *Cumul 2026*, *Cumul 2025*,
  *Cumul mois 2026*, *Cumul mois 2025 réel*, *Cumul mois 2025*
- Export BI Annuel (colonnes `-Centrale`, `-Rayon`, `N`, `N-1`)
- Export BI Mensuel (colonnes `-Centrale`, `-Rayon`, `N`, `N-1`)
- Export BI Semaines (colonnes `-Centrale`, `-Rayon`, `N`, `N-1`)

### Pour l'application Objectifs (en plus du SAC)
- Fichier **PREV** (`.xlsx` ou `.csv`) avec colonnes `Enseigne Client`, `Rayon`,
  `Pays`, `janvier`…`décembre`
- Export BI Annuel contenant **aussi** la colonne `Objectif`
- Export BI Mensuel contenant **aussi** la colonne `Objectif`

## ⚙️ Types de reporting

Chaque application propose 4 modes qui activent leurs propres règles métier et
filtres géographiques :

- **Diffusion** — France + DOM-TOM + Europe, Maurice ×0.019, règles Belgique/Celio/Hema
- **Brothers / Accessories USA** — ciblage États-Unis
- **Accessories Canada** — Canada + enseignes 100% canadiennes (Jean Coutu, Brunet, Red Apple)

## 🔍 Fonctionnalités transverses

- **Détection automatique** du mois en cours, des en-têtes, des cellules fusionnées
- **Convertisseur de nombres indestructible** (formats EU `1.234,56` et US `1,234.56`)
- **Clés de jointure anti-collision** (JEANCO ≠ GAPOUT ≠ REDAPP, etc.)
- **Tolérance d'écart** : un écart > 1 € est classé anomalie
- **Export Excel** de chaque onglet avec formatage conditionnel (lignes anomalies en rouge)
- **Recherche + filtre anomalies** sur chaque tableau
- **Progress bar** et diagnostics détaillés par test

## 📝 Notes

Le dépôt utilise le système natif de **multi-pages** de Streamlit. Tout fichier `.py`
placé dans `pages/` devient automatiquement une page accessible via le menu latéral.
Le préfixe numérique (`1_`, `2_`) définit l'ordre d'apparition.
