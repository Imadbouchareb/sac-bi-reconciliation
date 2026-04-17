import streamlit as st

# ============================================================
#                    CONFIGURATION DE PAGE
# ============================================================
st.set_page_config(
    page_title="Accueil · Réconciliation BI",
    layout="wide",
    page_icon="🏠",
    initial_sidebar_state="expanded"
)

# ============================================================
#                       STYLE CSS CUSTOM
# ============================================================
st.markdown("""
<style>
    .home-title {
        font-size: 2.6rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.25rem;
        letter-spacing: -0.02em;
        text-align: center;
        margin-top: 1rem;
    }
    .home-subtitle {
        color: #64748b;
        font-size: 1.05rem;
        margin-bottom: 3rem;
        text-align: center;
    }

    .app-card {
        background: #ffffff;
        padding: 2rem 1.8rem;
        border-radius: 16px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        height: 100%;
        transition: all 0.2s ease;
    }
    .app-card:hover {
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        border-color: #cbd5e1;
    }
    .app-card-sac    { border-top: 4px solid #0284c7; }
    .app-card-obj    { border-top: 4px solid #8b5cf6; }

    .app-icon {
        font-size: 3rem;
        margin-bottom: 0.5rem;
    }
    .app-title {
        font-size: 1.5rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.5rem;
    }
    .app-desc {
        color: #475569;
        font-size: 0.95rem;
        line-height: 1.5;
        margin-bottom: 1rem;
    }
    .app-features {
        color: #64748b;
        font-size: 0.85rem;
        line-height: 1.7;
        margin-bottom: 1.5rem;
    }
    .app-features li {
        margin-bottom: 0.25rem;
    }

    .badge {
        display: inline-block;
        padding: 0.25rem 0.6rem;
        border-radius: 6px;
        font-size: 0.72rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-bottom: 1rem;
    }
    .badge-sac { background: #e0f2fe; color: #075985; }
    .badge-obj { background: #ede9fe; color: #5b21b6; }

    .footer-info {
        margin-top: 3rem;
        padding: 1.5rem;
        background: #f8fafc;
        border-radius: 12px;
        border: 1px solid #e2e8f0;
        color: #475569;
        font-size: 0.88rem;
    }
    .footer-info b { color: #0f172a; }
</style>
""", unsafe_allow_html=True)


# ============================================================
#                      CONTENU DE LA PAGE
# ============================================================
st.markdown('<div class="home-title">📊 Outils de Réconciliation BI</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="home-subtitle">Choisis l\'application à lancer selon ton besoin de contrôle</div>',
    unsafe_allow_html=True
)

col1, col2 = st.columns(2, gap="large")

# ────────────────────────────────────────────────────────
#                   APPLICATION 1 — SAC
# ────────────────────────────────────────────────────────
with col1:
    st.markdown("""
    <div class="app-card app-card-sac">
        <div class="app-icon">📊</div>
        <div class="badge badge-sac">Contrôle des ventes</div>
        <div class="app-title">Réconciliation SAC vs Power BI</div>
        <div class="app-desc">
            Contrôle des écarts entre le fichier SAC et les exports Power BI
            sur les colonnes N (année en cours) et N-1 (année précédente).
        </div>
        <div class="app-features">
            <b>8 tests automatisés :</b>
            <ul>
                <li>Annuel 2026 / 2025</li>
                <li>Mensuel 2026 / 2025</li>
                <li>Semaine Dernière 2026 / 2025</li>
                <li>Semaine Avant-Dernière 2026 / 2025</li>
            </ul>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚀 Ouvrir l'application SAC",
                 type="primary", use_container_width=True, key="btn_sac"):
        st.switch_page("pages/1_📊_Reconciliation_SAC.py")


# ────────────────────────────────────────────────────────
#              APPLICATION 2 — OBJECTIFS PREV
# ────────────────────────────────────────────────────────
with col2:
    st.markdown("""
    <div class="app-card app-card-obj">
        <div class="app-icon">🎯</div>
        <div class="badge badge-obj">Contrôle complet + objectifs</div>
        <div class="app-title">Réconciliation SAC / PREV vs Power BI</div>
        <div class="app-desc">
            Version étendue qui ajoute le contrôle des objectifs PREV en plus
            des tests SAC classiques. Détection automatique du mois via la 1ère
            ligne du fichier BI.
        </div>
        <div class="app-features">
            <b>8 tests (6 SAC + 2 Objectifs) :</b>
            <ul>
                <li>Annuel 2026 / 2025</li>
                <li>Mensuel 2026 / 2025</li>
                <li>Semaine Dernière 2026 / 2025</li>
                <li>🎯 Objectif Annuel (cumul janvier → mois)</li>
                <li>🎯 Objectif Mensuel (colonne du mois)</li>
            </ul>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚀 Ouvrir l'application Objectifs",
                 type="primary", use_container_width=True, key="btn_obj"):
        st.switch_page("pages/2_🎯_Reconciliation_Objectifs.py")


# ────────────────────────────────────────────────────────
#                       FOOTER
# ────────────────────────────────────────────────────────
st.markdown("""
<div class="footer-info">
    💡 <b>Astuce :</b> Tu peux aussi naviguer directement entre les applications via
    le menu dans la barre latérale gauche. Chaque application conserve ses propres
    fichiers et résultats tant que la session est active.
</div>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown("### 🧭 Navigation")
    st.info(
        "Utilise le menu ci-dessus pour naviguer entre la page d'accueil "
        "et les deux applications de réconciliation."
    )
    st.markdown("---")
    st.caption("Outils de contrôle SAC / PREV vs Power BI")
