"""
app.py — Interface Streamlit pour le générateur de propale Neoma Conseil
"""

import streamlit as st
import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from datetime import datetime

# ── Questionnaire guidé ───────────────────────────────────────
QUESTIONNAIRE_SECTIONS = [
    {
        "titre": "🏢 Informations sur l'entreprise cliente",
        "questions": [
            {"key": "nom_entreprise",    "label": "Nom de l'entreprise",                        "type": "text",     "placeholder": "Ex : Carrefour, BNP Paribas…"},
            {"key": "adresse_ville",     "label": "Adresse et ville",                           "type": "text",     "placeholder": "Ex : 10 rue de la Paix, Paris"},
            {"key": "taille_entreprise", "label": "Taille de l'entreprise",                     "type": "select",   "options": ["— Sélectionner —", "TPE (< 10 salariés)", "PME (10–250 salariés)", "ETI (250–5 000 salariés)", "Grande entreprise (> 5 000 salariés)", "Association / ONG", "Startup"]},
            {"key": "nom_client",        "label": "Nom et prénom du représentant / contact",    "type": "text",     "placeholder": "Ex : Marie Dupont"},
            {"key": "tel_mail_client",   "label": "Téléphone et email du client",               "type": "text",     "placeholder": "Ex : 06 12 34 56 78 / marie@client.fr"},
        ],
    },
    {
        "titre": "🎯 Le projet et ses objectifs",
        "questions": [
            {"key": "titre_etude",        "label": "Nom / titre de l'étude",                                               "type": "text",     "placeholder": "Ex : Étude de marché — Restauration rapide"},
            {"key": "type_etude",         "label": "Type d'étude",                                                         "type": "select",   "options": ["— Sélectionner —", "Étude de marché", "Étude consommateurs", "Étude de satisfaction", "Étude de faisabilité", "Étude de notoriété", "Étude concurrentielle", "Étude de communication"]},
            {"key": "contexte_enjeux",    "label": "Contexte et enjeux — pourquoi ce projet maintenant ?",                 "type": "textarea", "placeholder": "Ex : L'entreprise souhaite lancer un nouveau produit et veut valider la demande…"},
            {"key": "objectif_principal", "label": "Objectif principal en une phrase",                                     "type": "text",     "placeholder": "Ex : Évaluer la demande pour un service de livraison bio"},
            {"key": "sous_objectifs",     "label": "Les 3 sous-objectifs spécifiques",                                     "type": "textarea", "placeholder": "1. …\n2. …\n3. …"},
            {"key": "resultats_attendus", "label": "Ce que le client veut accomplir grâce aux résultats",                  "type": "textarea", "placeholder": "Décisions à prendre, problème à résoudre, valeur attendue…"},
        ],
    },
    {
        "titre": "🛍️ Produit ou service étudié",
        "questions": [
            {"key": "produit_service", "label": "Nom propre du produit ou service",  "type": "text", "placeholder": "Ex : Nike Air Max, Netflix, application MyFitness"},
            {"key": "secteur",         "label": "Secteur d'activité",                "type": "text", "placeholder": "Ex : Restauration rapide, cosmétique bio, fintech…"},
        ],
    },
    {
        "titre": "📍 Cible et zone géographique",
        "questions": [
            {"key": "cible",      "label": "Description de la cible visée",                    "type": "text", "placeholder": "Ex : Consommateurs 25–45 ans, résidents Île-de-France"},
            {"key": "zone_geo",   "label": "Zone géographique de l'étude",                     "type": "text", "placeholder": "Ex : Région Normandie, Paris et petite couronne"},
            {"key": "villes_rep", "label": "Villes représentatives (si pertinent)",            "type": "text", "placeholder": "Ex : Rouen, Caen, Le Havre"},
            {"key": "nb_villes",  "label": "Nombre de zones / villes à couvrir",               "type": "text", "placeholder": "Ex : 3"},
        ],
    },
    {
        "titre": "🔬 Méthodologie",
        "questions": [
            {"key": "nb_questionnaires",  "label": "Nombre de questionnaires à administrer",   "type": "text",     "placeholder": "Ex : 150, ou 100 à 200 (sinon le DET le donne)"},
            {"key": "nb_intervenants",    "label": "Nombre d'intervenants prévus",              "type": "text",     "placeholder": "Ex : 5 (sinon le DET le donne)"},
            {"key": "mode_passation",     "label": "Mode de passation",                         "type": "select",   "options": ["— Sélectionner —", "Présentiel", "Téléphone", "En ligne", "Mixte (présentiel + en ligne)", "Mixte (présentiel + téléphone)"]},
            {"key": "duree_questionnaire","label": "Durée estimée du questionnaire",            "type": "text",     "placeholder": "Ex : 5 à 7 minutes"},
            {"key": "esd_details",        "label": "Entretiens semi-directifs (ESD) — si applicable",  "type": "textarea", "placeholder": "Nombre d'intervenants, objectifs… Laisser vide si non applicable."},
            {"key": "visite_mystere_det", "label": "Visites mystères — si applicable",                 "type": "textarea", "placeholder": "Objectif majeur, nombre de visites… Laisser vide si non applicable."},
        ],
    },
    {
        "titre": "⚔️ Concurrence et marché",
        "questions": [
            {"key": "concurrents_directs",   "label": "Concurrents directs connus",            "type": "text", "placeholder": "Ex : Leclerc, Aldi (sinon l'IA les déduit)"},
            {"key": "concurrents_indirects", "label": "Concurrents indirects connus",          "type": "text", "placeholder": "Ex : Amazon, vente en ligne (sinon l'IA les déduit)"},
            {"key": "marche_potentiel",      "label": "Estimation du marché potentiel",        "type": "text", "placeholder": "Ex : 2 millions de clients potentiels en IDF"},
        ],
    },
    {
        "titre": "📣 Communication (si étude de communication)",
        "questions": [
            {"key": "canaux_com",  "label": "Canaux de communication envisagés",   "type": "text",     "placeholder": "Ex : Instagram, presse locale, email marketing"},
            {"key": "actions_com", "label": "Actions de communication prévues",    "type": "textarea", "placeholder": "Ex : Campagne réseaux sociaux, relations presse…"},
            {"key": "cible_com",   "label": "Cible de communication",              "type": "text",     "placeholder": "Ex : 18–35 ans, urbains, CSP+"},
        ],
    },
    {
        "titre": "👥 L'équipe Neoma Conseil",
        "questions": [
            {"key": "chef_projet", "label": "Nom du chef de projet",                               "type": "text", "placeholder": "Ex : Thomas Martin"},
            {"key": "validateur",  "label": "Validateur / second chef de projet (si applicable)",  "type": "text", "placeholder": "Ex : Julie Bernard"},
        ],
    },
    {
        "titre": "📋 Livrables attendus",
        "questions": [
            {"key": "livrables", "label": "Éléments composant le rapport final (2 maximum)", "type": "textarea", "placeholder": "Ex :\n- Analyse documentaire du marché\n- Résultats de l'enquête et recommandations stratégiques"},
        ],
    },
]


def construire_brief_depuis_questionnaire(data: dict) -> str:
    """Compile les réponses du questionnaire en un brief structuré pour l'IA."""
    def ajouter(label, key, skip_values=None):
        val = data.get(key, "").strip()
        if not val or val in (skip_values or []):
            return ""
        return f"{label} : {val}"

    blocs = []

    # Entreprise
    bloc = [l for l in [
        ajouter("Entreprise cliente",       "nom_entreprise"),
        ajouter("Adresse",                  "adresse_ville"),
        ajouter("Taille",                   "taille_entreprise",  ["— Sélectionner —"]),
        ajouter("Contact client",           "nom_client"),
        ajouter("Coordonnées client",       "tel_mail_client"),
    ] if l]
    if bloc:
        blocs.append("=== ENTREPRISE CLIENTE ===\n" + "\n".join(bloc))

    # Projet
    bloc = [l for l in [
        ajouter("Titre de l'étude",         "titre_etude"),
        ajouter("Type d'étude",             "type_etude",         ["— Sélectionner —"]),
        ajouter("Contexte et enjeux",       "contexte_enjeux"),
        ajouter("Objectif principal",       "objectif_principal"),
        ajouter("Sous-objectifs",           "sous_objectifs"),
        ajouter("Résultats attendus",       "resultats_attendus"),
    ] if l]
    if bloc:
        blocs.append("=== PROJET ET OBJECTIFS ===\n" + "\n".join(bloc))

    # Produit / service
    bloc = [l for l in [
        ajouter("Produit / service étudié", "produit_service"),
        ajouter("Secteur d'activité",       "secteur"),
    ] if l]
    if bloc:
        blocs.append("=== PRODUIT / SERVICE ===\n" + "\n".join(bloc))

    # Cible & géo
    bloc = [l for l in [
        ajouter("Cible visée",              "cible"),
        ajouter("Zone géographique",        "zone_geo"),
        ajouter("Villes représentatives",   "villes_rep"),
        ajouter("Nombre de zones",          "nb_villes"),
    ] if l]
    if bloc:
        blocs.append("=== CIBLE ET GÉOGRAPHIE ===\n" + "\n".join(bloc))

    # Méthodologie
    bloc = [l for l in [
        ajouter("Nombre de questionnaires", "nb_questionnaires"),
        ajouter("Nombre d'intervenants",    "nb_intervenants"),
        ajouter("Mode de passation",        "mode_passation",     ["— Sélectionner —"]),
        ajouter("Durée du questionnaire",   "duree_questionnaire"),
        ajouter("Entretiens semi-directifs","esd_details"),
        ajouter("Visites mystères",         "visite_mystere_det"),
    ] if l]
    if bloc:
        blocs.append("=== MÉTHODOLOGIE ===\n" + "\n".join(bloc))

    # Marché / concurrence
    bloc = [l for l in [
        ajouter("Concurrents directs",      "concurrents_directs"),
        ajouter("Concurrents indirects",    "concurrents_indirects"),
        ajouter("Marché potentiel",         "marche_potentiel"),
    ] if l]
    if bloc:
        blocs.append("=== MARCHÉ ET CONCURRENCE ===\n" + "\n".join(bloc))

    # Communication
    bloc = [l for l in [
        ajouter("Canaux de communication",  "canaux_com"),
        ajouter("Actions de communication", "actions_com"),
        ajouter("Cible de communication",   "cible_com"),
    ] if l]
    if bloc:
        blocs.append("=== COMMUNICATION ===\n" + "\n".join(bloc))

    # Équipe
    bloc = [l for l in [
        ajouter("Chef de projet",           "chef_projet"),
        ajouter("Validateur",               "validateur"),
    ] if l]
    if bloc:
        blocs.append("=== ÉQUIPE NEOMA CONSEIL ===\n" + "\n".join(bloc))

    # Livrables
    bloc = [l for l in [
        ajouter("Livrables attendus",       "livrables"),
    ] if l]
    if bloc:
        blocs.append("=== LIVRABLES ===\n" + "\n".join(bloc))

    return "\n\n".join(blocs) if blocs else "(Aucune information fournie)"

# ── Config page ──────────────────────────────────────────────
st.set_page_config(
    page_title="Vador IA — MarketFlow",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ──────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp, .stMain { background-color: #ffffff; }

    .nc-header {
        padding: 12px 0 20px 0;
        border-bottom: 2px solid #1a3a6e;
        margin-bottom: 20px;
    }
    .nc-header h1 { margin: 0; font-size: 1.4rem; color: #1a3a6e; font-weight: 700; }
    .nc-header p  { margin: 2px 0 0 0; font-size: 0.82rem; color: #666; }

    /* Slide viewer */
    .viewer-main img {
        border-radius: 6px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.15);
    }
    .thumb-selected { outline: 3px solid #1a3a6e; border-radius: 4px; }
    .thumb-normal   { outline: 2px solid #e0e4ef; border-radius: 4px; }

    /* Nav buttons */
    div[data-testid="stHorizontalBlock"] button {
        border-radius: 6px;
    }

    /* Download */
    .stDownloadButton > button {
        background-color: #1a3a6e !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        width: 100%;
    }
    .stProgress > div > div { background-color: #1a3a6e; }
    footer, #MainMenu { visibility: hidden; }
</style>
""", unsafe_allow_html=True)


# ── Clé API ───────────────────────────────────────────────────
def get_api_key() -> str:
    try:
        return st.secrets["GROQ_API_KEY"]
    except Exception:
        pass
    return os.environ.get("GROQ_API_KEY", "")


# ── Conversion PPTX → images ─────────────────────────────────
def trouver_libreoffice() -> str | None:
    chemins = [
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/libreoffice",
        "/usr/bin/soffice",
        "libreoffice",
        "soffice",
    ]
    for c in chemins:
        if os.path.exists(c) or shutil.which(c):
            return c
    return None


def convertir_pptx_en_images(pptx_bytes: bytes, dpi: int = 150) -> list[bytes]:
    """Convertit chaque slide du PPTX en image PNG via LibreOffice + pymupdf."""
    import fitz

    soffice = trouver_libreoffice()
    if not soffice:
        return []

    with tempfile.TemporaryDirectory() as tmpdir:
        pptx_path = os.path.join(tmpdir, "presentation.pptx")
        with open(pptx_path, "wb") as f:
            f.write(pptx_bytes)

        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", tmpdir, pptx_path],
            capture_output=True,
            timeout=120,
        )

        pdf_path = os.path.join(tmpdir, "presentation.pdf")
        if not os.path.exists(pdf_path):
            return []

        zoom = dpi / 72
        mat  = fitz.Matrix(zoom, zoom)
        doc  = fitz.open(pdf_path)
        images = []
        for page in doc:
            pix = page.get_pixmap(matrix=mat, alpha=False)
            images.append(pix.tobytes("png"))
        doc.close()
        return images


# ── Visionneuse de slides ─────────────────────────────────────
def afficher_visionneuse(images: list[bytes], nom_fichier: str, pptx_bytes: bytes):
    n = len(images)
    if n == 0:
        st.info("Aperçu non disponible — téléchargez le fichier pour le visualiser.")
        return

    if "slide_idx" not in st.session_state or st.session_state.slide_idx >= n:
        st.session_state.slide_idx = 0

    idx = st.session_state.slide_idx

    col_viewer, col_right = st.columns([3, 1])

    with col_viewer:
        nav1, nav2, nav3, nav4, nav5 = st.columns([1, 1, 3, 1, 1])
        with nav1:
            if st.button("⏮", help="Première slide", use_container_width=True):
                st.session_state.slide_idx = 0
                st.rerun()
        with nav2:
            if st.button("◀", help="Précédente", disabled=(idx == 0), use_container_width=True):
                st.session_state.slide_idx -= 1
                st.rerun()
        with nav3:
            st.markdown(
                f"<p style='text-align:center;color:#666;padding-top:8px;font-size:0.9rem'>"
                f"Slide <strong>{idx + 1}</strong> / {n}</p>",
                unsafe_allow_html=True,
            )
        with nav4:
            if st.button("▶", help="Suivante", disabled=(idx == n - 1), use_container_width=True):
                st.session_state.slide_idx += 1
                st.rerun()
        with nav5:
            if st.button("⏭", help="Dernière slide", use_container_width=True):
                st.session_state.slide_idx = n - 1
                st.rerun()

        st.image(images[idx], use_container_width=True)

    with col_right:
        st.download_button(
            label="⬇️  Télécharger (.pptx)",
            data=pptx_bytes,
            file_name=nom_fichier,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )
        st.markdown("<br>", unsafe_allow_html=True)

        st.markdown(
            "<p style='font-size:0.8rem;color:#888;margin-bottom:8px'>Toutes les slides</p>",
            unsafe_allow_html=True,
        )
        for i, img in enumerate(images):
            is_sel = i == idx
            border = "3px solid #1a3a6e" if is_sel else "2px solid #e0e4ef"
            st.markdown(
                f'<div style="border:{border};border-radius:4px;overflow:hidden;margin-bottom:6px">',
                unsafe_allow_html=True,
            )
            st.image(img, use_container_width=True)
            st.markdown("</div>", unsafe_allow_html=True)
            if st.button(f"Slide {i + 1}", key=f"thumb_{i}", use_container_width=True):
                st.session_state.slide_idx = i
                st.rerun()


# ── Session state ─────────────────────────────────────────────
for key, default in [
    ("messages",         None),
    ("resume",           None),
    ("result",           None),
    ("images",           None),
    ("slide_idx",        0),
    ("all_replacements", None),
    ("template_path",    None),
    ("modif_messages",   None),
    ("tag_index",        None),
    ("history_stack",    None),
    ("slide_toggles",    None),
    ("mode_saisie",      "texte"),
    ("auto_generer",     False),
]:
    if key not in st.session_state:
        st.session_state[key] = default

if st.session_state.modif_messages is None:
    st.session_state.modif_messages = []
if st.session_state.history_stack is None:
    st.session_state.history_stack = []

# Init toggles depuis GROUPES_SLIDES (une seule fois)
if st.session_state.slide_toggles is None:
    from propale_core import GROUPES_SLIDES as _GS
    st.session_state.slide_toggles = {k: v["default"] for k, v in _GS.items()}

if st.session_state.messages is None:
    st.session_state.messages = [
        {
            "role": "assistant",
            "content": (
                "Bonjour ! Je suis **Vador IA**, votre assistant propale by MarketFlow.\n\n"
                "Décrivez votre mission ci-dessous pour générer votre proposition commerciale.\n\n"
                "**Incluez si possible :**\n"
                "- Nom et secteur du client\n"
                "- Type d'étude (marché, consommateurs, faisabilité...)\n"
                "- Zone géographique et cible\n"
                "- Objectifs et contexte\n\n"
                "*Plus le résumé est complet, plus la propale sera précise.*"
            ),
        }
    ]


# ── Sidebar : upload DETs + toggles sections ──────────────────
with st.sidebar:
    st.header("📂 Fichiers DET")
    st.caption("Uploadez jusqu'à 3 DETs Excel (un par option tarifaire)")

    det_files = []
    for i in range(1, 4):
        label = f"Option {i}" + (" (principale)" if i == 1 else "")
        det_files.append(st.file_uploader(label, type=["xlsx"], key=f"det{i}"))

    st.divider()
    uploaded_count = sum(1 for f in det_files if f is not None)
    if uploaded_count > 0:
        st.success(f"✅ {uploaded_count} DET(s) chargé(s)")
    else:
        st.info("ℹ️ Sans DET, l'IA générera les données financières.")

    st.divider()

    # ── Sections optionnelles ──
    from propale_core import GROUPES_SLIDES
    with st.expander("📑 Sections de la proposition", expanded=False):
        st.caption("Décochez les sections non pertinentes pour cette mission.")
        for key, group in GROUPES_SLIDES.items():
            default_val = st.session_state.slide_toggles.get(key, group["default"])
            st.toggle(
                group["label"],
                value=default_val,
                help=group["description"],
                key=f"tgl_{key}",
            )

    st.divider()
    with st.expander("⚙️ Template PowerPoint"):
        default_tpl = Path(__file__).parent / "template 2.pptx"
        if default_tpl.exists():
            size_kb = default_tpl.stat().st_size // 1024
            mtime   = datetime.fromtimestamp(default_tpl.stat().st_mtime).strftime("%d/%m/%Y %H:%M")
            st.caption(f"**{default_tpl.name}** · {size_kb} Ko · modifié le {mtime}")
        new_tpl = st.file_uploader(
            "Remplacer le template", type=["pptx"], key="template_upload",
            help="Le fichier est sauvegardé sur disque et reste actif après redémarrage."
        )
        if new_tpl is not None:
            tpl_dest = Path(__file__).parent / "template 2.pptx"
            tpl_dest.write_bytes(new_tpl.read())
            st.success("✅ Template sauvegardé — actif dès la prochaine génération.")

    st.divider()
    with st.expander("✏️ Personnaliser les prompts IA"):
        from propale_core import (
            BALISES_IA,
            charger_prompts_personnalises,
            sauvegarder_prompts_personnalises,
        )

        custom_prompts = charger_prompts_personnalises()
        nb_custom = len(custom_prompts)
        if nb_custom:
            st.caption(f"✏️ {nb_custom} prompt(s) personnalisé(s) actif(s)")
        else:
            st.caption("Tous les prompts utilisent les valeurs par défaut.")

        balise_choisie = st.selectbox(
            "Balise à modifier",
            options=list(BALISES_IA.keys()),
            key="prompt_selectbox",
        )

        is_custom   = balise_choisie in custom_prompts
        prompt_actif = custom_prompts.get(balise_choisie, BALISES_IA[balise_choisie])

        if is_custom:
            st.caption("✏️ Prompt personnalisé — diffère du défaut")
        else:
            st.caption("📄 Prompt par défaut")

        new_prompt = st.text_area(
            "Prompt envoyé à l'IA",
            value=prompt_actif,
            height=160,
            key=f"prompt_area_{balise_choisie}",
        )

        col_save, col_reset = st.columns(2)
        with col_save:
            if st.button("💾 Sauvegarder", use_container_width=True, key="btn_save_prompt"):
                custom_prompts[balise_choisie] = new_prompt
                sauvegarder_prompts_personnalises(custom_prompts)
                st.success("Sauvegardé !")
        with col_reset:
            if st.button(
                "↺ Défaut",
                use_container_width=True,
                key="btn_reset_prompt",
                disabled=not is_custom,
                help="Restaurer le prompt d'origine",
            ):
                del custom_prompts[balise_choisie]
                sauvegarder_prompts_personnalises(custom_prompts)
                st.rerun()

    st.caption("Vador IA by MarketFlow © 2025")


# ── Détection changement de toggles (après sidebar) ──────────
_current_toggles = {k: st.session_state.get(f"tgl_{k}", GROUPES_SLIDES[k]["default"])
                    for k in GROUPES_SLIDES}
_toggles_changed = (
    st.session_state.result is not None
    and _current_toggles != st.session_state.slide_toggles
)


# ── Helpers : historique & régénération ──────────────────────
def sauvegarder_historique():
    stack = st.session_state.history_stack
    stack.append({
        "replacements": dict(st.session_state.all_replacements),
        "toggles":      dict(st.session_state.slide_toggles),
    })
    if len(stack) > 15:
        st.session_state.history_stack = stack[-15:]


def calculer_slides_a_supprimer() -> list[int]:
    toggles = st.session_state.slide_toggles or {}
    out = []
    for key, group in GROUPES_SLIDES.items():
        if not toggles.get(key, group["default"]):
            out.extend(group["slides"])
    return out


def regenerer_tout():
    """Régénère propre + coloré + images selon replacements et toggles actuels."""
    from propale_core import regenerer_depuis_replacements, colorier_texte_remplace, supprimer_slides

    replacements  = st.session_state.all_replacements
    template_path = st.session_state.template_path

    pptx_clean                = regenerer_depuis_replacements(replacements, template_path)
    pptx_colored, tag_index   = colorier_texte_remplace(pptx_clean, replacements, template_path)

    a_supprimer = calculer_slides_a_supprimer()
    if a_supprimer:
        pptx_download = supprimer_slides(pptx_clean,   a_supprimer)
        pptx_preview  = supprimer_slides(pptx_colored, a_supprimer)
    else:
        pptx_download = pptx_clean
        pptx_preview  = pptx_colored

    images = convertir_pptx_en_images(pptx_preview, dpi=150)

    nom_client = st.session_state.result["nom_client"]
    timestamp  = datetime.now().strftime("%Y-%m-%d_%Hh%M")
    st.session_state.result["pptx_bytes"]  = pptx_download
    st.session_state.result["nom_fichier"] = f"Propale_{nom_client.replace(' ', '_')}_{timestamp}.pptx"
    st.session_state.images    = images
    st.session_state.slide_idx = 0
    st.session_state.tag_index = tag_index


def annuler_derniere_modification():
    if not st.session_state.history_stack:
        return False
    last = st.session_state.history_stack.pop()
    st.session_state.all_replacements = last["replacements"]
    st.session_state.slide_toggles    = last["toggles"]
    # Resync les widgets toggles pour éviter une re-détection de changement
    for key in last["toggles"]:
        st.session_state[f"tgl_{key}"] = last["toggles"][key]
    regenerer_tout()
    return True


# ── Génération initiale ───────────────────────────────────────
def lancer_generation(resume: str):
    from propale_core import generer_propale, colorier_texte_remplace, supprimer_slides

    api_key = get_api_key()
    if not api_key:
        st.error("❌ Clé API Groq manquante. Configurez `GROQ_API_KEY` dans les secrets Streamlit.")
        return

    template_path = Path(__file__).parent / "template 2.pptx"
    if not template_path.exists():
        st.error(f"❌ Template PPTX introuvable : `{template_path}`")
        return

    det_paths = []
    tmp_files = []
    for f in det_files:
        if f is not None:
            tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
            tmp.write(f.read())
            tmp.close()
            det_paths.append(tmp.name)
            tmp_files.append(tmp.name)
        else:
            det_paths.append(None)

    progress_bar = st.progress(0)
    status_text  = st.empty()

    def on_progress(step, total, msg):
        progress_bar.progress(step / total * 0.55)
        status_text.caption(f"⏳ {msg}")

    try:
        pptx_bytes, all_replacements, missed_tags, nom_client = generer_propale(
            resume=resume,
            det_paths=det_paths,
            api_key=api_key,
            template_path=str(template_path),
            progress_callback=on_progress,
        )

        status_text.caption("⏳ Génération de l'aperçu coloré...")
        progress_bar.progress(0.65)
        pptx_colored, tag_index = colorier_texte_remplace(pptx_bytes, all_replacements, str(template_path))

        # Appliquer les toggles de sections
        a_supprimer = []
        toggles = st.session_state.slide_toggles or {}
        for key, group in GROUPES_SLIDES.items():
            if not toggles.get(key, group["default"]):
                a_supprimer.extend(group["slides"])

        if a_supprimer:
            pptx_download = supprimer_slides(pptx_bytes,   a_supprimer)
            pptx_preview  = supprimer_slides(pptx_colored, a_supprimer)
        else:
            pptx_download = pptx_bytes
            pptx_preview  = pptx_colored

        status_text.caption("⏳ Conversion des slides en images...")
        progress_bar.progress(0.80)
        images = convertir_pptx_en_images(pptx_preview, dpi=150)
        progress_bar.progress(1.0)
        status_text.caption("✅ Proposition générée avec succès !")

        timestamp = datetime.now().strftime("%Y-%m-%d_%Hh%M")
        st.session_state.result = {
            "pptx_bytes":  pptx_download,
            "missed_tags": missed_tags,
            "nom_client":  nom_client,
            "nb_balises":  len(all_replacements),
            "timestamp":   timestamp,
            "nom_fichier": f"Propale_{nom_client.replace(' ', '_')}_{timestamp}.pptx",
        }
        st.session_state.images           = images
        st.session_state.slide_idx        = 0
        st.session_state.all_replacements = all_replacements
        st.session_state.template_path    = str(template_path)
        st.session_state.modif_messages   = []
        st.session_state.tag_index        = tag_index
        st.session_state.history_stack    = []  # Reset historique à chaque nouvelle génération

    except Exception as e:
        progress_bar.empty()
        status_text.empty()
        st.error(f"❌ Erreur : {e}")

    finally:
        for tmp in tmp_files:
            try:
                os.unlink(tmp)
            except Exception:
                pass


def appliquer_modification(request: str):
    from propale_core import interpreter_modification

    api_key      = get_api_key()
    replacements = st.session_state.all_replacements

    result = interpreter_modification(request, replacements, api_key, st.session_state.tag_index)

    if result["error"]:
        return False, result["error"]

    tag       = result["tag"]
    new_value = result["new_value"]
    old_value = replacements.get(tag, "")

    # Sauvegarder l'état actuel avant modification
    sauvegarder_historique()

    st.session_state.all_replacements[tag] = new_value
    regenerer_tout()

    tag_nom = tag.replace("{{", "").replace("}}", "")
    msg_ok  = f"✅ **`{tag_nom}`** mis à jour :\n> ~~{old_value[:80]}~~\n> **{new_value[:80]}**"
    return True, msg_ok


# ── Appliquer changement de toggles si détecté ───────────────
if _toggles_changed:
    sauvegarder_historique()
    st.session_state.slide_toggles = _current_toggles
    with st.spinner("Mise à jour des sections..."):
        regenerer_tout()
    st.rerun()


# ── Header ────────────────────────────────────────────────────
st.markdown("""
<div class="nc-header">
    <h1>⚡ Vador IA <span style="font-weight:400;font-size:1rem;color:#888">by MarketFlow</span></h1>
    <p>Remplissez le résumé de mission et uploadez vos DETs pour générer votre proposition commerciale.</p>
</div>
""", unsafe_allow_html=True)


# ── Zone de chat ──────────────────────────────────────────────
if st.session_state.result is not None:
    chat_col, _ = st.columns([1, 2])
else:
    chat_col = st.container()

with chat_col:
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"], avatar="⚡" if msg["role"] == "assistant" else "👤"):
            st.markdown(msg["content"])


# ── Interactions ─────────────────────────────────────────────
if st.session_state.resume is None:

    # ── Sélecteur de mode ─────────────────────────────────────
    with chat_col:
        st.markdown("**Comment souhaitez-vous saisir les informations ?**")
        m1, m2, _ = st.columns([1, 1, 2])
        with m1:
            if st.button(
                "✍️ Texte libre",
                use_container_width=True,
                type="primary" if st.session_state.mode_saisie == "texte" else "secondary",
                key="btn_mode_texte",
            ):
                st.session_state.mode_saisie = "texte"
                st.rerun()
        with m2:
            if st.button(
                "📋 Questionnaire",
                use_container_width=True,
                type="primary" if st.session_state.mode_saisie == "questionnaire" else "secondary",
                key="btn_mode_quest",
            ):
                st.session_state.mode_saisie = "questionnaire"
                st.rerun()

    # ── Mode texte libre ──────────────────────────────────────
    if st.session_state.mode_saisie == "texte":
        if prompt := st.chat_input("Décrivez votre mission ici..."):
            st.session_state.messages.append({"role": "user", "content": prompt})
            st.session_state.resume = prompt
            st.session_state.messages.append({
                "role": "assistant",
                "content": (
                    f"Parfait ! Résumé reçu ({len(prompt)} caractères). "
                    f"{uploaded_count} DET(s) chargé(s).\n\n"
                    "Cliquez sur **Générer la propale** pour lancer."
                ),
            })
            st.rerun()

    # ── Mode questionnaire ────────────────────────────────────
    else:
        st.markdown("---")
        st.caption("ℹ️ Seules les informations que vous connaissez sont nécessaires — l'IA complète le reste ou marque [À compléter].")

        with st.form("questionnaire_form", border=False):
            for section in QUESTIONNAIRE_SECTIONS:
                st.markdown(f"#### {section['titre']}")
                # Répartir en 2 colonnes les champs courts, pleine largeur pour les textareas
                short_qs  = [q for q in section["questions"] if q["type"] != "textarea"]
                long_qs   = [q for q in section["questions"] if q["type"] == "textarea"]

                if short_qs:
                    pairs = [short_qs[i:i+2] for i in range(0, len(short_qs), 2)]
                    for pair in pairs:
                        cols = st.columns(len(pair))
                        for col, q in zip(cols, pair):
                            with col:
                                if q["type"] == "text":
                                    st.text_input(q["label"], placeholder=q.get("placeholder", ""), key=f"q_{q['key']}")
                                elif q["type"] == "select":
                                    st.selectbox(q["label"], q["options"], key=f"q_{q['key']}")

                for q in long_qs:
                    st.text_area(q["label"], placeholder=q.get("placeholder", ""), height=90, key=f"q_{q['key']}")

                st.markdown("")  # espacement

            st.markdown("---")
            submitted = st.form_submit_button(
                "🚀 Générer la propale", type="primary", use_container_width=True
            )

        if submitted:
            data = {
                q["key"]: st.session_state.get(f"q_{q['key']}", "")
                for section in QUESTIONNAIRE_SECTIONS
                for q in section["questions"]
            }
            brief = construire_brief_depuis_questionnaire(data)
            sections_remplies = sum(
                1 for section in QUESTIONNAIRE_SECTIONS
                if any(
                    str(data.get(q["key"], "")).strip() and
                    str(data.get(q["key"], "")) not in ("— Sélectionner —", "")
                    for q in section["questions"]
                )
            )
            st.session_state.resume = brief
            st.session_state.messages.append({"role": "user", "content": f"📋 Questionnaire complété ({sections_remplies}/9 sections remplies)"})
            st.session_state.messages.append({
                "role": "assistant",
                "content": (
                    f"Questionnaire reçu — {sections_remplies} section(s) renseignée(s). "
                    f"{uploaded_count} DET(s) chargé(s). Génération en cours…"
                ),
            })
            st.session_state.auto_generer = True
            st.rerun()

elif st.session_state.result is None:
    # Auto-génération après questionnaire
    if st.session_state.auto_generer:
        st.session_state.auto_generer = False
        with st.spinner(""):
            lancer_generation(st.session_state.resume)
        if st.session_state.result is not None:
            st.rerun()

    with chat_col:
        st.markdown("---")
        c1, c2 = st.columns([3, 1])
        with c1:
            if st.button("🚀 Générer la propale", type="primary", use_container_width=True):
                with st.spinner(""):
                    lancer_generation(st.session_state.resume)
                if st.session_state.result is not None:
                    st.rerun()
        with c2:
            if st.button("🔄 Recommencer", use_container_width=True):
                for k in ["resume", "result", "images", "slide_idx", "all_replacements",
                          "template_path", "modif_messages", "history_stack", "auto_generer"]:
                    st.session_state[k] = None if k not in ("slide_idx", "auto_generer") else (0 if k == "slide_idx" else False)
                st.session_state.messages = st.session_state.messages[:1]
                st.rerun()

else:
    result = st.session_state.result
    with chat_col:
        st.success(f"✅ {result['nb_balises']} balises remplies")
        if st.button("🔄 Nouvelle propale", use_container_width=True):
            for k in ["resume", "result", "images", "slide_idx", "history_stack"]:
                st.session_state[k] = None if k != "slide_idx" else 0
            st.session_state.messages = st.session_state.messages[:1]
            st.rerun()

    # Visionneuse plein largeur
    st.markdown("---")
    afficher_visionneuse(
        st.session_state.images or [],
        result["nom_fichier"],
        result["pptx_bytes"],
    )

    # ── Chat de modification ──────────────────────────────────
    st.markdown("---")
    st.markdown("### ✏️ Modifier la proposition")

    # Ligne info + bouton undo côte à côte
    info_col, undo_col = st.columns([3, 1])
    with info_col:
        st.caption(
            "Les textes en **vert** portent un numéro **[N]** en exposant. "
            "Commencez par **[N] :** pour cibler la bonne balise. "
            "Ex : `[3] : Décathlon` · `[12] : 06 12 34 56 78`"
        )
    with undo_col:
        undo_disabled = len(st.session_state.history_stack) == 0
        if st.button(
            "↩️ Annuler",
            disabled=undo_disabled,
            use_container_width=True,
            help="Revenir à l'état précédent" if not undo_disabled else "Aucune modification à annuler",
        ):
            with st.spinner("Annulation..."):
                annuler_derniere_modification()
            st.session_state.modif_messages.append({
                "role": "assistant",
                "content": "↩️ Dernière modification annulée.",
            })
            st.rerun()

    # Historique des modifications
    for msg in (st.session_state.modif_messages or []):
        with st.chat_message(msg["role"], avatar="✏️" if msg["role"] == "assistant" else "👤"):
            st.markdown(msg["content"])

    if modif_prompt := st.chat_input(
        "Ex : [3] : Nike  •  [12] : 06 12 34 56 78  •  [7] : remplacer par 10 semaines",
        key="modif_input",
    ):
        st.session_state.modif_messages.append({"role": "user", "content": modif_prompt})

        with st.spinner("Modification en cours..."):
            ok, message = appliquer_modification(modif_prompt)

        st.session_state.modif_messages.append({"role": "assistant", "content": message})
        st.rerun()
