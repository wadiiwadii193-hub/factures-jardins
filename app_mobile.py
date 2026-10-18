"""
app_mobile.py — saisie d'une facture depuis un téléphone (envoi vers le Google Sheet).

Hébergée en ligne (Streamlit Community Cloud). Réglages dans les « Secrets » de l'application :
    SCRIPT_URL  = "https://script.google.com/macros/s/.../exec"
    TOKEN       = "le même code secret que dans le script Google"
    ACCESS_CODE = "le code d'accès à donner aux collègues"
"""
from __future__ import annotations

import datetime as dt
import hashlib
import hmac
import importlib.util
import pathlib
import sys
import unicodedata
import uuid

import streamlit as st

import cloud_client as cloud


def _charger_core():
    """Charge tresorerie_core.py, même si son nom a reçu un accent sur GitHub (trésorerie_core.py)."""
    try:
        import tresorerie_core
        return tresorerie_core
    except ModuleNotFoundError:
        pass
    ici = pathlib.Path(__file__).resolve().parent
    for fichier in ici.glob("*.py"):
        sans_accent = "".join(c for c in unicodedata.normalize("NFKD", fichier.stem)
                              if not unicodedata.combining(c))
        if sans_accent == "tresorerie_core":
            spec = importlib.util.spec_from_file_location("tresorerie_core", fichier)
            module = importlib.util.module_from_spec(spec)
            sys.modules["tresorerie_core"] = module
            spec.loader.exec_module(module)
            return module
    raise ModuleNotFoundError("Le fichier tresorerie_core.py est introuvable dans le dépôt GitHub.")


core = _charger_core()

st.set_page_config(page_title="Factures jardins familiaux", page_icon="🧾", layout="centered")
MOIS = [m.capitalize() for m in core.MOIS_FR]


def setting(name: str) -> str:
    try:
        value = st.secrets[name]
        if value:
            return str(value)
    except Exception:  # noqa: BLE001  (pas de fichier de secrets, ou clé absente)
        pass
    return cloud.setting(name)


def eur(x: float) -> str:
    return f"{x:,.2f}".replace(",", " ").replace(".", ",") + " €"


SCRIPT_URL, TOKEN, ACCESS_CODE = setting("SCRIPT_URL"), setting("TOKEN"), setting("ACCESS_CODE")
if not SCRIPT_URL or not TOKEN:
    st.error("Application non configurée : SCRIPT_URL et TOKEN manquent dans les réglages.")
    st.stop()

# ───────────────────────────── Code d'accès ─────────────────────────────
if ACCESS_CODE and not st.session_state.get("autorise"):
    st.title("🌱 Jardins familiaux")
    with st.form("connexion"):
        code = st.text_input("Code d'accès", type="password")
        entrer = st.form_submit_button("Entrer", type="primary")
    if entrer:
        if hmac.compare_digest(code.encode(), ACCESS_CODE.encode()):
            st.session_state["autorise"] = True
            st.rerun()
        else:
            st.error("Code incorrect.")
    st.stop()


@st.cache_data(ttl=300, show_spinner=False)
def categories(url: str, token: str) -> list[str]:
    return cloud.get_categories(url, token)


@st.cache_data(show_spinner="Lecture de la facture…")
def analyse(raw: bytes, filename: str, enhance: bool):
    preview, text, images = core.load_document(raw, filename)
    if enhance and images:
        images = [core.scan_document(i) for i in images]
        preview = images[0]
    if text is None:
        text = core.run_ocr(images, "Tesseract")
    return preview, text


# ───────────────────────────── Page ─────────────────────────────
st.title("🧾 Nouvelle facture")
st.caption("Dépense du jardin : la facture est envoyée au trésorier, qui la reporte dans le tableau.")

if "message" in st.session_state:
    st.success(st.session_state.pop("message"))
n = st.session_state.setdefault("n", 0)      # change après chaque envoi : remet les widgets à zéro

auteur = st.text_input("Votre prénom", key="auteur", placeholder="Ex. : Karim")

mode = st.radio("Comment ajouter la facture ?",
                ["📷 Prendre une photo", "🧾 Scanner un document", "📎 Joindre un fichier"])
data, name, enhance = None, "", False
if mode.startswith("📷"):
    st.caption("Si la caméra avant s'ouvre, choisissez plutôt « Joindre un fichier » puis « Prendre une photo ».")
    shot = st.camera_input("Cadrez la facture (bien à plat, bien éclairée)", key=f"cam_{n}")
    if shot:
        data, name = shot.getvalue(), "photo.jpg"
elif mode.startswith("🧾"):
    enhance = True
    st.caption("La feuille est détectée, redressée et passée en noir et blanc contrasté.")
    up = st.file_uploader("Photo ou PDF du document", type=["png", "jpg", "jpeg", "pdf"], key=f"scan_{n}")
    if up:
        data, name = up.getvalue(), up.name
else:
    up = st.file_uploader("Photo, image (PNG/JPG) ou PDF", type=["png", "jpg", "jpeg", "pdf"], key=f"file_{n}")
    if up:
        data, name = up.getvalue(), up.name

if data is None:
    st.info("Ajoutez une facture pour lancer la lecture automatique.")
    st.stop()

try:
    preview, text = analyse(data, name, enhance)
except core.OcrUnavailable as e:
    st.error(f"Lecture impossible : {e}")
    st.stop()
except Exception as e:  # noqa: BLE001
    st.error(f"Lecture impossible : {e}")
    st.stop()

try:
    cats = categories(SCRIPT_URL, TOKEN)
except cloud.CloudError as e:
    st.error(str(e))
    st.stop()
if not cats:
    st.warning("La liste des catégories n'est pas encore disponible : le trésorier doit d'abord "
               "ouvrir l'application sur son PC.")
    st.stop()

fh = hashlib.md5(data).hexdigest()[:10]
entry_id = st.session_state.setdefault(f"id_{fh}", uuid.uuid4().hex)   # même id si on touche deux fois « Envoyer »
amount, candidates = core.find_amount(text)
date = core.find_date(text)

st.image(preview, caption="Document analysé")
with st.expander("Texte lu par la machine"):
    st.text(text or "(aucun texte)")

st.markdown("**Vérifiez et corrigez si besoin**")
if amount is None:
    st.warning("Montant non détecté : saisissez-le.")
if date is None:
    st.warning("Date non détectée : choisissez-la.")
if len(candidates) > 1:
    st.caption("Autres montants lus : " + " · ".join(eur(c) for c in candidates if c != amount))

amount_in = st.number_input("Montant TTC (€)", min_value=0.0, value=float(amount or 0.0),
                            step=0.01, format="%.2f", key=f"amt_{fh}")
date_in = st.date_input("Date de la facture", value=date or dt.date.today(),
                        format="DD/MM/YYYY", key=f"date_{fh}")
cat = st.selectbox("Catégorie", cats, index=core.suggest_category(text, cats),
                   placeholder="Choisir une catégorie…", key=f"cat_{fh}")

if cat:
    st.info(f"➡️ **{cat}** · **{MOIS[date_in.month - 1]} {date_in.year}** · **{eur(amount_in)}**")

ready = bool(auteur.strip()) and bool(cat) and amount_in > 0
if not auteur.strip():
    st.caption("Indiquez votre prénom pour pouvoir envoyer.")

if st.button("📤 Envoyer au trésorier", type="primary", disabled=not ready):
    try:
        cloud.add_entry(SCRIPT_URL, TOKEN, auteur=auteur.strip(), date_facture=date_in.isoformat(),
                        categorie=cat, montant=amount_in, source=f"mobile:{name}", entry_id=entry_id)
    except cloud.CloudError as e:
        st.error(f"Envoi impossible : {e}")
    else:
        st.session_state["message"] = (f"✅ Envoyé : {eur(amount_in)} · {cat} · "
                                       f"{MOIS[date_in.month - 1]}. Merci !")
        st.session_state["n"] = n + 1
        st.rerun()
