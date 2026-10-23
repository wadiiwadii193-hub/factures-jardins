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
    """
    Charge la version À JOUR de tresorerie_core.py, même si GitHub en garde plusieurs exemplaires
    (nom avec ou sans accent : trésorerie_core.py).
    """
    ici = pathlib.Path(__file__).resolve().parent

    def sans_accent(texte):
        return "".join(c for c in unicodedata.normalize("NFKD", texte) if not unicodedata.combining(c))

    vus = []
    try:
        import tresorerie_core as exact
        if hasattr(exact, "prepare_justificatif"):
            return exact
        vus.append("tresorerie_core.py (ancienne version)")
    except ModuleNotFoundError:
        pass
    for i, fichier in enumerate(sorted(ici.glob("*.py"))):
        if sans_accent(fichier.stem) != "tresorerie_core":
            continue
        nom = f"tresorerie_core_{i}"
        spec = importlib.util.spec_from_file_location(nom, fichier)
        module = importlib.util.module_from_spec(spec)
        sys.modules[nom] = module
        spec.loader.exec_module(module)
        if hasattr(module, "prepare_justificatif"):
            sys.modules["tresorerie_core"] = module
            return module
        vus.append(f"{fichier.name} (ancienne version)")
    raise ModuleNotFoundError("Version à jour de tresorerie_core.py introuvable. Fichiers vus : "
                              + (", ".join(vus) or "aucun"))


st.set_page_config(page_title="Factures jardins familiaux", page_icon="🧾", layout="centered")
MOIS = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre",
        "Octobre", "Novembre", "Décembre"]


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


# Les modules lourds (lecture des factures) ne sont chargés qu'après le code d'accès :
# l'écran d'accueil s'affiche ainsi le plus vite possible.
with st.spinner("Chargement…"):
    try:
        core = _charger_core()
    except Exception as e:  # noqa: BLE001
        st.error(f"Démarrage impossible : {e}")
        st.stop()


@st.cache_data(ttl=300, show_spinner=False)
def categories(url: str, token: str) -> list[str]:
    return cloud.get_categories(url, token)


@st.cache_data(show_spinner="Lecture de la facture…")
def analyse(raw: bytes, filename: str, recadrer: bool):
    preview, text, images = core.load_document(raw, filename)
    qualite = core.image_quality(images[0]) if images else None
    if recadrer and images:
        images = [core.scan_document(i) for i in images]      # découpe la feuille, redresse, uniformise l'éclairage
        preview = images[0]
    if text is None:
        lecture = core.read_invoice(images, "Tesseract")
        text, amount, candidates = lecture["text"], lecture["amount"], lecture["candidates"]
    else:
        amount, candidates = core.find_amount(text)
    return {"preview": preview, "text": text, "amount": amount, "candidates": candidates,
            "floue": bool(qualite and qualite["floue"])}


# ───────────────────────────── Page ─────────────────────────────
st.title("🧾 Nouvelle facture")
st.caption("Dépense du jardin : la facture est envoyée au trésorier, qui la reporte dans le tableau.")

if "message" in st.session_state:
    st.success(st.session_state.pop("message"))
n = st.session_state.setdefault("n", 0)      # change après chaque envoi : remet les widgets à zéro

st.caption(f"Version 3 · {pathlib.Path(core.__file__).name}")
auteur = st.text_input("Votre prénom", key="auteur", placeholder="Ex. : Karim")

mode = st.radio("Comment ajouter la facture ?",
                ["📷 Prendre une photo", "🧾 Scanner un document", "📎 Joindre un fichier"])
recadrer = st.checkbox("Recadrer automatiquement la feuille (conseillé)", value=True)
data, name = None, ""
if mode.startswith("📷"):
    st.caption("Si la caméra avant s'ouvre, choisissez plutôt « Joindre un fichier » puis « Prendre une photo ».")
    shot = st.camera_input("Cadrez la facture (bien à plat, bien éclairée)", key=f"cam_{n}")
    if shot:
        data, name = shot.getvalue(), "photo.jpg"
elif mode.startswith("🧾"):
    st.caption("La feuille est détectée, découpée et redressée.")
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
    info = analyse(data, name, recadrer)
    preview, text = info["preview"], info["text"]
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
amount, candidates = info["amount"], info["candidates"]
date = core.find_date(text)

st.image(preview, caption="Document analysé")
with st.expander("Texte lu par la machine"):
    st.text(text or "(aucun texte)")

st.markdown("**Vérifiez et corrigez si besoin**")
if info["floue"]:
    st.warning("La photo est floue ou trop petite : la lecture automatique risque de se tromper. Rapprochez-vous de la facture, laissez le téléphone faire la mise au point, évitez les ombres, puis reprenez la photo.")
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
        photo, photo_type = core.prepare_justificatif(data, name, preview)
        res = cloud.add_entry(SCRIPT_URL, TOKEN, auteur=auteur.strip(), date_facture=date_in.isoformat(),
                              categorie=cat, montant=amount_in, source=f"mobile:{name}", entry_id=entry_id,
                              photo=photo, photo_type=photo_type)
    except cloud.CloudError as e:
        st.error(f"Envoi impossible : {e}")
    except Exception as e:  # noqa: BLE001
        st.error(f"Envoi impossible ({type(e).__name__}) : {e}")
    else:
        msg = f"✅ Envoyé : {eur(amount_in)} · {cat} · {MOIS[date_in.month - 1]}. Merci !"
        if res.get("photo_erreur"):
            msg += " (La facture est bien enregistrée, mais sa photo n'a pas pu être conservée.)"
        st.session_state["message"] = msg
        st.session_state["n"] = n + 1
        st.rerun()
