"""
app_mobile.py — saisie d'une facture depuis un téléphone (envoi vers le Google Sheet).

Hébergée en ligne (Streamlit Community Cloud). Réglages dans les « Secrets » de l'application :
    SCRIPT_URL  = "https://script.google.com/macros/s/.../exec"
    TOKEN       = "le même code secret que dans le script Google"
    ACCESS_CODE    = "le code d'accès donné aux collègues"   (ils saisissent aussi leur prénom)
    TRESORIER_CODE = "le code du trésorier"   (donne en plus l'écran « Factures reçues » : valider / rejeter)
"""
from __future__ import annotations

import base64
import datetime as dt
import hashlib
import hmac
import importlib.util
import io
import pathlib
import sys
import tempfile
import time
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


TRESORIER_CODE = setting("TRESORIER_CODE")

# Écran affiché après « Fermer »
if st.session_state.get("ferme"):
    st.title("🌱 Jardins familiaux")
    st.success("Session fermée. Vous pouvez fermer cet onglet.")
    if st.button("Se reconnecter"):
        st.session_state.clear()
        st.rerun()
    st.stop()


def _signature(role: str, expire: int) -> str:
    return hmac.new(TOKEN.encode(), f"{role}|{expire}".encode(), "sha256").hexdigest()[:20]


def _jeton(role: str) -> str:
    expire = int(time.time()) + 6 * 3600                 # reste valable 6 h
    return f"{role}.{expire}.{_signature(role, expire)}"


def _role_du_jeton(jeton: str):
    """« c » (collègue) ou « t » (trésorier) si le jeton de l'adresse est valide, sinon None."""
    try:
        role, expire, signature = jeton.split(".")
        if role in ("c", "t") and int(expire) > time.time() and hmac.compare_digest(signature, _signature(role, int(expire))):
            return role
    except (ValueError, AttributeError):
        pass
    return None


def fermer_session():
    """Réinitialise tout : session, mémoire des lectures, adresse (jeton de reconnexion)."""
    st.session_state.clear()
    st.cache_data.clear()
    st.query_params.clear()
    st.session_state["ferme"] = True
    st.rerun()


# Si Android a recyclé l'onglet pendant la prise de photo, la page se recharge : on reconnecte sans redemander le code
if not st.session_state.get("autorise") and st.query_params.get("acces"):
    _role = _role_du_jeton(st.query_params.get("acces"))
    if _role:
        st.session_state["reprise"] = True            # la page a été rechargée : la photo en cours sera retrouvée
        st.session_state["autorise"] = True
        st.session_state["tresorier"] = _role == "t"


def identifier(code: str):
    """(accepté ?, trésorier ?) : le code des collègues donne l'envoi de factures, celui du trésorier donne aussi la validation."""
    code = code.strip()
    if not code:
        return False, False
    est_tresorier = bool(TRESORIER_CODE) and hmac.compare_digest(code.encode(), TRESORIER_CODE.encode())
    est_collegue = bool(ACCESS_CODE) and hmac.compare_digest(code.encode(), ACCESS_CODE.encode())
    return est_tresorier or est_collegue, est_tresorier


# ───────────────────────────── Code d'accès ─────────────────────────────
if (ACCESS_CODE or TRESORIER_CODE) and not st.session_state.get("autorise"):
    st.title("🌱 Jardins familiaux")
    with st.form("connexion"):
        code = st.text_input("Code d'accès", type="password")
        entrer = st.form_submit_button("Entrer", type="primary")
    if entrer:
        accepte, tresorier = identifier(code)
        if accepte:
            st.session_state["autorise"] = True
            st.session_state["tresorier"] = tresorier
            st.query_params["acces"] = _jeton("t" if tresorier else "c")
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


def lire_facture(raw: bytes, nom: str, recadrer: bool, cats: list):
    """Lecture automatique (montant, date, catégorie probable) + aperçu recadré. Peut être lente : jamais bloquante."""
    preview, texte, images = core.load_document(raw, nom)
    qualite = core.image_quality(images[0]) if images else None
    if images:
        images = [core.scan_document(i, decouper=recadrer) for i in images]
        preview = images[0]
    if texte is None:
        lecture = core.read_invoice(images, "Tesseract")
        texte, montant, candidats = lecture["text"], lecture["amount"], lecture["candidates"]
    else:
        montant, candidats = core.find_amount(texte)
    apercu = io.BytesIO()
    vignette = preview.convert("RGB")
    vignette.thumbnail((1600, 1600))
    vignette.save(apercu, "JPEG", quality=80, optimize=True)
    return {"amount": montant, "candidates": candidats, "date": core.find_date(texte), "text": texte,
            "suggestion": core.suggest_category(texte, cats), "preview_jpeg": apercu.getvalue(),
            "floue": bool(qualite and qualite["floue"])}


def photo_a_envoyer(raw: bytes, nom: str, lect: dict | None):
    """(octets, type) de la pièce justificative conservée dans le Drive : PDF d'origine, ou image recadrée et légère."""
    if nom.lower().endswith(".pdf") or raw[:5] == b"%PDF-":
        return raw, "application/pdf"
    if lect and lect.get("preview_jpeg"):
        return lect["preview_jpeg"], "image/jpeg"
    from PIL import Image
    img = Image.open(io.BytesIO(raw))
    img.thumbnail((1800, 1800))
    sortie = io.BytesIO()
    img.convert("RGB").save(sortie, "JPEG", quality=82, optimize=True)
    return sortie.getvalue(), "image/jpeg"


def factures_recues():
    """Écran du trésorier : voir les factures reçues, les corriger, les valider ou les rejeter."""
    import pandas as pd

    st.subheader("Factures reçues")
    try:
        cats = categories(SCRIPT_URL, TOKEN)
        recues = cloud.list_pending(SCRIPT_URL, TOKEN)
        validees = cloud.list_validated(SCRIPT_URL, TOKEN)
    except cloud.CloudError as e:
        st.error(str(e))
        return
    if validees:
        st.info(f"✅ {len(validees)} facture(s) validée(s) : elles s'ajouteront au tableau Excel "
                "dès que vous ouvrirez l'application sur votre PC.")
    st.button("🔄 Actualiser")
    if not recues:
        st.success("Aucune facture en attente.")
        return
    tout = st.checkbox("Tout sélectionner")
    lignes, vues = [], set()
    for r in recues:
        cle = (r["date_facture"], r["auteur"].lower(), r["categorie"].lower(), round(float(r["montant"]), 2))
        doublon = cle in vues
        vues.add(cle)
        lignes.append({"Valider": tout and not doublon, "Date": dt.date.fromisoformat(r["date_facture"]),
                       "Auteur": r["auteur"], "Catégorie": r["categorie"], "Montant (€)": float(r["montant"]),
                       "Photo": r.get("photo_url") or None,
                       "Remarque": "Doublon possible" if doublon else "", "id": r["id"]})
    options = list(cats) + sorted({x["Catégorie"] for x in lignes if x["Catégorie"] not in cats})
    edite = st.data_editor(
        pd.DataFrame(lignes), hide_index=True, key=f"recues_{tout}", disabled=["Auteur", "Photo", "Remarque"],
        column_config={"id": None,
                       "Date": st.column_config.DateColumn("Date", format="DD/MM/YYYY", required=True),
                       "Catégorie": st.column_config.SelectboxColumn("Catégorie", options=options, required=True),
                       "Montant (€)": st.column_config.NumberColumn(format="%.2f", min_value=0.01),
                       "Photo": st.column_config.LinkColumn("Photo", display_text="Ouvrir")})
    cochees = edite[edite["Valider"]]
    st.caption("Touchez une date, une catégorie ou un montant pour le corriger, cochez les lignes, puis validez.")
    gauche, droite = st.columns(2)
    if gauche.button("✅ Valider", type="primary", disabled=cochees.empty):
        items = [{"id": r["id"], "categorie": r["Catégorie"], "montant": float(r["Montant (€)"]),
                  "date_facture": pd.Timestamp(r["Date"]).date().isoformat()} for _, r in cochees.iterrows()]
        try:
            nb = cloud.validate(SCRIPT_URL, TOKEN, items)
        except cloud.CloudError as e:
            st.error(f"Validation impossible : {e}")
        else:
            st.session_state["message"] = (f"✅ {nb} facture(s) validée(s). Elles seront ajoutées au tableau "
                                           "à l'ouverture de l'application sur le PC.")
            st.rerun()
    if droite.button("🗑️ Rejeter", disabled=cochees.empty):
        try:
            nb = cloud.reject(SCRIPT_URL, TOKEN, list(cochees["id"]))
        except cloud.CloudError as e:
            st.error(f"Rejet impossible : {e}")
        else:
            st.session_state["message"] = f"🗑️ {nb} facture(s) rejetée(s)."
            st.rerun()


# ───────────────────────────── Page ─────────────────────────────
est_tresorier = bool(st.session_state.get("tresorier"))
_gauche, _droite = st.columns([3, 1])
_gauche.caption(f"Version 9 · {pathlib.Path(core.__file__).name}" + (" · trésorier" if est_tresorier else ""))
if _droite.button("🚪 Fermer", key="fermer"):
    fermer_session()
if "message" in st.session_state:
    st.success(st.session_state.pop("message"))
if est_tresorier:
    vue = st.radio("Que voulez-vous faire ?", ["🧾 Envoyer une facture", "📥 Factures reçues"], horizontal=True)
    if vue.startswith("📥"):
        factures_recues()
        st.stop()

st.title("🧾 Nouvelle facture")
st.caption("Dépense du jardin : la facture est envoyée au trésorier, qui la reporte dans le tableau.")
n = st.session_state.setdefault("n", 0)      # change après chaque envoi : remet les widgets à zéro

auteur = st.text_input("Votre prénom", key="auteur", placeholder="Ex. : Karim")

mode = st.radio("Comment ajouter la facture ?",
                ["📷 Prendre une photo", "🧾 Scanner un document", "📎 Joindre un fichier"])
recadrer = st.checkbox("Recadrer automatiquement la feuille (conseillé)", value=True)
if mode.startswith("📷"):
    st.caption("Touchez le bouton, puis choisissez « Prendre une photo » (iPhone) ou « Appareil photo » (Android). "
               "Facture bien à plat, bien éclairée, en entier.")
    types, libelle = ["jpg", "jpeg", "png"], "📷 Prendre une photo de la facture"
elif mode.startswith("🧾"):
    st.caption("La feuille est détectée, découpée et redressée.")
    types, libelle = ["jpg", "jpeg", "png"], "🧾 Photographier le document"
else:
    st.caption("Une photo de la galerie, ou un document PDF.")
    types, libelle = ["pdf", "jpg", "jpeg", "png"], "📎 Choisir une photo ou un document"

envoi = st.file_uploader(libelle, type=types, key=f"envoi_{mode[:1]}_{n}")
if envoi:
    signature = (envoi.name, envoi.size)
    if st.session_state.get(f"sig_{n}") != signature:             # une seule fois par fichier
        st.session_state[f"sig_{n}"] = signature
        brut = envoi.getvalue()
        try:
            brut = core.reduire_photo(brut, envoi.name)             # photo réduite : peu de mémoire, envoi rapide
        except Exception:  # noqa: BLE001
            pass
        st.session_state[f"recu_{n}"] = (brut, envoi.name)         # mémorisé dès la réception
data, name = st.session_state.get(f"recu_{n}", (None, ""))

if data is None:
    st.info("Ajoutez une facture : le formulaire s'affichera aussitôt.")
    st.stop()

fh = hashlib.md5(data).hexdigest()[:10]
entry_id = st.session_state.setdefault(f"id_{fh}", uuid.uuid4().hex)   # même id si on touche deux fois « Envoyer »
st.success(f"📎 Document reçu ({max(1, len(data) // 1024)} Ko). Complétez le formulaire ci-dessous.")

try:
    cats = categories(SCRIPT_URL, TOKEN)
except cloud.CloudError as e:
    st.error(str(e))
    st.stop()
if not cats:
    st.warning("La liste des catégories n'est pas encore disponible : le trésorier doit d'abord "
               "ouvrir l'application sur son PC.")
    st.stop()

# Résultat de la lecture automatique (s'il existe déjà) : appliqué AVANT la création des champs
cle_lect = f"lect_{fh}_{int(recadrer)}"
lect = st.session_state.get(cle_lect)
if lect and not st.session_state.get("init_" + cle_lect):
    st.session_state["init_" + cle_lect] = True
    if lect.get("amount") and not st.session_state.get(f"amt_{fh}"):
        st.session_state[f"amt_{fh}"] = float(lect["amount"])
    if st.session_state.get(f"cat_{fh}") is None:
        if lect.get("date"):
            st.session_state[f"date_{fh}"] = lect["date"]
        if lect.get("suggestion") is not None and lect["suggestion"] < len(cats):
            st.session_state[f"cat_{fh}"] = cats[lect["suggestion"]]
st.session_state.setdefault(f"amt_{fh}", 0.0)
st.session_state.setdefault(f"date_{fh}", dt.date.today())

if lect and lect.get("preview_jpeg"):
    st.image(lect["preview_jpeg"], caption="Document recadré")
elif name.lower().endswith(".pdf") or data[:5] == b"%PDF-":
    st.caption("📄 Document PDF reçu.")
else:
    st.image(data, caption="Document reçu")

st.markdown("**Vérifiez et corrigez si besoin**")
if lect is None:
    st.caption("⏳ Lecture automatique du montant et de la date en cours… Vous pouvez déjà remplir le formulaire.")
elif lect.get("erreur"):
    st.warning(f"Lecture automatique impossible ({lect['erreur']}). Saisissez le montant et la date à la main.")
else:
    if lect["floue"]:
        st.warning("La photo est floue ou trop petite : la lecture automatique risque de se tromper. "
                   "Rapprochez-vous de la facture, évitez les ombres, puis reprenez la photo.")
    if lect["amount"] is None:
        st.warning("Montant non détecté : saisissez-le.")
    if lect["date"] is None:
        st.warning("Date non détectée : choisissez-la.")
    if len(lect["candidates"]) > 1:
        st.caption("Autres montants lus : " + " · ".join(eur(c) for c in lect["candidates"] if c != lect["amount"]))

amount_in = st.number_input("Montant TTC (€)", min_value=0.0, step=0.01, format="%.2f", key=f"amt_{fh}")
date_in = st.date_input("Date de la facture", format="DD/MM/YYYY", key=f"date_{fh}")
cat = st.selectbox("Catégorie", cats, index=None, placeholder="Choisir une catégorie…", key=f"cat_{fh}")

if cat:
    st.info(f"➡️ **{cat}** · **{MOIS[date_in.month - 1]} {date_in.year}** · **{eur(amount_in)}**")

ready = bool(auteur.strip()) and bool(cat) and amount_in > 0
if not auteur.strip():
    st.caption("Indiquez votre prénom pour pouvoir envoyer.")

if st.button("📤 Envoyer au trésorier", type="primary", disabled=not ready):
    try:
        photo, photo_type = photo_a_envoyer(data, name, lect)
        res = cloud.add_entry(SCRIPT_URL, TOKEN, auteur=auteur.strip(), date_facture=date_in.isoformat(),
                              categorie=cat, montant=amount_in, source=f"mobile:{name}", entry_id=entry_id,
                              photo=photo, photo_type=photo_type)
    except cloud.CloudError as e:
        st.error(f"Envoi impossible : {e}")
    except Exception as e:  # noqa: BLE001
        st.error(f"Envoi impossible ({type(e).__name__}) : {e}")
    else:
        envoye = core.maintenant_fr()
        msg = (f"✅ Envoyé le {envoye:%d/%m/%Y} à {envoye:%H:%M:%S} : {eur(amount_in)} · {cat} · "
               f"{MOIS[date_in.month - 1]}. Merci !")
        if res.get("photo_erreur"):
            msg += " (La facture est bien enregistrée, mais sa photo n'a pas pu être conservée.)"
        st.session_state["message"] = msg
        st.session_state["n"] = n + 1
        st.rerun()

# Lecture automatique : en DERNIER, une fois le formulaire affiché. Une erreur ou un dépassement de temps
# n'empêche jamais d'envoyer la facture (on saisit alors le montant à la main).
if lect is None:
    try:
        with st.spinner("Lecture automatique de la facture…"):
            lect = lire_facture(data, name, recadrer, cats)
    except Exception as e:  # noqa: BLE001
        lect = {"erreur": f"{type(e).__name__}"}
    st.session_state[cle_lect] = lect
    st.rerun()
