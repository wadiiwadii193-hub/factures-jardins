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
import streamlit.components.v1 as components

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


# ── Prise de photo stable sur Android : l'appareil photo du téléphone s'ouvre directement, et la photo est
#    RÉDUITE dans le téléphone (1800 px, JPEG) avant l'envoi : peu de mémoire, envoi rapide, pas de blocage.
CAPTURE_HTML = r"""<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<style>
 body{margin:0;font-family:sans-serif}
 #btn{display:block;box-sizing:border-box;width:100%;padding:18px;font-size:18px;border-radius:10px;
      background:#ff4b4b;color:#fff;text-align:center;cursor:pointer}
 #info{margin-top:8px;font-size:14px;color:#444;text-align:center}
 #apercu{display:none;max-width:100%;max-height:150px;margin:8px auto 0;border-radius:6px}
</style></head><body>
<label id="btn" for="f">📷 Prendre la photo</label>
<input id="f" type="file" accept="image/*" capture="environment" style="display:none">
<div id="info"></div><img id="apercu" alt="">
<script>
(function () {
  var MAX = 1800, QUALITE = 0.85;
  var f = document.getElementById("f"), info = document.getElementById("info"), ap = document.getElementById("apercu");
  function envoyer(type, extra) { window.parent.postMessage(Object.assign({isStreamlitMessage: true, type: type}, extra || {}), "*"); }
  function hauteur() { envoyer("streamlit:setFrameHeight", {height: document.documentElement.scrollHeight + 6}); }
  envoyer("streamlit:componentReady", {apiVersion: 1});
  window.addEventListener("message", function (e) {
    if (e.data && e.data.type === "streamlit:render") {
      var a = e.data.args || {};
      if (a.libelle) { document.getElementById("btn").textContent = a.libelle; }
      if (a.mode === "fichier") {
        f.removeAttribute("capture"); f.setAttribute("accept", "image/*,application/pdf");   // galerie, appareil photo ou PDF
      } else {
        f.setAttribute("capture", "environment"); f.setAttribute("accept", "image/*");       // appareil photo arrière
      }
      if (a.effacer) { try { localStorage.removeItem("photo_en_cours"); } catch (e) {} }
      if (a.restaurer && !a.effacer && !restaure) { restaure = true; retrouver(); }
      hauteur();
    }
  });
  var restaure = false;
  function garder(v) {
    try { if (v.data.length < 3500000) { localStorage.setItem("photo_en_cours", JSON.stringify({t: Date.now(), v: v})); } } catch (e) {}
  }
  function retrouver() {
    try {
      var s = JSON.parse(localStorage.getItem("photo_en_cours") || "null");
      if (s && s.v && Date.now() - s.t < 900000) {
        info.textContent = "✅ Photo retrouvée après le rechargement de la page"; hauteur();
        envoyer("streamlit:setComponentValue", {value: s.v, dataType: "json"});
      }
    } catch (e) {}
  }
  function taille(w, h) { var k = Math.min(1, MAX / Math.max(w, h)); return [Math.round(w * k), Math.round(h * k)]; }
  function versJpeg(source, w, h) {
    var t = taille(w, h), c = document.createElement("canvas");
    c.width = t[0]; c.height = t[1];
    c.getContext("2d").drawImage(source, 0, 0, t[0], t[1]);
    return new Promise(function (ok, ko) {
      c.toBlob(function (b) { c.width = 0; c.height = 0; if (b) { ok([b, t[0], t[1]]); } else { ko(new Error("toBlob")); } },
               "image/jpeg", QUALITE);
    });
  }
  function viaImage(file) {
    return new Promise(function (ok, ko) {
      var u = URL.createObjectURL(file), im = new Image();
      im.onload = function () { URL.revokeObjectURL(u); versJpeg(im, im.naturalWidth, im.naturalHeight).then(ok, ko); };
      im.onerror = function () { URL.revokeObjectURL(u); ko(new Error("image")); };
      im.src = u;
    });
  }
  function viaBitmap(file) {
    return createImageBitmap(file, {imageOrientation: "from-image"}).then(function (bmp) {
      return versJpeg(bmp, bmp.width, bmp.height).then(function (r) { if (bmp.close) { bmp.close(); } return r; });
    });
  }
  function finir(blob, w, h) {
    var lecteur = new FileReader();
    lecteur.onload = function () {
      var data = String(lecteur.result).split(",")[1];
      ap.src = URL.createObjectURL(blob); ap.style.display = "block";
      info.textContent = "✅ Photo prête (" + Math.round(blob.size / 1024) + " Ko)";
      hauteur();
      var v = {data: data, w: w, h: h, nom: "photo.jpg", type: "image/jpeg", t: Date.now()};
      garder(v);
      envoyer("streamlit:setComponentValue", {value: v, dataType: "json"});
    };
    lecteur.onerror = function () { info.textContent = "❌ Lecture impossible : reprenez la photo."; hauteur(); };
    lecteur.readAsDataURL(blob);
  }
  f.addEventListener("change", function () {
    var file = f.files && f.files[0];
    if (!file) { return; }
    if (file.type === "application/pdf" || /\.pdf$/i.test(file.name || "")) {
      if (file.size > 6000000) { info.textContent = "❌ PDF trop lourd (6 Mo maximum)."; hauteur(); f.value = ""; return; }
      info.textContent = "⏳ Préparation du document…"; ap.style.display = "none"; hauteur();
      var lp = new FileReader();
      lp.onload = function () {
        info.textContent = "✅ Document prêt (" + Math.round(file.size / 1024) + " Ko)"; hauteur();
        var v = {data: String(lp.result).split(",")[1], nom: file.name, type: "application/pdf", t: Date.now()};
        garder(v);
        envoyer("streamlit:setComponentValue", {value: v, dataType: "json"});
        f.value = "";
      };
      lp.onerror = function () { info.textContent = "❌ Lecture impossible : réessayez."; hauteur(); };
      lp.readAsDataURL(file);
      return;
    }
    info.textContent = "⏳ Préparation de la photo…"; ap.style.display = "none"; hauteur();
    var p = window.createImageBitmap ? viaBitmap(file).catch(function () { return viaImage(file); }) : viaImage(file);
    p.then(function (r) { finir(r[0], r[1], r[2]); })
     .catch(function () { info.textContent = "❌ Photo illisible : reprenez-la."; hauteur(); })
     .then(function () { f.value = ""; });
  });
  hauteur();
})();
</script></body></html>"""

_dossier_capture = pathlib.Path(tempfile.gettempdir()) / "composant_capture_photo_v3"
_dossier_capture.mkdir(parents=True, exist_ok=True)
(_dossier_capture / "index.html").write_text(CAPTURE_HTML, encoding="utf-8")
capture_photo = components.declare_component("capture_photo", path=str(_dossier_capture))


def photo_du_telephone(cle: str, libelle: str, mode: str = "photo"):
    """(octets, nom) de la photo ou du document choisi dans le téléphone (déjà réduit), ou (None, "")."""
    valeur = capture_photo(libelle=libelle, mode=mode, restaurer=bool(st.session_state.get("reprise")),
                           effacer=bool(st.session_state.pop("effacer_photo", False)), key=cle, default=None)
    if valeur and valeur.get("data"):
        st.session_state["recu_" + cle] = (base64.b64decode(valeur["data"]), valeur.get("nom") or "photo.jpg")
    return st.session_state.get("recu_" + cle, (None, ""))


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
_gauche.caption(f"Version 7 · {pathlib.Path(core.__file__).name}" + (" · trésorier" if est_tresorier else ""))
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
n = st.session_state.setdefault("n", 0)

auteur = st.text_input("Votre prénom", key="auteur", placeholder="Ex. : Karim")

mode = st.radio("Comment ajouter la facture ?",
                ["📷 Prendre une photo", "🧾 Scanner un document", "📎 Joindre un fichier"])
recadrer = st.checkbox("Recadrer automatiquement la feuille (conseillé)", value=True)
data, name = None, ""
if mode.startswith("📷"):
    st.caption("Facture bien à plat, bien éclairée, en entier dans l'image.")
    data, name = photo_du_telephone(f"photo_{n}", "📷 Prendre la photo", "photo")
elif mode.startswith("🧾"):
    st.caption("La feuille est détectée, découpée et redressée.")
    data, name = photo_du_telephone(f"scan_{n}", "🧾 Scanner la feuille", "photo")
else:
    st.caption("Une photo de la galerie ou de l'appareil photo, ou un document PDF.")
    data, name = photo_du_telephone(f"fichier_{n}", "📎 Choisir une photo ou un document", "fichier")

with st.expander("Ça ne fonctionne pas ? Envoi classique"):
    classique = st.file_uploader("Photo ou PDF", type=["png", "jpg", "jpeg", "pdf"], key=f"classique_{n}")
    if classique:
        st.session_state[f"recu_classique_{n}"] = (classique.getvalue(), classique.name)
if data is None and st.session_state.get(f"recu_classique_{n}"):
    data, name = st.session_state[f"recu_classique_{n}"]

if data is None:
    st.info("Ajoutez une facture : le formulaire s'affichera aussitôt.")
    st.stop()

fh = hashlib.md5(data).hexdigest()[:10]
entry_id = st.session_state.setdefault(f"id_{fh}", uuid.uuid4().hex)
st.success(f"📎 Document reçu ({max(1, len(data) // 1024)} Ko). Complétez le formulaire ci-dessous.")

try:
    cats = categories(SCRIPT_URL, TOKEN)
except cloud.CloudError as e:
    st.error(str(e))
    st.stop()
if not cats:
    st.warning("La liste des catégories n'est pas encore disponible : le trésorier doit d'abord ouvrir l'application sur son PC.")
    st.stop()

# Gestion propre du justificatif reçu
if name.lower().endswith(".pdf") or data[:5] == b"%PDF-":
    st.caption("📄 Document PDF reçu.")
else:
    st.image(data, caption="Document reçu", use_column_width=True)

st.markdown("**Saisissez les détails de la facture**")

st.session_state.setdefault(f"amt_{fh}", 0.0)
st.session_state.setdefault(f"date_{fh}", dt.date.today())

amount_in = st.number_input("Montant TTC (€)", min_value=0.0, step=0.01, format="%.2f", key=f"amt_{fh}")
date_in = st.date_input("Date de la facture", format="DD/MM/YYYY", key=f"date_{fh}")
if f"cat_{fh}" in st.session_state:
    cat = st.selectbox("Catégorie", cats, placeholder="Choisir une catégorie…", key=f"cat_{fh}")
else:
    cat = st.selectbox("Catégorie", cats, index=None, placeholder="Choisir me catégorie…", key=f"cat_{fh}")

if cat:
    st.info(f"➡️ **{cat}** · **{MOIS[date_in.month - 1]} {date_in.year}** · **{eur(amount_in)}**")

ready = bool(auteur.strip()) and bool(cat) and amount_in > 0
if not auteur.strip():
    st.caption("Indiquez votre prénom pour pouvoir envoyer.")

if st.button("📤 Envoyer au trésorier", type="primary", disabled=not ready):
    try:
        photo, photo_type = photo_a_envoyer(data, name, None)
        res = cloud.add_entry(SCRIPT_URL, TOKEN, auteur=auteur.strip(), date_facture=date_in.isoformat(),
                              categorie=cat, montant=amount_in, source=f"mobile:{name}", entry_id=entry_id,
                              photo=photo, photo_type=photo_type)
    except cloud.CloudError as e:
        st.error(f"Envoi impossible : {e}")
    except Exception as e:
        st.error(f"Envoi impossible ({type(e).__name__}) : {e}")
    else:
        envoye = core.maintenant_fr()
        msg = (f"✅ Envoyé le {envoye:%d/%m/%Y} à {envoye:%H:%M:%S} : {eur(amount_in)} · {cat} · "
               f"{MOIS[date_in.month - 1]}. Merci !")
        if res.get("photo_erreur"):
            msg += " (La facture est bien enregistrée, mais sa photo n'a pas pu être conservée.)"
        st.session_state["message"] = msg
        st.session_state["n"] = n + 1
        st.session_state["effacer_photo"] = True
        st.rerun()

