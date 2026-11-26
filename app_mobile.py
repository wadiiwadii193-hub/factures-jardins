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
import re
import sys
import tempfile
import time
import unicodedata
import uuid

import streamlit as st



def _sans_accent(texte):
    return "".join(c for c in unicodedata.normalize("NFKD", texte) if not unicodedata.combining(c))


def _nom_de_base(fichier):
    """« Trésorerie_core (1).py » → « tresorerie_core » : retrouve un fichier mal nommé au dépôt sur GitHub."""
    base = _sans_accent(fichier.stem).lower()
    base = re.sub(r"\s*\(\d+\)$", "", base)
    return re.sub(r"[\s_-]*(copie|copy)$", "", base)


def _charger(nom, attribut, exige=None):
    """
    Charge la version LA PLUS RÉCENTE (plus grand numéro `attribut`) parmi les fichiers « nom.py », « nom (1).py »,
    « trésorerie_core.py »… présents sur GitHub : une copie mal nommée ou en double ne bloque plus l'application.
    """
    ici = pathlib.Path(__file__).resolve().parent
    meilleur, meilleur_v = None, -1
    for i, fichier in enumerate(sorted(ici.glob("*.py"))):
        if _nom_de_base(fichier) != nom:
            continue
        nom_mod = f"{nom}__{i}"
        try:
            spec = importlib.util.spec_from_file_location(nom_mod, fichier)
            module = importlib.util.module_from_spec(spec)
            sys.modules[nom_mod] = module
            spec.loader.exec_module(module)
        except Exception:  # noqa: BLE001
            continue
        if exige and not hasattr(module, exige):
            continue
        v = getattr(module, attribut, 0)
        if meilleur is None or v > meilleur_v:
            meilleur, meilleur_v = module, v
    if meilleur is None:
        raise ModuleNotFoundError(f"{nom}.py est introuvable (ou illisible) dans le dépôt GitHub.")
    sys.modules[nom] = meilleur
    return meilleur


def _diagnostic():
    """Liste des fichiers du dépôt avec leur numéro de version : montre d'un coup d'œil ce qui n'est pas à jour."""
    ici = pathlib.Path(__file__).resolve().parent
    lignes = []
    for f in sorted(ici.glob("*.py")) + sorted(ici.glob("requirements*.txt")):
        try:
            t = f.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue
        v = re.search(r"^(?:CORE_VERSION|CLIENT_VERSION)\s*=\s*(\d+)", t, re.M) or re.search(r"Version (\d+) ·", t)
        lignes.append(f"{f.name:32s} " + (f"version {v.group(1)}" if v else ("" if f.suffix == ".txt" else "ancienne version (sans numéro)")))
    return "\n".join(lignes)


def _charger_core():
    return _charger("tresorerie_core", "CORE_VERSION", exige="prepare_justificatif")


try:
    cloud = _charger("cloud_client", "CLIENT_VERSION")
    _ERREUR_CLOUD = ""
except Exception as _e:  # noqa: BLE001
    cloud, _ERREUR_CLOUD = None, str(_e)


st.set_page_config(page_title="Factures jardins familiaux", page_icon="🧾", layout="centered")
if cloud is None:
    st.error(f"Démarrage impossible : {_ERREUR_CLOUD}")
    st.caption("Fichiers trouvés sur GitHub :")
    st.code(_diagnostic(), language=None)
    st.stop()
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


if getattr(core, "CORE_VERSION", 0) < 28 or getattr(cloud, "CLIENT_VERSION", 0) < 5:
    st.error("⚠️ Les fichiers de l'application ne sont pas tous à jour sur GitHub. Déposez ENSEMBLE app_mobile.py, "
             "tresorerie_core.py, cloud_client.py et requirements.txt (« Commit changes »), puis attendez 3 minutes. "
             f"(module {getattr(core, 'CORE_VERSION', 'ancien')} · client {getattr(cloud, 'CLIENT_VERSION', 'ancien')})")
    st.caption("Fichiers trouvés sur GitHub (il faut : tresorerie_core.py version 28 ou plus, cloud_client.py version 5 ou plus) :")
    st.code(_diagnostic(), language=None)
    st.stop()


@st.cache_data(ttl=300, show_spinner=False)
def categories(url: str, token: str) -> list[str]:
    return list(cloud.get_categories(url, token) or [])


@st.cache_data(ttl=300, show_spinner=False)
def categories_entrees(url: str, token: str) -> list[str]:
    """Catégories des versements : liste vide (et non un plantage) si elles ne sont pas disponibles."""
    try:
        return list(cloud.get_categories(url, token, kind="entrees") or [])
    except (cloud.CloudError, TypeError, KeyError, AttributeError):
        return []


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
        champs = core.extraire_champs(lecture)
        detail_total = lecture.get("detail_total", "")
    else:
        montant, candidats = core.find_amount(texte)
        champs = core.extraire_champs(None, texte)
        tri = core.triplet_tva(texte.splitlines())
        detail_total = core._texte_tri(tri) if tri else ""
    apercu = io.BytesIO()
    vignette = preview.convert("RGB")
    vignette.thumbnail((1600, 1600))
    vignette.save(apercu, "JPEG", quality=80, optimize=True)
    return {"amount": montant, "candidates": candidats, "detail_total": detail_total, "date": core.find_date(texte), "text": texte,
            "suggestion": core.suggest_category(texte, cats), "preview_jpeg": apercu.getvalue(),
            "fournisseur": champs["fournisseur"], "objet": champs["objet"], "numero": champs["numero"],
            "paiement": champs["paiement"], "floue": bool(qualite and qualite["floue"])}


def stocker_envois(cle_sig: str, cle_recu: str, fichiers) -> bool:
    """
    Mémorise le ou les fichiers reçus : une photo ou un PDF tel quel ; plusieurs photos réunies en un seul PDF
    (une page par photo, dans l'ordre). Une erreur est mémorisée pour être affichée, sans jamais faire planter la page.
    """
    if not fichiers:
        return False
    if not isinstance(fichiers, (list, tuple)):           # un seul fichier : même traitement
        fichiers = [fichiers]
    signature = tuple((f.name, f.size) for f in fichiers)
    if st.session_state.get(cle_sig) == signature:                  # une seule fois par lot de fichiers
        return False
    st.session_state[cle_sig] = signature
    try:
        lots = []
        for f in fichiers:
            brut = f.getvalue()
            try:
                brut = core.reduire_photo(brut, f.name)             # photo réduite : peu de mémoire, envoi rapide
            except Exception:  # noqa: BLE001
                pass
            lots.append((f.name, brut))
        st.session_state[cle_recu] = core.fusionner_pages(lots)
        st.session_state.pop(cle_recu + "_erreur", None)
    except Exception as e:  # noqa: BLE001
        st.session_state[cle_recu + "_erreur"] = str(e)
    return True


CAMERA_ANDROID_HTML = r"""<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;font-family:Arial,Helvetica,sans-serif">
<button id="b" style="width:100%;padding:14px 10px;font-size:17px;font-weight:600;border:0;border-radius:10px;background:#e8403c;color:#fff">📷 Ouvrir l'appareil photo</button>
<div id="m" style="font-size:14px;color:#b00020;margin-top:6px"></div>
<script>
(function () {
  var P = window.parent, D = null, msg = document.getElementById("m");
  try { D = P.document; D.body; } catch (e) { D = null; }
  function dire(t) { msg.textContent = t || ""; }
  function cible() {                       // l'envoi « un seul fichier » de Streamlit (pas celui de la galerie, qui accepte plusieurs fichiers)
    var choix = null;
    D.querySelectorAll('input[type="file"]').forEach(function (c) {
      if (!choix && !c.multiple && /jpe?g|png|image/i.test(c.accept || "")) { choix = c; }
    });
    return choix;
  }
  document.getElementById("b").addEventListener("click", function () {
    dire("");
    if (!D || !P.navigator || !P.navigator.mediaDevices) { dire("Appareil photo intégré indisponible : utilisez « Browse files » puis « Appareil photo »."); return; }
    if (!cible()) { dire("Zone d'envoi introuvable : utilisez « Browse files »."); return; }
    ouvrir();
  });
  function ouvrir() {
    var flux = null, fond = D.createElement("div");
    fond.id = "cam-fond";
    fond.style.cssText = "position:fixed;left:0;top:0;right:0;bottom:0;background:#000;z-index:2147483647;display:flex;flex-direction:column;font-family:Arial,sans-serif";
    var v = D.createElement("video");
    v.setAttribute("playsinline", ""); v.autoplay = true; v.muted = true;
    v.style.cssText = "flex:1;min-height:0;width:100%;object-fit:contain;background:#000";
    var barre = D.createElement("div");
    barre.style.cssText = "display:flex;align-items:center;justify-content:space-between;padding:14px 22px 26px;background:#000;color:#fff";
    var info = D.createElement("div"); info.style.cssText = "flex:1;font-size:14px;padding-right:10px"; info.textContent = "Facture à plat, bien éclairée, en entier dans l'image";
    var cap = D.createElement("button"); cap.id = "cam-capture"; cap.textContent = "";
    cap.style.cssText = "width:74px;height:74px;border-radius:50%;border:5px solid #fff;background:#e8403c;flex:none";
    var ann = D.createElement("button"); ann.id = "cam-annuler"; ann.textContent = "✕ Annuler";
    ann.style.cssText = "flex:1;text-align:right;background:none;border:0;color:#fff;font-size:16px";
    barre.appendChild(info); barre.appendChild(cap); barre.appendChild(ann);
    fond.appendChild(v); fond.appendChild(barre); D.body.appendChild(fond);
    function fermer() {
      if (flux) { flux.getTracks().forEach(function (t) { t.stop(); }); flux = null; }   // la caméra est libérée tout de suite (mémoire)
      if (fond.parentNode) { fond.parentNode.removeChild(fond); }
    }
    ann.addEventListener("click", fermer);
    P.navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal: "environment" }, width: { ideal: 1920 }, height: { ideal: 1080 } }, audio: false })
      .then(function (s) { flux = s; v.srcObject = s; })
      .catch(function (e) {
        info.textContent = "Appareil photo inaccessible (" + (e && e.name ? e.name : "erreur") + "). Autorisez-le dans les réglages du navigateur, ou utilisez « Browse files ».";
        cap.style.display = "none";
      });
    cap.addEventListener("click", function () {
      var w = v.videoWidth, h = v.videoHeight;
      if (!w || !h) { return; }
      var k = Math.min(1, 2000 / Math.max(w, h)), c = D.createElement("canvas");
      c.width = Math.round(w * k); c.height = Math.round(h * k);
      c.getContext("2d").drawImage(v, 0, 0, c.width, c.height);
      c.toBlob(function (blob) {
        fermer();
        if (!blob) { dire("La photo n'a pas pu être prise : réessayez."); return; }
        var champ = cible();
        if (!champ) { dire("Zone d'envoi introuvable : utilisez « Browse files »."); return; }
        var f = new P.File([blob], "photo_" + Date.now() + ".jpg", { type: "image/jpeg" });
        var dt = new P.DataTransfer(); dt.items.add(f);
        champ.files = dt.files;
        champ.dispatchEvent(new P.Event("change", { bubbles: true }));
      }, "image/jpeg", 0.85);
    });
  }
})();
</script></body></html>
"""


def est_android() -> bool:
    """Vrai sur un téléphone Android (d'après le navigateur). Jamais d'erreur : sinon on considère que ce n'est pas Android."""
    try:
        return "android" in str(st.context.headers.get("User-Agent", "")).lower()
    except Exception:  # noqa: BLE001
        return False


def envoi_photo(cle: str, libelle: str, camera: bool = True, ouvert: bool = False):
    """
    Retourne (octets, nom) du document à envoyer.
    • Appareil photo : UN fichier à la fois et seulement des images : c'est ce qui fait proposer « Appareil photo » sur
      Android (avec plusieurs fichiers ou un PDF, Android ne le propose plus). Chaque photo ajoute une page.
    • Fichiers déjà enregistrés (galerie, PDF) : plusieurs photos réunies en un PDF, ou un seul PDF.
    """
    pages = st.session_state.setdefault(f"pages_{cle}", [])
    if camera and est_android():
        # Android : l'appareil photo de l'application s'ouvre DANS la page (l'application « Appareil photo » du téléphone
        # fait recharger la page et perdre la photo). La photo prise est envoyée par la zone d'envoi ci-dessous.
        import streamlit.components.v1 as components
        components.html(CAMERA_ANDROID_HTML, height=78)
        st.caption("Si l'appareil photo ci-dessus ne s'ouvre pas, utilisez « Browse files » ci-dessous.")
    if camera:
        photo = st.file_uploader(libelle if not pages else "➕ Ajouter une page (facultatif)", type=["jpg", "jpeg", "png"],
                                 key=f"cam_{cle}_{len(pages)}")
        if photo is not None:
            brut = photo.getvalue()
            try:
                brut = core.reduire_photo(brut, photo.name)
            except Exception:  # noqa: BLE001
                pass
            pages.append((photo.name, brut))
            st.session_state[f"src_{cle}"] = "camera"
            st.rerun()                               # nouveau bouton vierge : on peut photographier la page suivante
        if pages:
            st.caption(f"📄 {len(pages)} page{'s' if len(pages) > 1 else ''} prise{'s' if len(pages) > 1 else ''}. "
                       "Touchez « Ajouter une page » pour la suivante, ou remplissez le formulaire plus bas.")
            if st.button("↩️ Retirer la dernière page", key=f"retire_{cle}_{len(pages)}"):
                pages.pop()
                st.rerun()
        with st.expander("📎 Ou choisir des photos / un PDF déjà enregistrés", expanded=ouvert):
            fichiers = st.file_uploader("Photos de la galerie (plusieurs possibles) ou un PDF", type=["pdf", "jpg", "jpeg", "png"],
                                        accept_multiple_files=True, key=f"gal_{cle}")
    else:
        fichiers = st.file_uploader(libelle, type=["pdf", "jpg", "jpeg", "png"], accept_multiple_files=True, key=f"gal_{cle}")
    if stocker_envois(f"sigg_{cle}", f"recug_{cle}", fichiers):
        st.session_state[f"src_{cle}"] = "fichiers"
    if not fichiers and st.session_state.get(f"src_{cle}") == "fichiers":      # les fichiers ont été retirés
        st.session_state.pop(f"recug_{cle}", None)
        st.session_state.pop(f"sigg_{cle}", None)
        st.session_state[f"src_{cle}"] = None
    if st.session_state.get(f"recug_{cle}_erreur"):
        st.error(st.session_state[f"recug_{cle}_erreur"])
    if st.session_state.get(f"src_{cle}") == "fichiers" and st.session_state.get(f"recug_{cle}"):
        return st.session_state[f"recug_{cle}"]
    if pages:
        signature = tuple(len(b) for _, b in pages)
        memo = st.session_state.get(f"fus_{cle}")
        if not memo or memo[0] != signature:
            try:
                memo = (signature, core.fusionner_pages(pages))
            except Exception as e:  # noqa: BLE001
                st.error(str(e))
                return None, ""
            st.session_state[f"fus_{cle}"] = memo
        return memo[1]
    return None, ""


def libelle_document(data: bytes, nom: str) -> str:
    """« 3 pages, 850 Ko » pour l'accusé de réception."""
    pages = core.nb_pages_pdf(data) if (nom.lower().endswith(".pdf") or data[:5] == b"%PDF-") else 1
    return (f"{pages} page{'s' if pages and pages > 1 else ''}, " if pages else "") + f"{max(1, len(data) // 1024)} Ko"


def photo_a_envoyer(raw: bytes, nom: str, lect: dict | None):
    """(octets, type) du justificatif conservé : PDF d'origine (allégé s'il est lourd), ou la photo (recadrée) en PDF."""
    if nom.lower().endswith(".pdf") or raw[:5] == b"%PDF-":
        return core.pdf_leger(raw), "application/pdf"
    source = lect["preview_jpeg"] if lect and lect.get("preview_jpeg") else raw
    try:
        return core.en_pdf(source, "photo.jpg"), "application/pdf"
    except Exception:  # noqa: BLE001      si la conversion échoue, on garde au moins la photo
        return core.reduire_photo(source, "photo.jpg"), "image/jpeg"


def factures_recues():
    """Écran du trésorier : voir les factures reçues, les corriger, les valider ou les rejeter."""
    import pandas as pd

    st.subheader("Factures et versements reçus")
    try:
        cats = categories(SCRIPT_URL, TOKEN)
        recues = cloud.list_pending(SCRIPT_URL, TOKEN)
        validees = cloud.list_validated(SCRIPT_URL, TOKEN)
    except cloud.CloudError as e:
        st.error(str(e))
        return
    _v = getattr(cloud, "DERNIERE_VERSION_SCRIPT", None)
    st.caption(f"Script Google : version {_v}" if _v else
               "⚠️ Script Google : ancienne version (ou mauvaise adresse). Redéployez le script (Déployer › Gérer les "
               "déploiements › Nouvelle version) et vérifiez SCRIPT_URL dans les Secrets Streamlit.")
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
        lignes.append({"Valider": tout and not doublon, "Type": "💶 Versement" if r.get("type") == "versement" else "🧾 Facture",
                       "Date": dt.date.fromisoformat(r["date_facture"]),
                       "Auteur": r["auteur"], "Fournisseur": r.get("fournisseur", ""), "Désignation": r.get("objet", ""),
                       "N° facture": r.get("numero_facture", ""), "Paiement": r.get("mode_paiement") or "Espèces",
                       "Catégorie": r["categorie"], "Montant (€)": float(r["montant"]),
                       "Photo": r.get("photo_url") or None,
                       "Remarque": "Doublon possible" if doublon else "", "id": r["id"]})
    cats_e = categories_entrees(SCRIPT_URL, TOKEN)
    options = list(cats) + [c for c in cats_e if c not in cats]
    options += sorted({x["Catégorie"] for x in lignes if x["Catégorie"] not in options})
    edite = st.data_editor(
        pd.DataFrame(lignes), hide_index=True, key=f"recues_{tout}", disabled=["Type", "Auteur", "Photo", "Remarque"],
        column_config={"id": None,
                       "Date": st.column_config.DateColumn("Date", format="DD/MM/YYYY", required=True),
                       "Catégorie": st.column_config.SelectboxColumn("Catégorie", options=options, required=True),
                       "Paiement": st.column_config.SelectboxColumn("Paiement", options=["CB", "Chèque", "Espèces", "Virement"]),
                       "Montant (€)": st.column_config.NumberColumn(format="%.2f", min_value=0.01),
                       "Photo": st.column_config.LinkColumn("Photo", display_text="Ouvrir")})
    cochees = edite[edite["Valider"]]
    st.caption("Touchez une date, une catégorie ou un montant pour le corriger, cochez les lignes, puis validez.")
    gauche, droite = st.columns(2)
    if gauche.button("✅ Valider", type="primary", disabled=cochees.empty):
        items = [{"id": r["id"], "categorie": r["Catégorie"], "montant": float(r["Montant (€)"]),
                  "date_facture": pd.Timestamp(r["Date"]).date().isoformat(),
                  "fournisseur": r["Fournisseur"], "objet": r["Désignation"], "numero_facture": r["N° facture"],
                  "mode_paiement": r["Paiement"]} for _, r in cochees.iterrows()]
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


def ecran_versement():
    """Versement d'espèces : photo du ticket de dépôt + saisie. Le trésorier valide, puis le PC met Excel à jour."""
    st.title("💶 Versement d'espèces")
    st.caption("Photographiez le ticket de dépôt de la banque. Le trésorier valide : le montant va dans les entrées et le "
               "ticket est rangé avec le versement.")
    auteur_v = st.text_input("Votre prénom", key="auteur", placeholder="Ex. : Karim")
    n = st.session_state.get("nv", 0)
    st.caption("Touchez le bouton, puis choisissez « Prendre une photo » (iPhone) ou « Appareil photo » (Android).")
    data, name = envoi_photo(f"v{n}", "📷 Photo du ticket de dépôt")
    sans_ticket = st.checkbox("Je n'ai pas de ticket", key=f"sans_{n}") if data is None else False
    if data is None and not sans_ticket:
        st.info("Ajoutez la photo du ticket : le formulaire s'affichera aussitôt.")
        return
    if data:
        st.success(f"📎 Ticket reçu ({libelle_document(data, name)}). Complétez le formulaire ci-dessous.")
    cats_e = categories_entrees(SCRIPT_URL, TOKEN)
    if not cats_e:
        st.warning("La liste des catégories de versement n'est pas encore disponible : le trésorier doit d'abord ouvrir "
                   "l'application sur son PC (et mettre à jour le script Google).")
        return
    fh = hashlib.md5(data).hexdigest()[:10] if data else "sans"
    cle = f"v_{n}_{fh}"
    cle_lect = f"lectv_{n}_{fh}"
    entry_id = st.session_state.setdefault(f"idv_{n}_{fh}", uuid.uuid4().hex)
    lect = st.session_state.get(cle_lect)
    if lect and not st.session_state.get("initv_" + cle_lect):          # valeurs proposées par la lecture du ticket
        st.session_state["initv_" + cle_lect] = True
        if lect.get("montant_v") and not st.session_state.get(f"{cle}_montant"):
            st.session_state[f"{cle}_montant"] = float(lect["montant_v"])
        if lect.get("date"):
            st.session_state[f"{cle}_date"] = lect["date"]
    st.session_state.setdefault(f"{cle}_montant", 0.0)
    st.session_state.setdefault(f"{cle}_date", dt.date.today())
    st.session_state.setdefault(f"{cle}_mode", "Espèces")
    if data:
        if lect and lect.get("preview_jpeg"):
            st.image(lect["preview_jpeg"], caption="Ticket de dépôt")
        elif name.lower().endswith(".pdf") or data[:5] == b"%PDF-":
            st.caption("📄 Document PDF reçu.")
        else:
            st.image(data, caption="Ticket reçu")
        if lect is None:
            st.caption("⏳ Lecture automatique du montant et de la date en cours… Vous pouvez déjà remplir le formulaire.")
        elif lect.get("erreur"):
            st.warning("Lecture automatique impossible : saisissez la date et le montant.")
    date_in = st.date_input("Date du versement", format="DD/MM/YYYY", key=f"{cle}_date")
    cat = st.selectbox("Catégorie (ex. Vente fumier)", cats_e, index=None, placeholder="Choisir une catégorie…", key=f"{cle}_cat")
    montant = st.number_input("Montant déposé (€)", min_value=0.0, step=0.01, format="%.2f", key=f"{cle}_montant")
    mode = st.selectbox("Mode", ["Espèces", "CB", "Virement", "Chèque"], key=f"{cle}_mode")
    precision = st.text_input("Précision (facultatif, ex. nom de l'acheteur)", key=f"{cle}_prec")
    if cat:
        st.info(f"➡️ **{cat}** · **{MOIS[date_in.month - 1]} {date_in.year}** · **{eur(montant)}**")
    ready = bool(auteur_v.strip()) and bool(cat) and montant > 0
    if not auteur_v.strip():
        st.caption("Indiquez votre prénom pour pouvoir envoyer.")
    if st.button("📤 Envoyer le versement au trésorier", type="primary", disabled=not ready, key=f"{cle}_envoi"):
        try:
            photo, photo_type = photo_a_envoyer(data, name, lect) if data else (None, "image/jpeg")
            res = cloud.add_entry(SCRIPT_URL, TOKEN, auteur=auteur_v.strip(), date_facture=date_in.isoformat(), categorie=cat,
                                  montant=montant, source="mobile:versement", entry_id=entry_id, photo=photo,
                                  photo_type=photo_type, fournisseur="", objet=precision.strip(), mode_paiement=mode,
                                  type_saisie="versement")
        except cloud.CloudError as e:
            st.error(f"Envoi impossible : {e}")
        except Exception as e:  # noqa: BLE001
            st.error(f"Envoi impossible ({type(e).__name__}) : {e}")
        else:
            envoye = core.maintenant_fr()
            msg = (f"✅ Versement envoyé le {envoye:%d/%m/%Y} à {envoye:%H:%M:%S} : {eur(montant)} · {cat} · "
                   f"{MOIS[date_in.month - 1]}. Merci !")
            if res.get("photo_erreur"):
                msg += " (Le versement est bien enregistré, mais la photo du ticket n'a pas pu être conservée.)"
            st.session_state["message"] = msg
            st.session_state["nv"] = n + 1
            st.rerun()
    if data and lect is None:                       # lecture du ticket en DERNIER : ne bloque jamais l'envoi
        try:
            with st.spinner("Lecture automatique du ticket…"):
                lect = lire_facture(data, name, True, [])
            lect["montant_v"] = core.find_amount_versement(lect.get("text", ""))[0]
        except Exception as e:  # noqa: BLE001
            lect = {"erreur": f"{type(e).__name__}"}
        st.session_state[cle_lect] = lect
        st.rerun()


# ───────────────────────────── Page ─────────────────────────────
est_tresorier = bool(st.session_state.get("tresorier"))
_gauche, _droite = st.columns([3, 1])
_gauche.caption(f"Version 21 · {pathlib.Path(core.__file__).name}" + (" · trésorier" if est_tresorier else ""))
if _droite.button("🚪 Fermer", key="fermer"):
    fermer_session()
if "message" in st.session_state:
    st.success(st.session_state.pop("message"))
_choix_vue = ["🧾 Envoyer une facture", "💶 Envoyer un versement d'espèces"] + (["📥 Factures reçues"] if est_tresorier else [])
vue = st.radio("Que voulez-vous faire ?", _choix_vue)
if vue.startswith("📥"):
    factures_recues()
    st.stop()
if vue.startswith("💶"):
    ecran_versement()
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

st.caption("Facture sur plusieurs pages ? Photographiez-les l'une après l'autre : elles sont réunies en un seul PDF, dans l'ordre.")
data, name = envoi_photo(f"f{n}", libelle, camera=(not mode.startswith("📎")) or est_android(), ouvert=mode.startswith("📎"))

if data is None:
    st.info("Ajoutez une facture : le formulaire s'affichera aussitôt.")
    st.stop()

fh = hashlib.md5(data).hexdigest()[:10]
entry_id = st.session_state.setdefault(f"id_{fh}", uuid.uuid4().hex)   # même id si on touche deux fois « Envoyer »
st.success(f"📎 Document reçu ({libelle_document(data, name)}). Complétez le formulaire ci-dessous.")

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
    if not st.session_state.get(f"four_{fh}") and lect.get("fournisseur"):
        st.session_state[f"four_{fh}"] = lect["fournisseur"]
    if not st.session_state.get(f"objet_{fh}") and lect.get("objet"):
        st.session_state[f"objet_{fh}"] = lect["objet"]
    if not st.session_state.get(f"num_{fh}") and lect.get("numero"):
        st.session_state[f"num_{fh}"] = lect["numero"]
    if lect.get("paiement"):
        st.session_state[f"pay_{fh}"] = lect["paiement"]
st.session_state.setdefault(f"amt_{fh}", 0.0)
st.session_state.setdefault(f"date_{fh}", dt.date.today())
st.session_state.setdefault(f"pay_{fh}", "Espèces")

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
    if lect.get("detail_total"):
        st.caption("🧮 Montant contrôlé : " + lect["detail_total"])
    if len(lect["candidates"]) > 1:
        st.caption("Autres montants lus : " + " · ".join(eur(c) for c in lect["candidates"] if c != lect["amount"]))

amount_in = st.number_input("Montant TTC (€)", min_value=0.0, step=0.01, format="%.2f", key=f"amt_{fh}")
date_in = st.date_input("Date de la facture", format="DD/MM/YYYY", key=f"date_{fh}")
cat = st.selectbox("Catégorie", cats, index=None, placeholder="Choisir une catégorie…", key=f"cat_{fh}")
fournisseur_in = st.text_input("Fournisseur (ex. Leroy Merlin)", key=f"four_{fh}")
objet_in = st.text_input("Désignation / Objet", key=f"objet_{fh}", placeholder="Ex. : arrosoir, gants")
numero_in = st.text_input("N° de facture", key=f"num_{fh}")
paiement_in = st.selectbox("Mode de paiement", ["CB", "Chèque", "Espèces"], key=f"pay_{fh}")

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
                              photo=photo, photo_type=photo_type, fournisseur=fournisseur_in.strip(),
                              objet=objet_in.strip(), numero_facture=numero_in.strip(), mode_paiement=paiement_in)
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
