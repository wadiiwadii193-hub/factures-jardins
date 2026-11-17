"""
cloud_client.py — dialogue avec la « boîte aux lettres » (Google Sheet + Apps Script).

Réglages lus dans cet ordre : variables d'environnement, puis fichier config_cloud.txt
(à côté de ce fichier), lignes de la forme  NOM=valeur.
"""
from __future__ import annotations

import base64
import json
import re
import time
import os
import uuid
from pathlib import Path

import requests

CLIENT_VERSION = 4                                   # doit correspondre à l'attente de l'application
TIMEOUT = 45
CONFIG_FILE = "config_cloud.txt"
APPLIED_FILE = "ids_appliques.txt"


class CloudError(RuntimeError):
    """Problème de communication avec le Google Sheet."""


def setting(name: str, default: str = "") -> str:
    if os.environ.get(name):
        return os.environ[name].strip()
    path = Path(__file__).with_name(CONFIG_FILE)
    if path.exists():
        for line in path.read_text(encoding="utf-8-sig").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            if key.strip() == name:
                return value.strip().strip('"').strip("'")
    return default


def _propre(x):
    """Rend la requête sûre pour le script Google : NaN, valeurs vides de pandas, dates et nombres numpy → JSON valide."""
    if x is None or isinstance(x, (str, bool, int)):
        return x
    if isinstance(x, float):
        return None if x != x or x in (float("inf"), float("-inf")) else x
    if isinstance(x, dict):
        return {str(k): _propre(v) for k, v in x.items()}
    if isinstance(x, (list, tuple, set)):
        return [_propre(v) for v in x]
    if hasattr(x, "isoformat"):
        try:
            return x.isoformat()
        except Exception:  # noqa: BLE001      (NaT…)
            return None
    if hasattr(x, "item"):                            # nombre numpy
        try:
            return _propre(x.item())
        except Exception:  # noqa: BLE001
            return None
    return str(x)


def _extrait(texte: str, n: int = 170) -> str:
    """Début lisible d'une réponse non JSON (page d'erreur Google) pour comprendre ce qui ne va pas."""
    brut = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", texte or "", flags=re.S | re.I)
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", brut)).strip()[:n]


def _call(url: str, token: str, action: str, payload: dict | None = None, method: str = "POST") -> dict:
    body = _propre({"token": token, "action": action, **(payload or {})})
    erreur = ""
    for tentative in (1, 2):                          # un 2ᵉ essai : Google renvoie parfois une page vide ou une erreur passagère
        try:
            if method == "GET":
                r = requests.get(url, params=body, timeout=TIMEOUT)
            else:
                r = requests.post(url, data=json.dumps(body, allow_nan=False), timeout=TIMEOUT,
                                  headers={"Content-Type": "text/plain;charset=utf-8"})
        except requests.RequestException as e:
            erreur = f"Connexion impossible avec le Google Sheet : {e}"
        else:
            r.encoding = "utf-8"                      # les pages de Google sont en UTF-8 (accents lisibles dans les messages)
            if r.status_code >= 500 or r.status_code == 429:
                erreur = f"Google a répondu « erreur {r.status_code} » : {_extrait(r.text) or 'réponse vide'}"
            elif r.status_code >= 400:
                raise CloudError(f"Google a répondu « erreur {r.status_code} » : {_extrait(r.text) or 'réponse vide'}. "
                                 "L'adresse du script (SCRIPT_URL) est-elle correcte ?")
            else:
                try:
                    data = r.json()
                except ValueError:
                    erreur = (f"Réponse inattendue de Google (au lieu de données) : « {_extrait(r.text) or 'réponse vide'} ». "
                              "Dans Apps Script : Déployer › Gérer les déploiements › crayon › Nouvelle version › Déployer, "
                              "avec l'accès « Tout le monde ».")
                else:
                    if not isinstance(data, dict) or not data.get("ok"):
                        raise CloudError((data or {}).get("error", "Erreur inconnue côté Google Sheet.") if isinstance(data, dict)
                                         else "Réponse inattendue du Google Sheet.")
                    return data
        if tentative == 1:
            time.sleep(1.5)
    raise CloudError(erreur)


def ping(url: str, token: str) -> bool:
    return bool(_call(url, token, "ping", method="GET").get("ok"))


def add_entry(url: str, token: str, *, auteur: str, date_facture: str, categorie: str,
              montant: float, source: str = "", entry_id: str | None = None,
              photo: bytes | None = None, photo_type: str = "image/jpeg",
              deja_appliquee: bool = False, fournisseur: str = "", objet: str = "",
              numero_facture: str = "", mode_paiement: str = "", type_saisie: str = "") -> dict:
    """Dépose une facture dans le Google Sheet (et sa photo dans le Drive si `photo` est fournie)."""
    payload = {"id": entry_id or uuid.uuid4().hex, "auteur": auteur, "date_facture": date_facture,
               "categorie": categorie, "montant": round(float(montant), 2), "source": source}
    if photo:
        payload["photo_b64"] = base64.b64encode(photo).decode("ascii")
        payload["photo_type"] = photo_type
    if deja_appliquee:
        payload["statut"] = "appliqué"
    payload.update({"fournisseur": fournisseur, "objet": objet, "numero_facture": numero_facture,
                    "mode_paiement": mode_paiement})
    if type_saisie:
        payload["type"] = type_saisie                # « versement » pour un versement d'espèces
    return _call(url, token, "add", payload)


def list_pending(url: str, token: str) -> list[dict]:
    return _call(url, token, "list_pending", method="GET")["rows"]


def get_file(url: str, token: str, reference: str):
    """(octets, nom, type) d'un justificatif du Drive, d'après son adresse : pour l'enregistrer sur le PC."""
    r = _call(url, token, "get_file", {"url": reference})
    return base64.b64decode(r["data"]), r.get("nom", "justificatif"), r.get("type", "application/pdf")


def list_all(url: str, token: str) -> list[dict]:
    """Toutes les factures (à traiter et appliquées), avec le lien de la photo."""
    return _call(url, token, "list_all", method="GET")["rows"]


def list_validated(url: str, token: str) -> list[dict]:
    """Factures validées depuis le téléphone du trésorier, en attente d'être ajoutées à Excel."""
    return _call(url, token, "list_validated", method="GET")["rows"]


def validate(url: str, token: str, items) -> int:
    """items : liste de {id, categorie?, date_facture?, montant?} → statut « validée »."""
    return int(_call(url, token, "validate", {"items": list(items)}).get("updated", 0))


def reject(url: str, token: str, ids) -> int:
    return int(_call(url, token, "reject", {"ids": list(ids)}).get("updated", 0))


def mark_applied(url: str, token: str, ids) -> int:
    return int(_call(url, token, "mark_applied", {"ids": list(ids)}).get("updated", 0))


def set_categories(url: str, token: str, categories, kind: str = "sorties") -> int:
    """kind : « sorties » (catégories de dépenses) ou « entrees » (catégories des versements)."""
    payload = {"categories": list(categories)}
    if kind == "entrees":
        payload["kind"] = "entrees"
    return int(_call(url, token, "set_categories", payload).get("count", 0))


def get_categories(url: str, token: str, kind: str = "sorties") -> list[str]:
    """Liste des catégories (jamais None : liste vide si le script n'en renvoie pas)."""
    if kind == "entrees":
        data = _call(url, token, "get_categories", {"kind": "entrees"})
    else:
        data = _call(url, token, "get_categories", method="GET")
    return [str(x).strip() for x in (data.get("categories") or []) if str(x).strip()]


# ---- mémoire locale des lignes déjà reportées dans Excel (garde-fou anti-doublon) ----
def _applied_path() -> Path:
    return Path(__file__).with_name(APPLIED_FILE)


def load_applied_ids() -> set[str]:
    p = _applied_path()
    return set(p.read_text(encoding="utf-8").split()) if p.exists() else set()


def remember_applied(entry_id: str) -> None:
    with open(_applied_path(), "a", encoding="utf-8") as f:
        f.write(entry_id + "\n")
