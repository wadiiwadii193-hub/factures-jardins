"""
cloud_client.py — dialogue avec la « boîte aux lettres » (Google Sheet + Apps Script).

Réglages lus dans cet ordre : variables d'environnement, puis fichier config_cloud.txt
(à côté de ce fichier), lignes de la forme  NOM=valeur.
"""
from __future__ import annotations

import json
import os
import uuid
from pathlib import Path

import requests

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


def _call(url: str, token: str, action: str, payload: dict | None = None, method: str = "POST") -> dict:
    body = {"token": token, "action": action, **(payload or {})}
    try:
        if method == "GET":
            r = requests.get(url, params=body, timeout=TIMEOUT)
        else:
            r = requests.post(url, data=json.dumps(body), timeout=TIMEOUT,
                              headers={"Content-Type": "text/plain;charset=utf-8"})
        r.raise_for_status()
        data = r.json()
    except requests.RequestException as e:
        raise CloudError(f"Connexion impossible avec le Google Sheet : {e}") from e
    except ValueError as e:
        raise CloudError("Réponse inattendue : l'adresse du script (SCRIPT_URL) est-elle correcte "
                         "et le déploiement ouvert à « Tout le monde » ?") from e
    if not data.get("ok"):
        raise CloudError(data.get("error", "Erreur inconnue côté Google Sheet."))
    return data


def ping(url: str, token: str) -> bool:
    return bool(_call(url, token, "ping", method="GET").get("ok"))


def add_entry(url: str, token: str, *, auteur: str, date_facture: str, categorie: str,
              montant: float, source: str = "", entry_id: str | None = None) -> dict:
    return _call(url, token, "add", {
        "id": entry_id or uuid.uuid4().hex, "auteur": auteur, "date_facture": date_facture,
        "categorie": categorie, "montant": round(float(montant), 2), "source": source})


def list_pending(url: str, token: str) -> list[dict]:
    return _call(url, token, "list_pending", method="GET")["rows"]


def mark_applied(url: str, token: str, ids) -> int:
    return int(_call(url, token, "mark_applied", {"ids": list(ids)}).get("updated", 0))


def set_categories(url: str, token: str, categories) -> int:
    return int(_call(url, token, "set_categories", {"categories": list(categories)}).get("count", 0))


def get_categories(url: str, token: str) -> list[str]:
    return _call(url, token, "get_categories", method="GET")["categories"]


# ---- mémoire locale des lignes déjà reportées dans Excel (garde-fou anti-doublon) ----
def _applied_path() -> Path:
    return Path(__file__).with_name(APPLIED_FILE)


def load_applied_ids() -> set[str]:
    p = _applied_path()
    return set(p.read_text(encoding="utf-8").split()) if p.exists() else set()


def remember_applied(entry_id: str) -> None:
    with open(_applied_path(), "a", encoding="utf-8") as f:
        f.write(entry_id + "\n")
