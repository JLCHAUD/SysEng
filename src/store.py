"""Gestion du store central JSON — source de vérité des variables GLOBAL."""
import json
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

STORE_PATH = Path(__file__).parent.parent / "output" / "store.json"


def _load() -> dict:
    if not STORE_PATH.exists():
        return {"version_store": "1", "derniere_maj": None, "variables": {}}
    with open(STORE_PATH, encoding="utf-8") as f:
        return json.load(f)


def _save(data: dict) -> None:
    STORE_PATH.parent.mkdir(parents=True, exist_ok=True)
    data["derniere_maj"] = datetime.now().isoformat(timespec="seconds")
    with open(STORE_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)


def get(nom_global: str) -> Optional[Any]:
    """Lit une variable du store. Retourne None si absente."""
    return _load()["variables"].get(nom_global)


def set(nom_global: str, valeur: Any) -> None:
    """Écrit une variable dans le store."""
    data = _load()
    data["variables"][nom_global] = valeur
    _save(data)


def set_many(variables: Dict[str, Any]) -> None:
    """Écrit plusieurs variables en une seule opération."""
    data = _load()
    data["variables"].update(variables)
    _save(data)


def get_all() -> Dict[str, Any]:
    """Retourne toutes les variables du store."""
    return _load()["variables"]


def delete(nom_global: str) -> None:
    data = _load()
    data["variables"].pop(nom_global, None)
    _save(data)


def snapshot() -> dict:
    """Retourne une copie complète du store (pour rapport)."""
    return _load()
