"""
ExoSync — store central JSON
============================
Source de vérité pour toutes les variables GLOBAL échangées entre fichiers.

Deux modes d'utilisation :

1. Module (rétro-compat) — fonctions get/set/set_many/get_all/delete/snapshot
   qui opèrent sur le store par défaut (output/store.json) :

       from src import store as Store
       Store.get("uo.UO-001.avancement")

2. Classe JsonStore — pour les tests ou les contextes multi-stores :

       from src.store import JsonStore
       store = JsonStore(tmp_path / "store.json")
       store.set("uo.UO-001.avancement", 65.0)
"""
import json
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

DEFAULT_STORE_PATH = Path(__file__).parent.parent / "output" / "store.json"


# ─── Classe JsonStore ─────────────────────────────────────────────────────────

class JsonStore:
    """
    Store persistant JSON avec interface get/set/set_many/get_all/clear.

    Chaque opération de lecture charge le fichier ; chaque opération d'écriture
    recharge puis sauvegarde (sûr pour des usages séquentiels, pas concurrent).

    Args:
        path: chemin du fichier JSON (créé automatiquement si absent)
    """

    def __init__(self, path: Path = DEFAULT_STORE_PATH):
        self.path = Path(path)

    # ── Lecture/écriture bas niveau ───────────────────────────────────────────

    def _load(self) -> dict:
        if not self.path.exists():
            return {"version_store": "1", "derniere_maj": None, "variables": {}}
        with open(self.path, encoding="utf-8") as f:
            return json.load(f)

    def _save(self, data: dict) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        data["derniere_maj"] = datetime.now().isoformat(timespec="seconds")
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2, default=str)

    # ── API publique ──────────────────────────────────────────────────────────

    def get(self, nom_global: str) -> Optional[Any]:
        """Lit une variable du store. Retourne None si absente."""
        return self._load()["variables"].get(nom_global)

    def set(self, nom_global: str, valeur: Any) -> None:
        """Écrit une variable dans le store."""
        data = self._load()
        data["variables"][nom_global] = valeur
        self._save(data)

    def set_many(self, variables: Dict[str, Any]) -> None:
        """Écrit plusieurs variables en une seule opération atomique."""
        data = self._load()
        data["variables"].update(variables)
        self._save(data)

    def get_all(self) -> Dict[str, Any]:
        """Retourne toutes les variables du store."""
        return self._load()["variables"]

    def delete(self, nom_global: str) -> None:
        """Supprime une variable du store (silencieux si absente)."""
        data = self._load()
        data["variables"].pop(nom_global, None)
        self._save(data)

    def clear(self) -> None:
        """Vide toutes les variables du store."""
        data = self._load()
        data["variables"] = {}
        self._save(data)

    def snapshot(self) -> dict:
        """Retourne une copie complète du store (pour rapport ou debug)."""
        return self._load()

    def keys(self, prefix: str = "") -> list:
        """Retourne les clés du store, filtrées par préfixe optionnel."""
        all_keys = list(self._load()["variables"].keys())
        if prefix:
            return [k for k in all_keys if k.startswith(prefix)]
        return all_keys

    def __repr__(self) -> str:
        n = len(self._load()["variables"])
        return f"JsonStore({self.path}, {n} variable(s))"


# ─── Store par défaut (module-level, rétro-compat) ───────────────────────────

_default = JsonStore(DEFAULT_STORE_PATH)


def get(nom_global: str) -> Optional[Any]:
    """Lit une variable du store par défaut."""
    return _default.get(nom_global)


def set(nom_global: str, valeur: Any) -> None:
    """Écrit une variable dans le store par défaut."""
    _default.set(nom_global, valeur)


def set_many(variables: Dict[str, Any]) -> None:
    """Écrit plusieurs variables dans le store par défaut."""
    _default.set_many(variables)


def get_all() -> Dict[str, Any]:
    """Retourne toutes les variables du store par défaut."""
    return _default.get_all()


def delete(nom_global: str) -> None:
    """Supprime une variable du store par défaut."""
    _default.delete(nom_global)


def snapshot() -> dict:
    """Retourne un snapshot du store par défaut."""
    return _default.snapshot()
