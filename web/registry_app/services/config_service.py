"""Config service dynamique N2 — chemin configurable via set_active_config().

La différence avec web/services/config_service.py :
- Tous les chemins passent par get_active_config(), modifiable via set_active_config().
- Cela permet de gérer plusieurs écosystèmes (dossiers config distincts).
- L'écosystème actif est persisté dans .active_registry_ecosystem pour survivre aux reloads.
"""
import json
from pathlib import Path
import yaml

_PROJECT_ROOT   = Path(__file__).parent.parent.parent.parent
_DEFAULT_CONFIG = _PROJECT_ROOT / "config"
_ACTIVE_FILE    = _PROJECT_ROOT / ".active_registry_ecosystem"


def _load_persisted() -> Path:
    """Lit le chemin persisté sur disque, fallback sur _DEFAULT_CONFIG."""
    if _ACTIVE_FILE.exists():
        try:
            p = Path(_ACTIVE_FILE.read_text(encoding="utf-8").strip())
            if p.exists():
                return p
        except Exception:
            pass
    return _DEFAULT_CONFIG


# Écosystème actif — initialisé depuis le fichier persisté
_active_config: Path = _load_persisted()


def set_active_config(path: Path) -> None:
    global _active_config
    _active_config = Path(path)
    try:
        _ACTIVE_FILE.write_text(str(_active_config), encoding="utf-8")
    except Exception:
        pass


def get_active_config() -> Path:
    return _active_config


def get_file_types_path() -> Path:
    return _active_config / "file_types.yaml"


def _p(filename: str) -> Path:
    return _active_config / filename


# ── file_types.yaml ────────────────────────────────────────────────────────────

def load_file_types() -> dict:
    p = _p("file_types.yaml")
    if not p.exists():
        return {}
    with open(p, encoding="utf-8") as f:
        data = yaml.safe_load(f)
    return (data or {}).get("file_types", {})


def save_file_types(types: dict) -> None:
    p = _p("file_types.yaml")
    existing = {}
    if p.exists():
        with open(p, encoding="utf-8") as f:
            existing = yaml.safe_load(f) or {}
    existing["file_types"] = types
    with open(p, "w", encoding="utf-8") as f:
        yaml.dump(existing, f, allow_unicode=True, default_flow_style=False, sort_keys=False)


# ── registre.json ─────────────────────────────────────────────────────────────

def load_registre() -> list:
    p = _p("registre.json")
    if not p.exists():
        return []
    with open(p, encoding="utf-8") as f:
        data = json.load(f)
    return data.get("fichiers", [])


def save_registre(fichiers: list) -> None:
    with open(_p("registre.json"), "w", encoding="utf-8") as f:
        json.dump({"version": "1", "fichiers": fichiers}, f, ensure_ascii=False, indent=2)


# ── acteurs.json ──────────────────────────────────────────────────────────────

def load_acteurs() -> list:
    p = _p("acteurs.json")
    if not p.exists():
        return []
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def save_acteurs(acteurs: list) -> None:
    with open(_p("acteurs.json"), "w", encoding="utf-8") as f:
        json.dump(acteurs, f, ensure_ascii=False, indent=2)


# ── hierarchy.json ────────────────────────────────────────────────────────────

def load_hierarchy() -> dict:
    p = _p("hierarchy.json")
    if not p.exists():
        return {"version": "1", "lists": [], "collects": [], "pulls": []}
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def save_hierarchy(data: dict) -> None:
    with open(_p("hierarchy.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ── tables.json ───────────────────────────────────────────────────────────────

def load_tables() -> dict:
    p = _p("tables.json")
    if not p.exists():
        return {"version": "1", "tables": {}}
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def save_tables(data: dict) -> None:
    with open(_p("tables.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ── schema_relations.json ─────────────────────────────────────────────────────

def load_relations() -> list:
    p = _p("schema_relations.json")
    if not p.exists():
        return []
    with open(p, encoding="utf-8") as f:
        return json.load(f).get("relations", [])


def save_relations(relations: list) -> None:
    with open(_p("schema_relations.json"), "w", encoding="utf-8") as f:
        json.dump({"version": "1", "relations": relations}, f, ensure_ascii=False, indent=2)


# ── functions.json ────────────────────────────────────────────────────────────

def load_functions() -> list:
    p = _p("functions.json")
    if not p.exists():
        return []
    with open(p, encoding="utf-8") as f:
        return json.load(f).get("functions", [])


def save_functions(functions: list) -> None:
    with open(_p("functions.json"), "w", encoding="utf-8") as f:
        json.dump({"version": "1", "functions": functions}, f, ensure_ascii=False, indent=2)


# ── templates.json ────────────────────────────────────────────────────────────

def load_templates() -> list:
    p = _p("templates.json")
    if not p.exists():
        return []
    with open(p, encoding="utf-8") as f:
        return json.load(f).get("templates", [])


def save_templates(templates: list) -> None:
    with open(_p("templates.json"), "w", encoding="utf-8") as f:
        json.dump({"version": "1", "templates": templates}, f, ensure_ascii=False, indent=2)


# ── directories.json ─────────────────────────────────────────────────────────

def load_dirs() -> dict:
    p = _p("directories.json")
    if not p.exists():
        return {"posts_dir": ""}
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def save_dirs(data: dict) -> None:
    with open(_p("directories.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ── ecosystem.json (lecture seule) ────────────────────────────────────────────

def load_ecosystem() -> dict:
    # output/ecosystem.json est dans le dossier PARENT du config actif
    eco_path = _active_config.parent / "output" / "ecosystem.json"
    if not eco_path.exists():
        return {"version": "2.0", "files": {}, "edges": [], "tables": {}, "variables": {}}
    with open(eco_path, encoding="utf-8") as f:
        return json.load(f)
