"""Config service configurable pour le Schema Designer N1.

La différence avec web/services/config_service.py :
- Tous les chemins passent par get_config_dir(), modifiable via set_active_config().
- Cela permet de gérer plusieurs écosystèmes (dossiers config distincts).
"""
import json
from pathlib import Path
import yaml

# Écosystème actif — modifiable via set_active_config()
_DEFAULT_CONFIG = Path(__file__).parent.parent.parent.parent / "config"
_active_config: Path = _DEFAULT_CONFIG


def set_active_config(path: Path) -> None:
    global _active_config
    _active_config = Path(path)


def get_active_config() -> Path:
    return _active_config


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


# ── tables.json (tables std des Classes) ──────────────────────────────────────

def load_tables() -> dict:
    p = _p("tables.json")
    if not p.exists():
        return {"version": "1", "tables": {}}
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def save_tables(data: dict) -> None:
    with open(_p("tables.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
