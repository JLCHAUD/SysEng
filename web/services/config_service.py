"""Lecture/écriture des fichiers config — source de vérité unique."""
import json
from pathlib import Path
import yaml

CONFIG_DIR = Path(__file__).parent.parent.parent / "config"
OUTPUT_DIR = Path(__file__).parent.parent.parent / "output"

FILE_TYPES_PATH = CONFIG_DIR / "file_types.yaml"
REGISTRE_PATH   = CONFIG_DIR / "registre.json"
ACTEURS_PATH    = CONFIG_DIR / "acteurs.json"
ECOSYSTEM_PATH  = OUTPUT_DIR / "ecosystem.json"
HIERARCHY_PATH   = CONFIG_DIR / "hierarchy.json"
TABLES_PATH      = CONFIG_DIR / "tables.json"
RELATIONS_PATH   = CONFIG_DIR / "schema_relations.json"
FUNCTIONS_PATH   = CONFIG_DIR / "functions.json"
TEMPLATES_PATH   = CONFIG_DIR / "templates.json"


# ── file_types.yaml ────────────────────────────────────────────────────────────

def load_file_types() -> dict:
    with open(FILE_TYPES_PATH, encoding="utf-8") as f:
        data = yaml.safe_load(f)
    return data.get("file_types", {})


def save_file_types(types: dict) -> None:
    existing = {}
    if FILE_TYPES_PATH.exists():
        with open(FILE_TYPES_PATH, encoding="utf-8") as f:
            existing = yaml.safe_load(f) or {}
    existing["file_types"] = types
    with open(FILE_TYPES_PATH, "w", encoding="utf-8") as f:
        yaml.dump(existing, f, allow_unicode=True, default_flow_style=False, sort_keys=False)


# ── registre.json ─────────────────────────────────────────────────────────────

def load_registre() -> list:
    with open(REGISTRE_PATH, encoding="utf-8") as f:
        data = json.load(f)
    return data.get("fichiers", [])


def save_registre(fichiers: list) -> None:
    with open(REGISTRE_PATH, "w", encoding="utf-8") as f:
        json.dump({"version": "1", "fichiers": fichiers}, f, ensure_ascii=False, indent=2)


# ── acteurs.json ──────────────────────────────────────────────────────────────

def load_acteurs() -> list:
    with open(ACTEURS_PATH, encoding="utf-8") as f:
        return json.load(f)


def save_acteurs(acteurs: list) -> None:
    with open(ACTEURS_PATH, "w", encoding="utf-8") as f:
        json.dump(acteurs, f, ensure_ascii=False, indent=2)


# ── hierarchy.json ────────────────────────────────────────────────────────────

def load_hierarchy() -> dict:
    if not HIERARCHY_PATH.exists():
        return {"version": "1", "lists": [], "collects": []}
    with open(HIERARCHY_PATH, encoding="utf-8") as f:
        return json.load(f)


def save_hierarchy(data: dict) -> None:
    with open(HIERARCHY_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ── tables.json ───────────────────────────────────────────────────────────────

def load_tables() -> dict:
    if not TABLES_PATH.exists():
        return {"version": "1", "tables": {}}
    with open(TABLES_PATH, encoding="utf-8") as f:
        return json.load(f)


def save_tables(data: dict) -> None:
    with open(TABLES_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# ── schema_relations.json ─────────────────────────────────────────────────────

def load_relations() -> list:
    if not RELATIONS_PATH.exists():
        return []
    with open(RELATIONS_PATH, encoding="utf-8") as f:
        return json.load(f).get("relations", [])


def save_relations(relations: list) -> None:
    with open(RELATIONS_PATH, "w", encoding="utf-8") as f:
        json.dump({"version": "1", "relations": relations}, f, ensure_ascii=False, indent=2)


# ── functions.json ────────────────────────────────────────────────────────────

def load_functions() -> list:
    if not FUNCTIONS_PATH.exists():
        return []
    with open(FUNCTIONS_PATH, encoding="utf-8") as f:
        return json.load(f).get("functions", [])


def save_functions(functions: list) -> None:
    with open(FUNCTIONS_PATH, "w", encoding="utf-8") as f:
        json.dump({"version": "1", "functions": functions}, f, ensure_ascii=False, indent=2)


# ── templates.json ────────────────────────────────────────────────────────────

def load_templates() -> list:
    if not TEMPLATES_PATH.exists():
        return []
    with open(TEMPLATES_PATH, encoding="utf-8") as f:
        return json.load(f).get("templates", [])


def save_templates(templates: list) -> None:
    with open(TEMPLATES_PATH, "w", encoding="utf-8") as f:
        json.dump({"version": "1", "templates": templates}, f, ensure_ascii=False, indent=2)


# ── ecosystem.json (lecture seule) ────────────────────────────────────────────

def load_ecosystem() -> dict:
    if not ECOSYSTEM_PATH.exists():
        return {"version": "2.0", "files": {}, "edges": [], "tables": {}, "variables": {}}
    with open(ECOSYSTEM_PATH, encoding="utf-8") as f:
        return json.load(f)
