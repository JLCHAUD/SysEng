"""Gestion des espaces de travail (écosystèmes) du Schema Designer N1.

Un écosystème = un dossier config/ contenant :
  file_types.yaml, schema_relations.json, functions.json, templates.json, tables.json

L'app maintient une liste d'écosystèmes connus dans .schema_ecosystems.json
à la racine du projet, et un écosystème actif en mémoire.
"""
import json
from pathlib import Path
from typing import Optional

import yaml
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from web.schema_app.services.config_service import set_active_config, get_active_config

router = APIRouter()

# Racine du projet (SysEng/)
_PROJECT_ROOT = Path(__file__).parent.parent.parent.parent
_REGISTRY_FILE = _PROJECT_ROOT / ".schema_ecosystems.json"

# Fichiers requis pour un écosystème valide, avec leur contenu initial
_REQUIRED_FILES: dict[str, dict | str] = {
    "file_types.yaml":       None,   # YAML, géré séparément
    "schema_relations.json": {"version": "1", "relations": []},
    "functions.json":        {"version": "1", "functions": []},
    "templates.json":        {"version": "1", "templates": []},
    "tables.json":           {"version": "1", "tables": {}},
    "namespaces.json":       {"version": "1", "namespaces": []},
}


# ── Persistance de la liste d'écosystèmes ─────────────────────────────────────

def _load_known() -> list[dict]:
    if not _REGISTRY_FILE.exists():
        default = str(_PROJECT_ROOT / "config")
        return [{"name": "SysEng (défaut)", "path": default}]
    with open(_REGISTRY_FILE, encoding="utf-8") as f:
        return json.load(f).get("ecosystems", [])


def _save_known(ecosystems: list[dict]) -> None:
    with open(_REGISTRY_FILE, "w", encoding="utf-8") as f:
        json.dump({"ecosystems": ecosystems}, f, ensure_ascii=False, indent=2)


def _is_valid(path: Path) -> bool:
    return path.is_dir() and (path / "file_types.yaml").exists()


def _init_ecosystem(path: Path) -> None:
    """Crée les fichiers vides d'un nouvel écosystème."""
    path.mkdir(parents=True, exist_ok=True)
    # file_types.yaml
    ft_path = path / "file_types.yaml"
    if not ft_path.exists():
        with open(ft_path, "w", encoding="utf-8") as f:
            yaml.dump({"file_types": {}}, f, allow_unicode=True, default_flow_style=False)
    # fichiers JSON
    for fname, init_data in _REQUIRED_FILES.items():
        if fname == "file_types.yaml" or init_data is None:
            continue
        fp = path / fname
        if not fp.exists():
            with open(fp, "w", encoding="utf-8") as f:
                json.dump(init_data, f, ensure_ascii=False, indent=2)


# ── Modèles ────────────────────────────────────────────────────────────────────

class EcosystemInfo(BaseModel):
    name: str
    path: str
    active: bool = False
    valid: bool = True
    class_count: int = 0


class NewEcosystemRequest(BaseModel):
    name: str
    path: str   # chemin absolu ou relatif au projet


class OpenEcosystemRequest(BaseModel):
    path: str


# ── Routes ────────────────────────────────────────────────────────────────────

@router.get("/list", response_model=list[EcosystemInfo])
def list_ecosystems():
    known = _load_known()
    active_path = str(get_active_config())
    result = []
    for eco in known:
        p = Path(eco["path"])
        count = 0
        if _is_valid(p):
            try:
                import yaml as _yaml
                with open(p / "file_types.yaml", encoding="utf-8") as f:
                    ft = _yaml.safe_load(f) or {}
                count = len(ft.get("file_types", {}))
            except Exception:
                pass
        result.append(EcosystemInfo(
            name=eco["name"],
            path=eco["path"],
            active=(str(p.resolve()) == str(Path(active_path).resolve())),
            valid=_is_valid(p),
            class_count=count,
        ))
    return result


@router.post("/new", response_model=EcosystemInfo, status_code=201)
def new_ecosystem(body: NewEcosystemRequest):
    path = Path(body.path)
    if not path.is_absolute():
        path = _PROJECT_ROOT / body.path
    if path.exists() and any(path.iterdir()):
        raise HTTPException(409, f"Le dossier '{path}' existe déjà et n'est pas vide")

    _init_ecosystem(path)

    known = _load_known()
    if not any(k["path"] == str(path) for k in known):
        known.append({"name": body.name, "path": str(path)})
        _save_known(known)

    set_active_config(path)
    return EcosystemInfo(name=body.name, path=str(path), active=True, valid=True, class_count=0)


@router.post("/open", response_model=EcosystemInfo)
def open_ecosystem(body: OpenEcosystemRequest):
    path = Path(body.path)
    if not path.is_absolute():
        path = _PROJECT_ROOT / body.path
    if not path.exists():
        raise HTTPException(404, f"Dossier introuvable : {path}")
    if not _is_valid(path):
        raise HTTPException(422, f"Dossier invalide : file_types.yaml manquant dans '{path}'")

    known = _load_known()
    name = next((k["name"] for k in known if Path(k["path"]).resolve() == path.resolve()), path.name)
    if not any(Path(k["path"]).resolve() == path.resolve() for k in known):
        known.append({"name": name, "path": str(path)})
        _save_known(known)

    set_active_config(path)
    return EcosystemInfo(name=name, path=str(path), active=True, valid=True)


@router.get("/current", response_model=EcosystemInfo)
def current_ecosystem():
    path = get_active_config()
    known = _load_known()
    name = next((k["name"] for k in known if Path(k["path"]).resolve() == path.resolve()), path.name)
    count = 0
    if _is_valid(path):
        try:
            import yaml as _yaml
            with open(path / "file_types.yaml", encoding="utf-8") as f:
                ft = _yaml.safe_load(f) or {}
            count = len(ft.get("file_types", {}))
        except Exception:
            pass
    return EcosystemInfo(name=name, path=str(path), active=True, valid=_is_valid(path), class_count=count)


@router.post("/save")
def save_ecosystem():
    """Force-save : vérifie que tous les fichiers requis existent, les crée si absent."""
    path = get_active_config()
    for fname, init_data in _REQUIRED_FILES.items():
        fp = path / fname
        if not fp.exists():
            if fname == "file_types.yaml":
                import yaml as _yaml
                with open(fp, "w", encoding="utf-8") as f:
                    _yaml.dump({"file_types": {}}, f, allow_unicode=True)
            elif init_data is not None:
                with open(fp, "w", encoding="utf-8") as f:
                    json.dump(init_data, f, ensure_ascii=False, indent=2)
    return {"ok": True, "path": str(path), "message": "Écosystème sauvegardé"}


@router.delete("/{eco_name}", status_code=204)
def delete_ecosystem(eco_name: str):
    known = _load_known()
    eco = next((k for k in known if k["name"] == eco_name), None)
    if not eco:
        raise HTTPException(404, f"Écosystème '{eco_name}' introuvable")

    path = Path(eco["path"])
    active = get_active_config()
    if path.resolve() == active.resolve():
        raise HTTPException(409, "Impossible de supprimer l'écosystème actif")

    # Retire de la liste (ne supprime PAS les fichiers — action irréversible)
    known = [k for k in known if k["name"] != eco_name]
    _save_known(known)
