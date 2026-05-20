"""Gestion des Gabarits — points de départ réutilisables pour créer des Affaires.

Un **Gabarit** est un dossier contenant uniquement le schéma N1 :
  file_types.yaml, schema_relations.json, functions.json, templates.json
  (optionnellement : namespaces.json)

Une **Affaire** est un écosystème complet N1+N2 :
  tout le N1 + registre.json, acteurs.json, hierarchy.json, tables.json

Cloner un Gabarit crée une Affaire indépendante (copie, pas une référence).
"""
import json
import shutil
from pathlib import Path
from typing import Optional

import yaml
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from web.registry_app.services.config_service import set_active_config, get_active_config
from web.workspace_service import get_gabarits_dir, get_workspace_dir

router = APIRouter()

_PROJECT_ROOT = Path(__file__).parent.parent.parent.parent
_INVALID_NAME_CHARS = set(r'\/:*?"<>|')
_GABARITS_FILE = _PROJECT_ROOT / ".gabarits.json"

# Fichiers N1 copiés lors d'un clonage
_N1_FILES = [
    "file_types.yaml",
    "schema_relations.json",
    "functions.json",
    "templates.json",
    "namespaces.json",   # optionnel — ignoré s'il n'existe pas
]

# Fichiers N2 créés vides dans la nouvelle Affaire
_N2_INIT: dict[str, dict | list] = {
    "registre.json":  {"version": "1", "fichiers": []},
    "acteurs.json":   [],
    "hierarchy.json": {"version": "1", "lists": [], "collects": [], "pulls": []},
    "tables.json":    {"version": "1", "tables": {}},
}


# ── Persistance ───────────────────────────────────────────────────────────────

def _load_known() -> list[dict]:
    if not _GABARITS_FILE.exists():
        return []
    with open(_GABARITS_FILE, encoding="utf-8") as f:
        return json.load(f).get("gabarits", [])


def _save_known(gabarits: list[dict]) -> None:
    with open(_GABARITS_FILE, "w", encoding="utf-8") as f:
        json.dump({"gabarits": gabarits}, f, ensure_ascii=False, indent=2)


def _is_valid_gabarit(path: Path) -> bool:
    """Un gabarit valide contient au minimum file_types.yaml."""
    return path.is_dir() and (path / "file_types.yaml").exists()


def _count_classes(path: Path) -> int:
    """Nombre de classes déclarées dans file_types.yaml."""
    ft_path = path / "file_types.yaml"
    if not ft_path.exists():
        return 0
    try:
        with open(ft_path, encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
        return len(data.get("file_types", {}))
    except Exception:
        return 0


# ── Modèles ───────────────────────────────────────────────────────────────────

class GabaritInfo(BaseModel):
    name: str
    path: str
    valid: bool = True
    class_count: int = 0
    description: Optional[str] = ""


class NewGabaritRequest(BaseModel):
    name: str           # nom court → gabarits_dir/<name>/ créé automatiquement
    description: str = ""


class OpenGabaritRequest(BaseModel):
    path: str
    name: str = ""
    description: str = ""


class CloneRequest(BaseModel):
    source_path: str        # chemin du gabarit (ou affaire) source
    dest_path: str = ""     # chemin de destination — si vide : workspace_dir/name
    name: str               # nom de la nouvelle Affaire
    activate: bool = True   # activer immédiatement après clonage


class SaveAsGabaritRequest(BaseModel):
    name: str           # nom court → gabarits_dir/<name>/ créé automatiquement
    description: str = ""


# ── Routes ────────────────────────────────────────────────────────────────────

@router.get("/list", response_model=list[GabaritInfo])
def list_gabarits():
    """Liste les gabarits disponibles.

    Priorité :
    1. Si gabarits_dir défini dans le workspace → scan automatique du dossier
    2. Sinon → fallback sur .gabarits.json (liste explicite)
    """
    gabarits_dir = get_gabarits_dir()

    if gabarits_dir and gabarits_dir.is_dir():
        # Scan automatique : chaque sous-dossier valide = un gabarit
        result = []
        for sub in sorted(gabarits_dir.iterdir()):
            if sub.is_dir() and _is_valid_gabarit(sub):
                # Lire description depuis un éventuel gabarit.json
                desc = ""
                meta_file = sub / "gabarit.json"
                if meta_file.exists():
                    try:
                        import json as _json
                        meta = _json.loads(meta_file.read_text(encoding="utf-8"))
                        desc = meta.get("description", "")
                    except Exception:
                        pass
                result.append(GabaritInfo(
                    name=sub.name,
                    path=str(sub),
                    valid=True,
                    class_count=_count_classes(sub),
                    description=desc,
                ))
        return result

    # Fallback : liste explicite .gabarits.json
    known = _load_known()
    result = []
    for g in known:
        p = Path(g["path"])
        result.append(GabaritInfo(
            name=g["name"],
            path=g["path"],
            valid=_is_valid_gabarit(p),
            class_count=_count_classes(p),
            description=g.get("description", ""),
        ))
    return result


@router.post("/new", response_model=GabaritInfo, status_code=201)
def new_gabarit(body: NewGabaritRequest):
    """Crée un nouveau Gabarit vide sous gabarits_dir/<name>/."""
    gabarits_dir = get_gabarits_dir()
    if not gabarits_dir:
        raise HTTPException(422, "gabarits_dir non configuré — définissez-le dans N1 (Workspace)")

    if not body.name or any(c in _INVALID_NAME_CHARS for c in body.name):
        raise HTTPException(422, f"Nom invalide : '{body.name}' (caractères interdits: \\ / : * ? \" < > |)")

    path = gabarits_dir / body.name

    if path.exists() and any(path.iterdir()):
        raise HTTPException(409, f"Le dossier '{path}' existe déjà et n'est pas vide")

    path.mkdir(parents=True, exist_ok=True)

    with open(path / "file_types.yaml", "w", encoding="utf-8") as f:
        yaml.dump({"file_types": {}}, f, allow_unicode=True, default_flow_style=False)

    for fname, init in [
        ("schema_relations.json", {"version": "1", "relations": []}),
        ("functions.json",        {"version": "1", "functions": []}),
        ("templates.json",        {"version": "1", "templates": []}),
        ("namespaces.json",       {"version": "1", "namespaces": []}),
    ]:
        with open(path / fname, "w", encoding="utf-8") as f:
            json.dump(init, f, ensure_ascii=False, indent=2)

    return GabaritInfo(name=body.name, path=str(path), valid=True, class_count=0,
                       description=body.description)


@router.post("/open", response_model=GabaritInfo)
def open_gabarit(body: OpenGabaritRequest):
    """Référence un dossier existant comme Gabarit."""
    path = Path(body.path)
    if not path.is_absolute():
        path = _PROJECT_ROOT / body.path

    if not path.exists():
        raise HTTPException(404, f"Dossier introuvable : {path}")
    if not _is_valid_gabarit(path):
        raise HTTPException(422, f"Gabarit invalide : file_types.yaml manquant dans '{path}'")

    name = body.name or path.name
    known = _load_known()
    if not any(Path(k["path"]).resolve() == path.resolve() for k in known):
        known.append({"name": name, "path": str(path), "description": body.description})
        _save_known(known)

    return GabaritInfo(name=name, path=str(path), valid=True,
                       class_count=_count_classes(path), description=body.description)


@router.post("/from-current", response_model=GabaritInfo, status_code=201)
def save_current_as_gabarit(body: SaveAsGabaritRequest):
    """Exporte le schéma N1 de l'Affaire active comme nouveau Gabarit."""
    source = get_active_config()
    if not _is_valid_gabarit(source):
        raise HTTPException(422, "L'affaire active n'a pas de file_types.yaml valide")

    gabarits_dir = get_gabarits_dir()
    if not gabarits_dir:
        raise HTTPException(422, "gabarits_dir non configuré — définissez-le dans N1 (Workspace)")

    if not body.name or any(c in _INVALID_NAME_CHARS for c in body.name):
        raise HTTPException(422, f"Nom invalide : '{body.name}'")

    dest = gabarits_dir / body.name

    if dest.exists() and any(dest.iterdir()):
        raise HTTPException(409, f"Le dossier '{dest}' existe déjà et n'est pas vide")

    dest.mkdir(parents=True, exist_ok=True)

    for fname in _N1_FILES:
        src_file = source / fname
        if src_file.exists():
            shutil.copy2(src_file, dest / fname)

    return GabaritInfo(name=body.name, path=str(dest), valid=True,
                       class_count=_count_classes(dest), description=body.description)


@router.post("/clone", response_model=GabaritInfo, status_code=201)
def clone_gabarit(body: CloneRequest):
    """Clone un Gabarit (ou une Affaire) vers une nouvelle Affaire indépendante."""
    source = Path(body.source_path)
    if not source.is_absolute():
        source = _PROJECT_ROOT / body.source_path

    if not source.exists():
        raise HTTPException(404, f"Source introuvable : {source}")
    if not _is_valid_gabarit(source):
        raise HTTPException(422, f"Source invalide : file_types.yaml manquant dans '{source}'")

    # dest_path optionnel — défaut : workspace_dir/name
    if body.dest_path:
        dest = Path(body.dest_path)
        if not dest.is_absolute():
            dest = _PROJECT_ROOT / body.dest_path
    else:
        ws_dir = get_workspace_dir()
        if not ws_dir:
            raise HTTPException(422, "dest_path non fourni et workspace_dir non défini dans le workspace global")
        dest = ws_dir / body.name

    if dest.exists() and any(dest.iterdir()):
        raise HTTPException(409, f"Le dossier '{dest}' existe déjà et n'est pas vide")

    dest.mkdir(parents=True, exist_ok=True)

    # 1. Copier les fichiers N1 depuis la source
    for fname in _N1_FILES:
        src_file = source / fname
        if src_file.exists():
            shutil.copy2(src_file, dest / fname)

    # 2. Créer les fichiers N2 vides (sauf s'ils viennent de la source)
    for fname, init_data in _N2_INIT.items():
        dest_file = dest / fname
        if not dest_file.exists():
            with open(dest_file, "w", encoding="utf-8") as f:
                json.dump(init_data, f, ensure_ascii=False, indent=2)

    # 3. Enregistrer comme écosystème N2 connu
    from web.registry_app.api.ecosystem_manager import _load_known as _load_eco, _save_known as _save_eco
    eco_known = _load_eco()
    if not any(Path(k["path"]).resolve() == dest.resolve() for k in eco_known):
        eco_known.append({"name": body.name, "path": str(dest)})
        _save_eco(eco_known)

    # 4. Activer si demandé
    if body.activate:
        set_active_config(dest)

    return GabaritInfo(name=body.name, path=str(dest), valid=True,
                       class_count=_count_classes(dest))


@router.delete("/{gabarit_name}", status_code=204)
def delete_gabarit(gabarit_name: str):
    """Retire un Gabarit de la liste (ne supprime PAS les fichiers)."""
    known = _load_known()
    if not any(k["name"] == gabarit_name for k in known):
        raise HTTPException(404, f"Gabarit '{gabarit_name}' introuvable")
    _save_known([k for k in known if k["name"] != gabarit_name])
