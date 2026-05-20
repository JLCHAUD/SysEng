"""Gestion des répertoires Affaire et Posts — N2 Registry Populator.

MVP :
  - Un répertoire Affaire (racine du projet Excel)
  - Un répertoire Posts (toujours sous Affaire), relatif à affaire_dir

Stockage : directories.json dans le dossier config actif N2.
"""
from pathlib import Path
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from web.registry_app.services.config_service import load_dirs, save_dirs

router = APIRouter()


class DirsConfig(BaseModel):
    affaire_dir: str = ""
    posts_dir: str = ""    # chemin absolu — peut pointer vers OneDrive ou tout autre emplacement


def get_posts_base_path() -> Path | None:
    """Retourne le chemin absolu du répertoire Posts.
    Retourne None si posts_dir n'est pas défini.
    Utilisé par le moteur sync pour résoudre les chemins relatifs des Posts.
    """
    d = load_dirs()
    posts = d.get("posts_dir", "")
    if not posts:
        return None
    return Path(posts)


# ── Routes ─────────────────────────────────────────────────────────────────────

@router.get("", response_model=DirsConfig)
def get_dirs():
    return DirsConfig(**load_dirs())


@router.put("", response_model=DirsConfig)
def update_dirs(body: DirsConfig):
    errors = []

    # Valider affaire_dir
    if body.affaire_dir:
        p = Path(body.affaire_dir)
        if not p.exists():
            errors.append(f"Répertoire Affaire introuvable : {body.affaire_dir}")
        elif not p.is_dir():
            errors.append(f"Affaire : ce chemin n'est pas un répertoire")

    # Valider posts_dir — doit être absolu
    if body.posts_dir:
        p = Path(body.posts_dir)
        if not p.is_absolute():
            errors.append("Répertoire Posts : le chemin doit être absolu (ex: C:\\...  ou  /home/...)")
        elif p.exists() and not p.is_dir():
            errors.append(f"Posts : ce chemin n'est pas un répertoire")
        # Note : on accepte un chemin qui n'existe pas encore (ex: OneDrive non encore synchronisé)

    if errors:
        raise HTTPException(400, " | ".join(errors))

    save_dirs(body.model_dump())
    return body


@router.delete("", status_code=204)
def reset_dirs():
    """Remet les répertoires à zéro."""
    save_dirs({"affaire_dir": "", "posts_dir": ""})
