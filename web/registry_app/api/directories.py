"""Gestion du répertoire Posts — N2 Registry Populator.

Un seul répertoire par Affaire : posts_dir (chemin absolu des fichiers Excel Posts).
Stockage : directories.json dans le dossier config actif N2.
"""
from pathlib import Path
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from web.registry_app.services.config_service import load_dirs, save_dirs

router = APIRouter()


class DirsConfig(BaseModel):
    posts_dir: str = ""    # chemin absolu — peut pointer vers OneDrive ou tout autre emplacement


def get_posts_base_path() -> Path | None:
    """Retourne le chemin absolu du répertoire Posts.
    Retourne None si posts_dir n'est pas défini.
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
    if body.posts_dir:
        p = Path(body.posts_dir)
        if not p.is_absolute():
            raise HTTPException(400, "Répertoire Posts : le chemin doit être absolu")
        if p.exists() and not p.is_dir():
            raise HTTPException(400, "Posts : ce chemin n'est pas un répertoire")
        # chemin non existant accepté (OneDrive non encore synchronisé)

    save_dirs(body.model_dump())
    return body


@router.delete("", status_code=204)
def reset_dirs():
    """Remet les répertoires à zéro."""
    save_dirs({"posts_dir": ""})
