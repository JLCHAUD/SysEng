"""API Workspace — N2 Registry Populator (lecture + écriture partagée avec N1).

Même endpoints que N1, même fichier source (.exosync_workspace.json).
N2 lit le workspace pour trouver les gabarits disponibles.
N2 peut aussi modifier workspace_dir / gabarits_dir si N1 n'est pas ouvert.
"""
from pathlib import Path
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from web.workspace_service import load_workspace, save_workspace, get_gabarits_dir

router = APIRouter()


class WorkspaceConfig(BaseModel):
    gabarits_dir: str = ""
    workspace_dir: str = ""


@router.get("", response_model=WorkspaceConfig)
def get_workspace():
    return WorkspaceConfig(**load_workspace())


@router.put("", response_model=WorkspaceConfig)
def update_workspace(body: WorkspaceConfig):
    errors = []
    for field, label in [("gabarits_dir", "Gabarits"), ("workspace_dir", "Workspace")]:
        val = getattr(body, field)
        if val:
            p = Path(val)
            if not p.is_absolute():
                errors.append(f"{label} : le chemin doit être absolu")
            elif p.exists() and not p.is_dir():
                errors.append(f"{label} : ce chemin n'est pas un répertoire")
    if errors:
        raise HTTPException(400, " | ".join(errors))
    save_workspace(body.model_dump())
    return body


@router.delete("", status_code=204)
def reset_workspace():
    save_workspace({"gabarits_dir": "", "workspace_dir": ""})
