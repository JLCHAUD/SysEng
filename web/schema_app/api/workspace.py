"""API Workspace — N1 Schema Designer.

N1 est propriétaire de gabarits_dir uniquement.
workspace_dir est géré par N2 et préservé lors de la mise à jour.
"""
from pathlib import Path
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from web.workspace_service import load_workspace, save_workspace

router = APIRouter()


class WorkspaceConfig(BaseModel):
    gabarits_dir: str = ""


@router.get("", response_model=WorkspaceConfig)
def get_workspace():
    data = load_workspace()
    return WorkspaceConfig(gabarits_dir=data.get("gabarits_dir", ""))


@router.put("", response_model=WorkspaceConfig)
def update_workspace(body: WorkspaceConfig):
    if body.gabarits_dir:
        p = Path(body.gabarits_dir)
        if not p.is_absolute():
            raise HTTPException(400, "Gabarits : le chemin doit être absolu")
        if p.exists() and not p.is_dir():
            raise HTTPException(400, "Gabarits : ce chemin n'est pas un répertoire")

    # Mise à jour partielle — préserver workspace_dir géré par N2
    current = load_workspace()
    current["gabarits_dir"] = body.gabarits_dir
    save_workspace(current)
    return body


@router.delete("", status_code=204)
def reset_workspace():
    # Préserver workspace_dir lors du reset N1
    current = load_workspace()
    current["gabarits_dir"] = ""
    save_workspace(current)
