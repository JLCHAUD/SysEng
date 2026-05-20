"""API Workspace — N2 Registry Populator.

N2 est propriétaire de workspace_dir uniquement.
gabarits_dir est géré par N1 : affiché en lecture seule, préservé lors de la mise à jour.
"""
from pathlib import Path
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from web.workspace_service import load_workspace, save_workspace

router = APIRouter()


class WorkspaceResponse(BaseModel):
    """Réponse GET : les deux champs pour affichage."""
    gabarits_dir: str = ""
    workspace_dir: str = ""


class WorkspaceDirUpdate(BaseModel):
    """Corps PUT : uniquement workspace_dir (propriété N2)."""
    workspace_dir: str = ""


@router.get("", response_model=WorkspaceResponse)
def get_workspace():
    data = load_workspace()
    return WorkspaceResponse(
        gabarits_dir=data.get("gabarits_dir", ""),
        workspace_dir=data.get("workspace_dir", ""),
    )


@router.put("", response_model=WorkspaceResponse)
def update_workspace(body: WorkspaceDirUpdate):
    if body.workspace_dir:
        p = Path(body.workspace_dir)
        if not p.is_absolute():
            raise HTTPException(400, "Workspace : le chemin doit être absolu")
        if p.exists() and not p.is_dir():
            raise HTTPException(400, "Workspace : ce chemin n'est pas un répertoire")

    # Mise à jour partielle — préserver gabarits_dir géré par N1
    current = load_workspace()
    current["workspace_dir"] = body.workspace_dir
    save_workspace(current)
    return WorkspaceResponse(
        gabarits_dir=current.get("gabarits_dir", ""),
        workspace_dir=body.workspace_dir,
    )


@router.delete("", status_code=204)
def reset_workspace():
    # Préserver gabarits_dir lors du reset N2
    current = load_workspace()
    current["workspace_dir"] = ""
    save_workspace(current)
