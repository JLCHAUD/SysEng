"""API Workspace — N1 Schema Designer.

Permet à l'architecte de définir :
  - gabarits_dir : bibliothèque de gabarits (dossier contenant des sous-dossiers gabarits)
  - workspace_dir : répertoire par défaut pour créer de nouvelles Affaires

Ces paramètres sont partagés avec N2 via .exosync_workspace.json (racine projet).
"""
from pathlib import Path
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from web.workspace_service import load_workspace, save_workspace

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
