"""Génère le template MXL d'une Classe (N1).

Produit le *template Manifeste* que l'on placerait dans une feuille _Manifeste
pour un fichier de cette Classe. Délègue la génération à web.mxl_service.
"""
from fastapi import APIRouter, HTTPException
from fastapi.responses import PlainTextResponse
from pydantic import BaseModel
from typing import Optional

from web.schema_app.services.config_service import load_file_types, load_tables, load_relations
from web.mxl_service import build_class_mxl_lines

router = APIRouter()


def _generate_class_mxl(class_id: str) -> str:
    ft_all = load_file_types()
    if class_id not in ft_all:
        raise HTTPException(404, f"Classe '{class_id}' introuvable")

    ft = ft_all[class_id]

    std_tables = [
        t for t in load_tables().get("tables", {}).values()
        if t.get("file_id") == f"__class__{class_id}"
    ]

    placeholder_id = class_id.upper() + "_ID"
    lines = build_class_mxl_lines(
        class_id   = class_id,
        file_id    = placeholder_id,
        ft         = ft,
        std_tables = std_tables,
        relations  = load_relations(),
    )
    return "\n".join(lines).strip()


# ── Routes ─────────────────────────────────────────────────────────────────────

@router.get("/generate/{class_id}")
def generate_mxl(class_id: str):
    mxl_text = _generate_class_mxl(class_id)
    return {"class_id": class_id, "mxl": mxl_text}


@router.get("/generate/{class_id}/text", response_class=PlainTextResponse)
def generate_mxl_text(class_id: str):
    return PlainTextResponse(_generate_class_mxl(class_id), media_type="text/plain")


class PreviewRequest(BaseModel):
    class_id: str
    push_prefix: Optional[str] = None
    extra_sheets: list[str] = []


@router.post("/preview")
def preview_mxl(body: PreviewRequest):
    """Génère depuis payload (sans modifier les fichiers)."""
    mxl_text = _generate_class_mxl(body.class_id)
    return {"class_id": body.class_id, "mxl": mxl_text}
