"""Lecture seule des types de fichiers (Classes N1) — utilisé par le wizard Import Excel N2."""
from fastapi import APIRouter
from web.schemas.models import FileType
from web.services.config_service import load_file_types

router = APIRouter()


@router.get("", response_model=list[FileType])
def list_file_types():
    return [FileType(id=k, **v) for k, v in load_file_types().items()]
