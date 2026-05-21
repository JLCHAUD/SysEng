from fastapi import APIRouter, HTTPException
from web.schemas.models import FileInstance, FileInstanceCreate
from web.registry_app.services.config_service import (
    load_registre, save_registre, load_file_types,
)

router = APIRouter()


def _enrich(post: dict) -> dict:
    type_id = post.get("type_fichier")
    if not type_id:
        return {**post, "schema_outdated": None}
    ft = load_file_types()
    class_version = ft.get(type_id, {}).get("schema_version", 1)
    post_version = post.get("schema_version") or 0
    return {**post, "schema_outdated": post_version < class_version}


@router.get("", response_model=list[FileInstance])
def list_registry():
    return [FileInstance(**_enrich(f)) for f in load_registre()]


@router.get("/{file_id}", response_model=FileInstance)
def get_file(file_id: str):
    for f in load_registre():
        if f["id"] == file_id:
            return FileInstance(**_enrich(f))
    raise HTTPException(status_code=404, detail="Fichier non trouvé")


@router.post("", response_model=FileInstance, status_code=201)
def create_file(body: FileInstanceCreate):
    fichiers = load_registre()
    if any(f["id"] == body.id for f in fichiers):
        raise HTTPException(status_code=409, detail="ID déjà existant")
    entry = body.model_dump()
    fichiers.append(entry)
    save_registre(fichiers)
    return FileInstance(**_enrich(entry))


@router.put("/{file_id}", response_model=FileInstance)
def update_file(file_id: str, body: FileInstanceCreate):
    fichiers = load_registre()
    for i, f in enumerate(fichiers):
        if f["id"] == file_id:
            updated = {**f, **body.model_dump()}
            fichiers[i] = updated
            save_registre(fichiers)
            return FileInstance(**_enrich(updated))
    raise HTTPException(status_code=404, detail="Fichier non trouvé")


@router.delete("/{file_id}", status_code=204)
def delete_file(file_id: str):
    fichiers = load_registre()
    new = [f for f in fichiers if f["id"] != file_id]
    if len(new) == len(fichiers):
        raise HTTPException(status_code=404, detail="Fichier non trouvé")
    save_registre(new)
