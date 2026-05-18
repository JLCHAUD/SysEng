from fastapi import APIRouter, HTTPException
from web.schemas.models import FileInstance, FileInstanceCreate
from web.services.config_service import load_registre, save_registre

router = APIRouter()


@router.get("", response_model=list[FileInstance])
def list_registry():
    return [FileInstance(**f) for f in load_registre()]


@router.get("/{file_id}", response_model=FileInstance)
def get_file(file_id: str):
    for f in load_registre():
        if f["id"] == file_id:
            return FileInstance(**f)
    raise HTTPException(status_code=404, detail="Fichier non trouvé")


@router.post("", response_model=FileInstance, status_code=201)
def create_file(body: FileInstanceCreate):
    fichiers = load_registre()
    if any(f["id"] == body.id for f in fichiers):
        raise HTTPException(status_code=409, detail="ID déjà existant")
    entry = body.model_dump()
    fichiers.append(entry)
    save_registre(fichiers)
    return FileInstance(**entry)


@router.put("/{file_id}", response_model=FileInstance)
def update_file(file_id: str, body: FileInstanceCreate):
    fichiers = load_registre()
    for i, f in enumerate(fichiers):
        if f["id"] == file_id:
            updated = {**f, **body.model_dump()}
            fichiers[i] = updated
            save_registre(fichiers)
            return FileInstance(**updated)
    raise HTTPException(status_code=404, detail="Fichier non trouvé")


@router.delete("/{file_id}", status_code=204)
def delete_file(file_id: str):
    fichiers = load_registre()
    new = [f for f in fichiers if f["id"] != file_id]
    if len(new) == len(fichiers):
        raise HTTPException(status_code=404, detail="Fichier non trouvé")
    save_registre(new)
