from fastapi import APIRouter, HTTPException
from web.schemas.models import FileType, FileTypeBase, FileTypeCreate
from web.services.config_service import load_file_types, save_file_types

router = APIRouter()


def _row_to_model(id: str, data: dict) -> FileType:
    return FileType(id=id, **{k: v for k, v in data.items()})


@router.get("", response_model=list[FileType])
def list_file_types():
    types = load_file_types()
    return [_row_to_model(fid, fdata) for fid, fdata in types.items()]


@router.get("/{type_id}", response_model=FileType)
def get_file_type(type_id: str):
    types = load_file_types()
    if type_id not in types:
        raise HTTPException(status_code=404, detail="Type non trouvé")
    return _row_to_model(type_id, types[type_id])


@router.post("", response_model=FileType, status_code=201)
def create_file_type(body: FileTypeCreate):
    types = load_file_types()
    if body.id in types:
        raise HTTPException(status_code=409, detail="ID déjà existant")
    entry = body.model_dump(exclude={"id"})
    types[body.id] = entry
    save_file_types(types)
    return _row_to_model(body.id, entry)


@router.put("/{type_id}", response_model=FileType)
def update_file_type(type_id: str, body: FileTypeBase):
    types = load_file_types()
    if type_id not in types:
        raise HTTPException(status_code=404, detail="Type non trouvé")
    entry = body.model_dump()
    types[type_id] = entry
    save_file_types(types)
    return _row_to_model(type_id, entry)


@router.delete("/{type_id}", status_code=204)
def delete_file_type(type_id: str):
    types = load_file_types()
    if type_id not in types:
        raise HTTPException(status_code=404, detail="Type non trouvé")
    del types[type_id]
    save_file_types(types)
