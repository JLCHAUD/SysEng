"""CRUD Fonctions Acteurs — functions.json."""
import uuid
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from web.schema_app.services.config_service import load_functions, save_functions

router = APIRouter()


class FunctionDef(BaseModel):
    id: str = ""
    label: str
    description: str = ""
    side: str = "interne"    # interne | externe


@router.get("", response_model=list[FunctionDef])
def list_functions():
    return [FunctionDef(**f) for f in load_functions()]


@router.get("/{func_id}", response_model=FunctionDef)
def get_function(func_id: str):
    funcs = load_functions()
    f = next((f for f in funcs if f["id"] == func_id), None)
    if not f:
        raise HTTPException(404, "Fonction non trouvée")
    return FunctionDef(**f)


@router.post("", response_model=FunctionDef, status_code=201)
def create_function(body: FunctionDef):
    funcs = load_functions()
    if not body.id:
        body.id = "fn-" + str(uuid.uuid4())[:8]
    if any(f["id"] == body.id for f in funcs):
        raise HTTPException(409, "ID déjà existant")
    funcs.append(body.model_dump())
    save_functions(funcs)
    return body


@router.put("/{func_id}", response_model=FunctionDef)
def update_function(func_id: str, body: FunctionDef):
    funcs = load_functions()
    for i, f in enumerate(funcs):
        if f["id"] == func_id:
            updated = body.model_dump()
            updated["id"] = func_id
            funcs[i] = updated
            save_functions(funcs)
            return FunctionDef(**updated)
    raise HTTPException(404, "Fonction non trouvée")


@router.delete("/{func_id}", status_code=204)
def delete_function(func_id: str):
    funcs = load_functions()
    new = [f for f in funcs if f["id"] != func_id]
    if len(new) == len(funcs):
        raise HTTPException(404, "Fonction non trouvée")
    save_functions(new)
