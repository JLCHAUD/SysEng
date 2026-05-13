"""CRUD Templates de Classe — templates.json."""
import uuid
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from web.schema_app.services.config_service import (
    load_templates, save_templates, load_file_types,
)
from web.schema_app.api.classes import ColumnMin, TableStd

router = APIRouter()


class TemplateSchema(BaseModel):
    id: str = ""
    label: str
    class_id: str
    description: str = ""
    extra_sheets: list[str] = []
    field_defaults: dict = {}          # {field_name: default_value}
    std_tables: list[TableStd] = []    # surcharge / ajout vs Classe
    mxl_defaults: dict = {}            # push_prefix, pull_keys…
    source_file: str = ""


@router.get("", response_model=list[TemplateSchema])
def list_templates():
    return [TemplateSchema(**t) for t in load_templates()]


@router.get("/{tpl_id}", response_model=TemplateSchema)
def get_template(tpl_id: str):
    tpls = load_templates()
    t = next((t for t in tpls if t["id"] == tpl_id), None)
    if not t:
        raise HTTPException(404, "Template non trouvé")
    return TemplateSchema(**t)


@router.post("", response_model=TemplateSchema, status_code=201)
def create_template(body: TemplateSchema):
    if not body.id:
        body.id = "tpl-" + str(uuid.uuid4())[:8]
    tpls = load_templates()
    if any(t["id"] == body.id for t in tpls):
        raise HTTPException(409, "ID déjà existant")
    if body.class_id and body.class_id not in load_file_types():
        raise HTTPException(422, f"Classe '{body.class_id}' introuvable")
    tpls.append(body.model_dump())
    save_templates(tpls)
    return body


@router.put("/{tpl_id}", response_model=TemplateSchema)
def update_template(tpl_id: str, body: TemplateSchema):
    tpls = load_templates()
    for i, t in enumerate(tpls):
        if t["id"] == tpl_id:
            updated = body.model_dump()
            updated["id"] = tpl_id
            tpls[i] = updated
            save_templates(tpls)
            return TemplateSchema(**updated)
    raise HTTPException(404, "Template non trouvé")


@router.delete("/{tpl_id}", status_code=204)
def delete_template(tpl_id: str):
    tpls = load_templates()
    new = [t for t in tpls if t["id"] != tpl_id]
    if len(new) == len(tpls):
        raise HTTPException(404, "Template non trouvé")
    save_templates(new)


