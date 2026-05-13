"""API Niveau 1 — Schéma d'écosystème : Classes, Relations P/F, Fonctions."""
import uuid
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from typing import Optional

from web.services.config_service import (
    load_file_types, save_file_types,
    load_relations, save_relations,
    load_functions, save_functions,
    load_tables, save_tables,
    load_templates, save_templates,
)

router = APIRouter()


# ── Modèles ────────────────────────────────────────────────────────────────────

class ColumnMin(BaseModel):
    name: str
    col_type: str = "string"
    header: str = ""
    write: str = ""
    is_key: bool = False
    required: bool = True
    description: str = ""


class TableMin(BaseModel):
    name: str
    sheet: str
    columns: list[ColumnMin] = []
    description: str = ""


class FieldMin(BaseModel):
    name: str
    label: str = ""
    field_type: str = "string"    # string | int | float | date | bool | ref
    nature: str = "identitaire"   # identitaire | operationnel | parametre
    source: str = "user_input"    # user_input | system | computed | reference
    required: bool = True
    pushable: bool = False
    ref_class: str = ""           # si source=reference
    description: str = ""


class ClassSchema(BaseModel):
    id: str
    label: str
    description: str = ""
    owner_function: str
    min_sheets: list[str] = []
    optional_sheets: list[str] = []
    allowed_namespaces: list[str] = []
    push_prefix: str = ""
    template: str = ""
    min_fields: list[FieldMin] = []   # champs scalaires (DEF)
    std_tables: list[TableMin] = []   # tables minimales standard


class TemplateSchema(BaseModel):
    id: str
    label: str
    class_id: str
    description: str = ""
    extra_sheets: list[str] = []
    field_defaults: dict = {}         # valeurs par défaut pour min_fields
    std_tables: list[TableMin] = []   # surcharge / ajout vs Classe
    mxl_defaults: dict = {}           # push_prefix, pull_keys, etc.
    source_file: str = ""             # chemin vers fichier Excel de base


class RelationPF(BaseModel):
    id: str
    parent_class: str
    child_class: str
    qualifier: str = "TYPICAL"        # PRESCRIBED | TYPICAL
    cardinality: str = "1..N"
    description: str = ""


class FunctionDef(BaseModel):
    id: str
    label: str
    description: str = ""
    side: str = "interne"   # interne | externe


# ── Helpers ────────────────────────────────────────────────────────────────────

def _ft_to_class(fid: str, ft: dict) -> ClassSchema:
    """Convertit un file_type YAML en ClassSchema enrichi avec ses tables standard."""
    tables_data = load_tables().get("tables", {})
    std_tables = [
        TableMin(**t)
        for t in tables_data.values()
        if t.get("file_id") == f"__class__{fid}"
    ]
    min_fields = [FieldMin(**f) for f in ft.get("min_fields", [])]
    return ClassSchema(
        id=fid,
        label=ft.get("label", ""),
        description=ft.get("description", ""),
        owner_function=ft.get("owner_role") or ft.get("owner_function", ""),
        min_sheets=ft.get("min_sheets") or ft.get("required_sheets", []),
        optional_sheets=ft.get("optional_sheets", []),
        allowed_namespaces=ft.get("allowed_namespaces", []),
        push_prefix=ft.get("push_prefix", ""),
        template=ft.get("template", ""),
        min_fields=min_fields,
        std_tables=std_tables,
    )


def _class_to_ft(cls: ClassSchema) -> dict:
    return {
        "label": cls.label,
        "description": cls.description,
        "owner_role": cls.owner_function,
        "required_sheets": cls.min_sheets,
        "min_sheets": cls.min_sheets,
        "optional_sheets": cls.optional_sheets,
        "allowed_namespaces": cls.allowed_namespaces,
        "push_prefix": cls.push_prefix,
        "template": cls.template,
        "min_fields": [f.model_dump() for f in cls.min_fields],
    }


def _save_std_tables(class_id: str, tables: list[TableMin]) -> None:
    data = load_tables()
    # Supprimer les anciennes tables de cette Classe
    data["tables"] = {
        k: v for k, v in data.get("tables", {}).items()
        if v.get("file_id") != f"__class__{class_id}"
    }
    for tbl in tables:
        tid = f"__class__{class_id}.{tbl.name}"
        data["tables"][tid] = {
            "id": tid,
            "file_id": f"__class__{class_id}",
            "table_name": tbl.name,
            "sheet": tbl.sheet,
            "description": tbl.description,
            "columns": [c.model_dump() for c in tbl.columns],
        }
    save_tables(data)


# ── Classes ────────────────────────────────────────────────────────────────────

@router.get("/classes", response_model=list[ClassSchema])
def list_classes():
    return [_ft_to_class(fid, ft) for fid, ft in load_file_types().items()]


@router.get("/classes/{class_id}", response_model=ClassSchema)
def get_class(class_id: str):
    ft = load_file_types()
    if class_id not in ft:
        raise HTTPException(404, "Classe non trouvée")
    return _ft_to_class(class_id, ft[class_id])


@router.post("/classes", response_model=ClassSchema, status_code=201)
def create_class(body: ClassSchema):
    ft = load_file_types()
    if body.id in ft:
        raise HTTPException(409, "ID déjà existant")
    ft[body.id] = _class_to_ft(body)
    save_file_types(ft)
    _save_std_tables(body.id, body.std_tables)
    return _ft_to_class(body.id, ft[body.id])


@router.put("/classes/{class_id}", response_model=ClassSchema)
def update_class(class_id: str, body: ClassSchema):
    ft = load_file_types()
    if class_id not in ft:
        raise HTTPException(404, "Classe non trouvée")
    ft[class_id] = _class_to_ft(body)
    save_file_types(ft)
    _save_std_tables(class_id, body.std_tables)
    return _ft_to_class(class_id, ft[class_id])


@router.delete("/classes/{class_id}", status_code=204)
def delete_class(class_id: str):
    ft = load_file_types()
    if class_id not in ft:
        raise HTTPException(404, "Classe non trouvée")
    del ft[class_id]
    save_file_types(ft)
    _save_std_tables(class_id, [])  # nettoie les tables


# ── Relations P/F ──────────────────────────────────────────────────────────────

@router.get("/relations", response_model=list[RelationPF])
def list_relations():
    return [RelationPF(**r) for r in load_relations()]


@router.post("/relations", response_model=RelationPF, status_code=201)
def create_relation(body: RelationPF):
    rels = load_relations()
    if not body.id:
        body.id = "rel-" + str(uuid.uuid4())[:8]
    if any(r["id"] == body.id for r in rels):
        raise HTTPException(409, "ID déjà existant")
    rels.append(body.model_dump())
    save_relations(rels)
    return body


@router.put("/relations/{rel_id}", response_model=RelationPF)
def update_relation(rel_id: str, body: RelationPF):
    rels = load_relations()
    for i, r in enumerate(rels):
        if r["id"] == rel_id:
            rels[i] = body.model_dump()
            save_relations(rels)
            return body
    raise HTTPException(404, "Relation non trouvée")


@router.delete("/relations/{rel_id}", status_code=204)
def delete_relation(rel_id: str):
    rels = load_relations()
    new = [r for r in rels if r["id"] != rel_id]
    if len(new) == len(rels):
        raise HTTPException(404, "Relation non trouvée")
    save_relations(new)


# ── Fonctions ──────────────────────────────────────────────────────────────────

@router.get("/functions", response_model=list[FunctionDef])
def list_functions():
    return [FunctionDef(**f) for f in load_functions()]


@router.post("/functions", response_model=FunctionDef, status_code=201)
def create_function(body: FunctionDef):
    funcs = load_functions()
    if any(f["id"] == body.id for f in funcs):
        raise HTTPException(409, "ID déjà existant")
    funcs.append(body.model_dump())
    save_functions(funcs)
    return body


@router.put("/functions/{func_id}", response_model=FunctionDef)
def update_function(func_id: str, body: FunctionDef):
    funcs = load_functions()
    for i, f in enumerate(funcs):
        if f["id"] == func_id:
            funcs[i] = body.model_dump()
            save_functions(funcs)
            return body
    raise HTTPException(404, "Fonction non trouvée")


@router.delete("/functions/{func_id}", status_code=204)
def delete_function(func_id: str):
    funcs = load_functions()
    new = [f for f in funcs if f["id"] != func_id]
    if len(new) == len(funcs):
        raise HTTPException(404, "Fonction non trouvée")
    save_functions(new)


# ── Templates ──────────────────────────────────────────────────────────────────

@router.get("/templates", response_model=list[TemplateSchema])
def list_templates():
    return [TemplateSchema(**t) for t in load_templates()]


@router.get("/templates/{tpl_id}", response_model=TemplateSchema)
def get_template(tpl_id: str):
    tpls = load_templates()
    t = next((t for t in tpls if t["id"] == tpl_id), None)
    if not t:
        raise HTTPException(404, "Template non trouvé")
    return TemplateSchema(**t)


@router.post("/templates", response_model=TemplateSchema, status_code=201)
def create_template(body: TemplateSchema):
    tpls = load_templates()
    if not body.id:
        body.id = "tpl-" + str(uuid.uuid4())[:8]
    if any(t["id"] == body.id for t in tpls):
        raise HTTPException(409, "ID déjà existant")
    # vérifier que la Classe existe
    if body.class_id not in load_file_types():
        raise HTTPException(422, f"Classe '{body.class_id}' introuvable")
    tpls.append(body.model_dump())
    save_templates(tpls)
    return body


@router.put("/templates/{tpl_id}", response_model=TemplateSchema)
def update_template(tpl_id: str, body: TemplateSchema):
    tpls = load_templates()
    for i, t in enumerate(tpls):
        if t["id"] == tpl_id:
            tpls[i] = body.model_dump()
            save_templates(tpls)
            return body
    raise HTTPException(404, "Template non trouvé")


@router.delete("/templates/{tpl_id}", status_code=204)
def delete_template(tpl_id: str):
    tpls = load_templates()
    new = [t for t in tpls if t["id"] != tpl_id]
    if len(new) == len(tpls):
        raise HTTPException(404, "Template non trouvé")
    save_templates(new)


# ── Blueprint data ─────────────────────────────────────────────────────────────

@router.get("/blueprint")
def get_blueprint():
    classes = list_classes()
    relations = list_relations()
    # Compter les Posts par Classe (depuis registre)
    from web.services.config_service import load_registre
    registre = load_registre()
    post_counts = {}
    for f in registre:
        cid = f.get("type_fichier", "")
        post_counts[cid] = post_counts.get(cid, 0) + 1

    return {
        "classes": [
            {**c.model_dump(), "post_count": post_counts.get(c.id, 0)}
            for c in classes
        ],
        "relations": [r.model_dump() for r in relations],
    }
