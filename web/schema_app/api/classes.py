"""CRUD Classes — file_types.yaml + tables std (tables.json)."""
import uuid
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from web.schema_config import SchemaConfigService


# ── Modèles (module-level — importables par templates.py) ─────────────────────

class ColumnMin(BaseModel):
    name: str
    col_type: str = "string"
    header: str = ""
    write: str = ""
    is_key: bool = False
    required: bool = True
    description: str = ""


class TableStd(BaseModel):
    name: str
    sheet: str
    columns: list[ColumnMin] = []
    description: str = ""


class FieldMin(BaseModel):
    name: str
    label: str = ""
    field_type: str = "string"
    nature: str = "identitaire"
    source: str = "user_input"
    required: bool = True
    pushable: bool = False
    ref_class: str = ""
    description: str = ""


class ClassSchema(BaseModel):
    id: str
    label: str
    description: str = ""
    owner_function: str = ""
    min_sheets: list[str] = []
    optional_sheets: list[str] = []
    allowed_namespaces: list[str] = []
    push_prefix: str = ""
    template: str = ""
    min_fields: list[FieldMin] = []
    std_tables: list[TableStd] = []
    schema_version: int = 1


# ── Factory ────────────────────────────────────────────────────────────────────

def make_router(cfg: SchemaConfigService) -> APIRouter:

    def _ft_to_class(fid: str, ft: dict) -> ClassSchema:
        tables_data = cfg.load_tables().get("tables", {})
        std_tables = [
            TableStd(
                name=t.get("table_name", t.get("name", "")),
                sheet=t.get("sheet", ""),
                description=t.get("description", ""),
                columns=[
                    ColumnMin(**{k: v for k, v in c.items() if k in ColumnMin.model_fields})
                    for c in t.get("columns", [])
                ],
            )
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
            schema_version=ft.get("schema_version", 1),
        )

    def _class_to_ft(cls: ClassSchema, existing_version: int = 0, increment: bool = False) -> dict:
        version = (existing_version + 1) if increment else max(existing_version, 1)
        return {
            "label": cls.label,
            "description": cls.description,
            "owner_role": cls.owner_function,
            "min_sheets": cls.min_sheets,
            "required_sheets": cls.min_sheets,
            "optional_sheets": cls.optional_sheets,
            "allowed_namespaces": cls.allowed_namespaces,
            "push_prefix": cls.push_prefix,
            "template": cls.template,
            "min_fields": [f.model_dump() for f in cls.min_fields],
            "schema_version": version,
        }

    def _save_std_tables(class_id: str, tables: list[TableStd]) -> None:
        data = cfg.load_tables()
        data.setdefault("tables", {})
        prefix = f"__class__{class_id}"
        data["tables"] = {k: v for k, v in data["tables"].items() if v.get("file_id") != prefix}
        for tbl in tables:
            tid = f"{prefix}.{tbl.name}"
            data["tables"][tid] = {
                "id": tid,
                "file_id": prefix,
                "table_name": tbl.name,
                "sheet": tbl.sheet,
                "description": tbl.description,
                "columns": [c.model_dump() for c in tbl.columns],
            }
        cfg.save_tables(data)

    router = APIRouter()

    @router.get("", response_model=list[ClassSchema])
    def list_classes():
        return [_ft_to_class(fid, ft) for fid, ft in cfg.load_file_types().items()]

    @router.get("/{class_id}", response_model=ClassSchema)
    def get_class(class_id: str):
        ft = cfg.load_file_types()
        if class_id not in ft:
            raise HTTPException(404, "Classe non trouvée")
        return _ft_to_class(class_id, ft[class_id])

    @router.post("", response_model=ClassSchema, status_code=201)
    def create_class(body: ClassSchema):
        ft = cfg.load_file_types()
        if body.id in ft:
            raise HTTPException(409, "ID déjà existant")
        ft[body.id] = _class_to_ft(body, existing_version=0, increment=False)
        cfg.save_file_types(ft)
        _save_std_tables(body.id, body.std_tables)
        return _ft_to_class(body.id, ft[body.id])

    @router.put("/{class_id}", response_model=ClassSchema)
    def update_class(class_id: str, body: ClassSchema):
        ft = cfg.load_file_types()
        if class_id not in ft:
            raise HTTPException(404, "Classe non trouvée")
        existing_version = ft[class_id].get("schema_version", 1)
        ft[class_id] = _class_to_ft(body, existing_version=existing_version, increment=True)
        cfg.save_file_types(ft)
        _save_std_tables(class_id, body.std_tables)
        return _ft_to_class(class_id, ft[class_id])

    @router.delete("/{class_id}", status_code=204)
    def delete_class(class_id: str):
        ft = cfg.load_file_types()
        if class_id not in ft:
            raise HTTPException(404, "Classe non trouvée")
        del ft[class_id]
        cfg.save_file_types(ft)
        _save_std_tables(class_id, [])

    return router
