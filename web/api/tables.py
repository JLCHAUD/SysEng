from fastapi import APIRouter, HTTPException
from web.schemas.models import TableDef
from web.services.config_service import load_tables, save_tables, load_ecosystem

router = APIRouter()


def _all_tables(data: dict) -> list[TableDef]:
    out = []
    for tid, t in data.get("tables", {}).items():
        out.append(TableDef(**t))
    return out


@router.get("", response_model=list[TableDef])
def list_tables(file_id: str = None):
    data = load_tables()
    tables = _all_tables(data)
    if file_id:
        tables = [t for t in tables if t.file_id == file_id]
    return tables


@router.get("/{table_id}", response_model=TableDef)
def get_table(table_id: str):
    data = load_tables()
    if table_id not in data.get("tables", {}):
        raise HTTPException(status_code=404, detail="Table non trouvée")
    return TableDef(**data["tables"][table_id])


@router.post("", response_model=TableDef, status_code=201)
def create_table(body: TableDef):
    data = load_tables()
    if body.id in data.get("tables", {}):
        raise HTTPException(status_code=409, detail="ID déjà existant")
    entry = body.model_dump()
    data.setdefault("tables", {})[body.id] = entry
    save_tables(data)
    return TableDef(**entry)


@router.put("/{table_id}", response_model=TableDef)
def update_table(table_id: str, body: TableDef):
    data = load_tables()
    if table_id not in data.get("tables", {}):
        raise HTTPException(status_code=404, detail="Table non trouvée")
    entry = body.model_dump()
    data["tables"][table_id] = entry
    save_tables(data)
    return TableDef(**entry)


@router.delete("/{table_id}", status_code=204)
def delete_table(table_id: str):
    data = load_tables()
    if table_id not in data.get("tables", {}):
        raise HTTPException(status_code=404, detail="Table non trouvée")
    del data["tables"][table_id]
    save_tables(data)


@router.get("/from-ecosystem/{file_id}", response_model=list[TableDef])
def tables_from_ecosystem(file_id: str):
    """Importe les tables découvertes dans ecosystem.json pour un fichier."""
    eco = load_ecosystem()
    result = []
    for tid, t in eco.get("tables", {}).items():
        if t.get("source_file_id") == file_id:
            cols = [
                {"name": n, "col_type": c.get("col_type","string"),
                 "header": c.get("header",""), "write": c.get("write",""),
                 "is_key": c.get("col_type","") == "KEY", "description": c.get("description","")}
                for n, c in t.get("columns", {}).items()
            ]
            result.append(TableDef(
                id=f"{file_id}.{t['table_name']}",
                file_id=file_id,
                table_name=t.get("table_name", tid),
                sheet=t.get("source_sheet", ""),
                columns=cols,
                description=t.get("description", ""),
            ))
    return result
