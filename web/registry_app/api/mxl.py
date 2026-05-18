"""Génère le contenu MXL d'un fichier depuis sa config (tables + ecosystem + pulls déclarés)."""
from fastapi import APIRouter, HTTPException
from fastapi.responses import PlainTextResponse
from web.registry_app.services.config_service import (
    load_registre, load_tables, load_hierarchy, load_ecosystem
)

router = APIRouter()


# ── Helpers ───────────────────────────────────────────────────────────────────

def _var_from_eco_key(eco_key: str) -> str:
    return "$" + eco_key.split(".")[-1]


def _var_from_table_name(table_name: str) -> str:
    name = table_name[3:] if table_name.startswith("Tab") else table_name
    return "$" + name[0].lower() + name[1:] if name else "$unknown"


def _build_sheet_index(tables_data: dict, eco: dict) -> dict:
    index = {}
    for tbl in tables_data.get("tables", {}).values():
        index[tbl["table_name"]] = tbl["sheet"]
    for tbl in eco.get("tables", {}).values():
        index[tbl["table_name"]] = tbl.get("source_sheet", "")
    return index


def _col_line(var: str, col_name: str, col_type: str, header: str, write: str) -> str:
    is_key = col_type == "KEY"
    attrs_parts = []
    if is_key:
        attrs_parts.append("KEY")
    if write:
        attrs_parts.append(f"WRITE={write}")
    if header and header != col_name:
        attrs_parts.append(f'HEADER="{header}"')
    if not attrs_parts:
        return ""
    return f"  COL {var}.{col_name} : {'  '.join(attrs_parts)}"


# ── Générateur principal ───────────────────────────────────────────────────────

def _generate_mxl(file_id: str) -> str:
    registre = {f["id"]: f for f in load_registre()}
    if file_id not in registre:
        raise HTTPException(status_code=404, detail=f"Fichier {file_id} non trouvé dans le registre")

    file_rec = registre[file_id]
    tables_data = load_tables()
    hierarchy = load_hierarchy()
    eco = load_ecosystem()
    sheet_index = _build_sheet_index(tables_data, eco)

    lines = []

    # ── En-tête ───────────────────────────────────────────────────────────────
    lines += [
        f"FILE_TYPE  {file_rec.get('type_fichier') or 'libre'}",
        f"FILE_ID    {file_id}",
        f"VERSION    1",
        "",
    ]

    # ── DEF / COL — tables configurées manuellement (tables.json) ─────────────
    file_tables = [t for t in tables_data.get("tables", {}).values()
                   if t.get("file_id") == file_id]
    manual_table_names = set()

    for tbl in file_tables:
        var = _var_from_table_name(tbl["table_name"])
        manual_table_names.add(tbl["table_name"])
        lines.append(f"DEF {var} = GET_TABLE({tbl['sheet']}, {tbl['table_name']})")
        for col in tbl.get("columns", []):
            col_type = "KEY" if col.get("is_key") else col.get("col_type", "")
            cl = _col_line(var, col["name"], col_type,
                           col.get("header", col["name"]), col.get("write", ""))
            if cl:
                lines.append(cl)
        lines.append("")

    # ── DEF / COL — tables découvertes dans ecosystem.json ────────────────────
    eco_tables = [
        (k, t) for k, t in eco.get("tables", {}).items()
        if t.get("source_file_id") == file_id
        and t.get("table_name") not in manual_table_names
    ]
    for eco_key, tbl in eco_tables:
        var = _var_from_eco_key(eco_key)
        sheet = tbl.get("source_sheet", "")
        table = tbl.get("table_name", "")
        lines.append(f"DEF {var} = GET_TABLE({sheet}, {table})")
        for cname, col in tbl.get("columns", {}).items():
            cl = _col_line(var, cname, col.get("col_type", ""),
                           col.get("header", cname), col.get("write", ""))
            if cl:
                lines.append(cl)
        lines.append("")

    # ── DEF — variables calculées (ecosystem.json) ────────────────────────────
    eco_vars = [
        (k, v) for k, v in eco.get("variables", {}).items()
        if v.get("source_file_id") == file_id
    ]
    for eco_key, var in eco_vars:
        var_name = _var_from_eco_key(eco_key)
        formula = var.get("formula", "")
        lines.append(f"DEF {var_name} = {formula}")
    if eco_vars:
        lines.append("")

    # ── PULL déclarés via Tissage (hierarchy.json["pulls"]) ───────────────────
    declared_pulls = [p for p in hierarchy.get("pulls", [])
                      if p.get("target_file_id") == file_id]
    for p in declared_pulls:
        src_table = p["source_table"]
        tgt_table = p["target_table"]
        tgt_sheet = sheet_index.get(tgt_table, tgt_table)
        cols_part = ""
        if p.get("columns"):
            cols_part = f"  COLS={','.join(p['columns'])}"
        lines.append(f"PULL {p['source_file_id']}::{src_table} -> FILL_TABLE({tgt_sheet}, {tgt_table}){cols_part}")
    if declared_pulls:
        lines.append("")

    # ── PULL edges (ecosystem.json) ───────────────────────────────────────────
    eco_pulls = [e for e in eco.get("edges", [])
                 if e["edge_type"] == "PULL"
                 and e.get("to_node", "").startswith(f"{file_id}::")]
    if eco_pulls:
        for e in eco_pulls:
            src = e["from_node"].replace("store::", "")
            table_name = e["to_node"].split("::", 1)[1]
            sheet = sheet_index.get(table_name, "")
            mode = f"  MODE={e['mode']}" if e.get("mode") else ""
            lines.append(f"PULL {src} -> FILL_TABLE({sheet}, {table_name}){mode}")
        lines.append("")

    # ── PUSH edges (ecosystem.json) ───────────────────────────────────────────
    pushes = [e for e in eco.get("edges", [])
              if e["edge_type"] == "PUSH"
              and e.get("from_node", "").startswith(f"{file_id}::")]
    if pushes:
        for e in pushes:
            var_name = e["from_node"].split("::", 1)[1]
            store_path = e["to_node"].replace("store::", "")
            lines.append(f"PUSH {var_name} -> {store_path}")
        lines.append("")

    # ── LIST declarations (hierarchy.json) ────────────────────────────────────
    file_lists = [l for l in hierarchy.get("lists", [])
                  if l.get("owner_file_id") == file_id]
    for lst in file_lists:
        line = f"LIST  {lst['list_name']}  FORM={lst.get('form', 'TABLE')}"
        if lst.get("source_table"):
            line += f"  SOURCE={lst['source_table']}"
        if lst.get("filter_type"):
            line += f"  FILTER={lst['filter_type']}"
        if lst.get("filter_where"):
            line += f"  WHERE={lst['filter_where']}"
        lines.append(line)
    if file_lists:
        lines.append("")

    # ── COLLECT declarations (hierarchy.json) ─────────────────────────────────
    file_collects = [c for c in hierarchy.get("collects", [])
                     if c.get("owner_file_id") == file_id]
    for col in file_collects:
        line = f"COLLECT  {col['source_table']}  FROM_LIST={col['list_name']}  INTO={col['target_table']}"
        if col.get("where_clause"):
            line += f"  WHERE={col['where_clause']}"
        if col.get("cols_filter"):
            line += f"  COLS={col['cols_filter']}"
        if col.get("with_fields"):
            line += f"  WITH={col['with_fields']}"
        lines.append(line)
    if file_collects:
        lines.append("")

    return "\n".join(lines).strip()


# ── Routes ────────────────────────────────────────────────────────────────────

@router.get("/{file_id}")
def generate_mxl(file_id: str):
    mxl = _generate_mxl(file_id)
    return {"file_id": file_id, "mxl": mxl}


@router.get("/{file_id}/text")
def generate_mxl_text(file_id: str):
    mxl = _generate_mxl(file_id)
    return PlainTextResponse(mxl, media_type="text/plain")
