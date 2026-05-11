"""Génère le contenu MXL d'un fichier depuis sa config (tables + hierarchy)."""
from fastapi import APIRouter, HTTPException
from web.services.config_service import (
    load_registre, load_file_types, load_tables, load_hierarchy, load_ecosystem
)

router = APIRouter()

COL_TYPE_MAP = {
    "KEY": "KEY", "string": "string", "float": "float", "int": "int",
    "date": "date", "pct": "pct", "bool": "bool",
}


def _generate_mxl(file_id: str) -> str:
    registre = {f["id"]: f for f in load_registre()}
    if file_id not in registre:
        raise HTTPException(status_code=404, detail=f"Fichier {file_id} non trouvé dans le registre")

    file_rec = registre[file_id]
    file_types = load_file_types()
    ft = file_types.get(file_rec["type_fichier"], {})
    tables_data = load_tables().get("tables", {})
    hierarchy = load_hierarchy()
    eco = load_ecosystem()

    lines = []

    # ── En-tête ────────────────────────────────────────────────────────────────
    lines += [
        f"FILE_TYPE  {file_rec['type_fichier']}",
        f"FILE_ID    {file_id}",
        f"VERSION    1",
        "",
    ]

    # ── DEF / COL — tables configurées manuellement ───────────────────────────
    file_tables = [t for t in tables_data.values() if t.get("file_id") == file_id]

    for tbl in file_tables:
        lines.append(f"DEF  {tbl['table_name']}  SHEET={tbl['sheet']}")
        for col in tbl.get("columns", []):
            parts = [f"  COL  {col['name']}"]
            parts.append(f"TYPE={COL_TYPE_MAP.get(col['col_type'], col['col_type'])}")
            if col.get("header"):
                parts.append(f'HEADER="{col["header"]}"')
            if col.get("write"):
                parts.append(f"WRITE={col['write']}")
            if col.get("is_key") or col.get("col_type") == "KEY":
                parts.append("KEY")
            lines.append("  ".join(parts))
        lines.append("")

    # ── Tables découvertes dans ecosystem.json ────────────────────────────────
    eco_tables = [t for t in eco.get("tables", {}).values() if t.get("source_file_id") == file_id]
    manual_names = {t["table_name"] for t in file_tables}

    for tbl in eco_tables:
        if tbl.get("table_name") in manual_names:
            continue
        lines.append(f"DEF  {tbl['table_name']}  SHEET={tbl.get('source_sheet','')}")
        for cname, col in tbl.get("columns", {}).items():
            ctype = COL_TYPE_MAP.get(col.get("col_type","string"), "string")
            hdr = f'HEADER="{col["header"]}"' if col.get("header") else ""
            wrt = f"WRITE={col['write']}" if col.get("write") else ""
            key = "KEY" if col.get("col_type") == "KEY" else ""
            parts = [p for p in [f"  COL  {cname}", f"TYPE={ctype}", hdr, wrt, key] if p]
            lines.append("  ".join(parts))
        lines.append("")

    # ── PULL edges depuis ecosystem.json ──────────────────────────────────────
    pulls = [e for e in eco.get("edges", []) if e["edge_type"] == "PULL" and f"{file_id}::" in e["to_node"]]
    if pulls:
        for e in pulls:
            src = e["from_node"].replace("store::", "")
            tgt = e["to_node"].split("::")[1]
            mode = f"  MODE={e['mode']}" if e.get("mode") else ""
            lines.append(f"PULL  {tgt}  FROM={src}{mode}")
        lines.append("")

    # ── PUSH edges depuis ecosystem.json ──────────────────────────────────────
    pushes = [e for e in eco.get("edges", []) if e["edge_type"] == "PUSH" and f"{file_id}::" in e["from_node"]]
    if pushes:
        for e in pushes:
            src = e["from_node"].split("::")[1].lstrip("$")
            tgt = e["to_node"].replace("store::", "")
            lines.append(f"PUSH  {src}  TO={tgt}")
        lines.append("")

    # ── LIST declarations (hierarchy) ─────────────────────────────────────────
    file_lists = [l for l in hierarchy.get("lists", []) if l.get("owner_file_id") == file_id]
    for lst in file_lists:
        line = f"LIST  {lst['list_name']}  FORM={lst.get('form','TABLE')}"
        if lst.get("source_table"):
            line += f"  SOURCE={lst['source_table']}"
        if lst.get("filter_type"):
            line += f"  FILTER={lst['filter_type']}"
        if lst.get("filter_where"):
            line += f"  WHERE={lst['filter_where']}"
        lines.append(line)
    if file_lists:
        lines.append("")

    # ── COLLECT declarations (hierarchy) ──────────────────────────────────────
    file_collects = [c for c in hierarchy.get("collects", []) if c.get("owner_file_id") == file_id]
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

    # ── Variables COMPUTE depuis ecosystem.json ───────────────────────────────
    eco_vars = [v for v in eco.get("variables", {}).values() if v.get("source_file_id") == file_id]
    for var in eco_vars:
        var_name = var["id"].split(".")[-1]
        lines.append(f"COMPUTE  {var_name}  FORMULA={var.get('formula','')}")
    if eco_vars:
        lines.append("")

    return "\n".join(lines).strip()


@router.get("/{file_id}")
def generate_mxl(file_id: str):
    mxl = _generate_mxl(file_id)
    return {"file_id": file_id, "mxl": mxl}


@router.get("/{file_id}/text", response_class=None)
def generate_mxl_text(file_id: str):
    from fastapi.responses import PlainTextResponse
    mxl = _generate_mxl(file_id)
    return PlainTextResponse(mxl, media_type="text/plain")
