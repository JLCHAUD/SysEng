"""Génère le template MXL d'une Classe (N1).

Contrairement au générateur N2 (qui produit le MXL d'un Post instance),
ce générateur produit le *template manifeste* qu'un architecte placerait
dans une feuille _Manifeste pour un fichier de cette Classe.

Sortie type pour la Classe `uo_instance` :
  FILE_TYPE  uo_instance
  FILE_ID    UO_INSTANCE_ID

  # -- Champs scalaires --
  DEF $project_id = GET_CELL(_Manifeste, project_id)
    COL $project_id : HEADER="Projet"

  # -- Tables standard --
  DEF $activites = GET_TABLE(Activites, TabActivites)
    COL $activites.id        : KEY
    COL $activites.label     : HEADER="Activité"
    COL $activites.hours     : WRITE=ingenieur_sys

  # -- PUSH --
  PUSH $activites -> uo.{id}.activites
"""
from fastapi import APIRouter, HTTPException
from fastapi.responses import PlainTextResponse
from pydantic import BaseModel
from typing import Optional

from web.schema_app.services.config_service import load_file_types, load_tables

router = APIRouter()


# ── Helpers ────────────────────────────────────────────────────────────────────

def _col_line(var: str, col_name: str, is_key: bool, header: str, write: str) -> str:
    parts = []
    if is_key:
        parts.append("KEY")
    if write:
        parts.append(f"WRITE={write}")
    if header and header != col_name:
        parts.append(f'HEADER="{header}"')
    if not parts:
        return ""
    return f"  COL {var}.{col_name} : {' '.join(parts)}"


def _generate_class_mxl(class_id: str) -> str:
    ft_all = load_file_types()
    if class_id not in ft_all:
        raise HTTPException(404, f"Classe '{class_id}' introuvable")

    ft = ft_all[class_id]
    push_prefix = ft.get("push_prefix", "")
    min_fields = ft.get("min_fields", [])

    # Tables std liées à cette Classe dans tables.json
    tables_data = load_tables().get("tables", {})
    std_tables = [
        t for t in tables_data.values()
        if t.get("file_id") == f"__class__{class_id}"
    ]

    lines: list[str] = []

    # ── En-tête ────────────────────────────────────────────────────────────────
    placeholder_id = class_id.upper() + "_ID"
    default_style = ft.get("default_style", "default")
    lines += [
        f"FILE_TYPE  {class_id}",
        f"FILE_ID    {placeholder_id}",
        f"VERSION    1",
        f"STYLE:     {default_style}",
        "",
    ]

    # ── Champs scalaires (min_fields → DEF GET_CELL) ───────────────────────────
    if min_fields:
        lines.append("# -- Champs scalaires --")
        for field in min_fields:
            fname = field.get("name", "")
            label = field.get("label", fname)
            ftype = field.get("field_type", "string")
            pushable = field.get("pushable", False)
            var = "$" + fname

            lines.append(f"DEF {var} = GET_CELL(_Manifeste, {fname})")
            attrs = []
            if label and label != fname:
                attrs.append(f'HEADER="{label}"')
            if ftype not in ("string", ""):
                attrs.append(f"TYPE={ftype}")
            if attrs:
                lines.append(f"  COL {var} : {' '.join(attrs)}")

            if pushable and push_prefix:
                store_key = push_prefix.rstrip(".") + "." + fname
                lines.append(f"  PUSH {var} -> {store_key}")

            lines.append("")

    # ── Tables standard (std_tables → DEF GET_TABLE + COL) ────────────────────
    if std_tables:
        lines.append("# -- Tables standard --")
    for tbl in std_tables:
        tname = tbl.get("table_name", "")
        sheet = tbl.get("sheet", tname)
        var_name = tname[3:] if tname.startswith("Tab") else tname
        var = "$" + var_name[0].lower() + var_name[1:] if var_name else "$table"

        lines.append(f"DEF {var} = GET_TABLE({sheet}, {tname})")
        for col in tbl.get("columns", []):
            cname = col.get("name", "")
            cl = _col_line(var, cname, col.get("is_key", False),
                           col.get("header", ""), col.get("write", ""))
            if cl:
                lines.append(cl)

        # PUSH auto si push_prefix défini
        if push_prefix:
            seg = var_name[0].lower() + var_name[1:]
            store_key = push_prefix.rstrip(".") + "." + seg
            lines.append(f"PUSH {var} -> {store_key}")

        lines.append("")

    return "\n".join(lines).strip()


# ── Routes ─────────────────────────────────────────────────────────────────────

@router.get("/generate/{class_id}")
def generate_mxl(class_id: str):
    mxl_text = _generate_class_mxl(class_id)
    return {"class_id": class_id, "mxl": mxl_text}


@router.get("/generate/{class_id}/text", response_class=PlainTextResponse)
def generate_mxl_text(class_id: str):
    return PlainTextResponse(_generate_class_mxl(class_id), media_type="text/plain")


class PreviewRequest(BaseModel):
    class_id: str
    push_prefix: Optional[str] = None
    extra_sheets: list[str] = []


@router.post("/preview")
def preview_mxl(body: PreviewRequest):
    """Génère depuis payload (sans modifier les fichiers)."""
    mxl_text = _generate_class_mxl(body.class_id)
    return {"class_id": body.class_id, "mxl": mxl_text}
