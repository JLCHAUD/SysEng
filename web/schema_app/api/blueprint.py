"""Endpoint Blueprint N1 — graphe Cytoscape Classes + Relations."""
from fastapi import APIRouter

from web.schema_app.services.config_service import load_file_types, load_relations

router = APIRouter()


@router.get("")
def get_blueprint():
    ft = load_file_types()
    rels = load_relations()

    classes = [
        {
            "id": fid,
            "label": v.get("label", fid),
            "owner_function": v.get("owner_role") or v.get("owner_function", ""),
            "description": v.get("description", ""),
            "min_sheets": v.get("min_sheets") or v.get("required_sheets", []),
            "push_prefix": v.get("push_prefix", ""),
        }
        for fid, v in ft.items()
    ]

    return {"classes": classes, "relations": rels}
