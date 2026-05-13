"""CRUD Relations père/fils — schema_relations.json."""
import uuid
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from web.schema_app.services.config_service import load_relations, save_relations

router = APIRouter()


class RelationPF(BaseModel):
    id: str = ""
    parent_class: str
    child_class: str
    qualifier: str = "TYPICAL"    # PRESCRIBED | TYPICAL
    cardinality: str = "1..N"
    description: str = ""


@router.get("", response_model=list[RelationPF])
def list_relations():
    return [RelationPF(**r) for r in load_relations()]


@router.get("/{rel_id}", response_model=RelationPF)
def get_relation(rel_id: str):
    rels = load_relations()
    r = next((r for r in rels if r["id"] == rel_id), None)
    if not r:
        raise HTTPException(404, "Relation non trouvée")
    return RelationPF(**r)


@router.post("", response_model=RelationPF, status_code=201)
def create_relation(body: RelationPF):
    rels = load_relations()
    if not body.id:
        body.id = "rel-" + str(uuid.uuid4())[:8]
    if any(r["id"] == body.id for r in rels):
        raise HTTPException(409, "ID déjà existant")
    rels.append(body.model_dump())
    save_relations(rels)
    return body


@router.put("/{rel_id}", response_model=RelationPF)
def update_relation(rel_id: str, body: RelationPF):
    rels = load_relations()
    for i, r in enumerate(rels):
        if r["id"] == rel_id:
            updated = body.model_dump()
            updated["id"] = rel_id
            rels[i] = updated
            save_relations(rels)
            return RelationPF(**updated)
    raise HTTPException(404, "Relation non trouvée")


@router.delete("/{rel_id}", status_code=204)
def delete_relation(rel_id: str):
    rels = load_relations()
    new = [r for r in rels if r["id"] != rel_id]
    if len(new) == len(rels):
        raise HTTPException(404, "Relation non trouvée")
    save_relations(new)
