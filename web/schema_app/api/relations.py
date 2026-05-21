"""CRUD Relations père/fils — schema_relations.json."""
import uuid
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel, model_validator
from typing import Optional

from web.schema_config import SchemaConfigService


class FluxEntry(BaseModel):
    id: str = ""
    table: str = ""
    source: str = "child"
    mode: str = "PULL"
    description: str = ""
    from_table: str = ""
    to_table: str = ""

    @model_validator(mode="after")
    def _migrate(self):
        if not self.table and self.from_table:
            self.table = self.from_table
        return self


class RelationPF(BaseModel):
    id: str = ""
    parent_class: str
    child_class: str
    qualifier: str = "TYPICAL"
    cardinality: str = "1..N"
    description: str = ""
    flux: list[FluxEntry] = []


def make_router(cfg: SchemaConfigService) -> APIRouter:
    router = APIRouter()

    @router.get("", response_model=list[RelationPF])
    def list_relations():
        return [RelationPF(**r) for r in cfg.load_relations()]

    @router.get("/{rel_id}", response_model=RelationPF)
    def get_relation(rel_id: str):
        r = next((r for r in cfg.load_relations() if r["id"] == rel_id), None)
        if not r:
            raise HTTPException(404, "Relation non trouvée")
        return RelationPF(**r)

    @router.post("", response_model=RelationPF, status_code=201)
    def create_relation(body: RelationPF):
        rels = cfg.load_relations()
        if not body.id:
            body.id = "rel-" + str(uuid.uuid4())[:8]
        if any(r["id"] == body.id for r in rels):
            raise HTTPException(409, "ID déjà existant")
        rels.append(body.model_dump())
        cfg.save_relations(rels)
        return body

    @router.put("/{rel_id}", response_model=RelationPF)
    def update_relation(rel_id: str, body: RelationPF):
        rels = cfg.load_relations()
        for i, r in enumerate(rels):
            if r["id"] == rel_id:
                updated = body.model_dump()
                updated["id"] = rel_id
                if not updated.get("flux"):
                    updated["flux"] = r.get("flux", [])
                rels[i] = updated
                cfg.save_relations(rels)
                return RelationPF(**updated)
        raise HTTPException(404, "Relation non trouvée")

    @router.delete("/{rel_id}", status_code=204)
    def delete_relation(rel_id: str):
        rels = cfg.load_relations()
        new = [r for r in rels if r["id"] != rel_id]
        if len(new) == len(rels):
            raise HTTPException(404, "Relation non trouvée")
        cfg.save_relations(new)

    def _get_rel(rel_id: str, rels: list) -> tuple[int, dict]:
        for i, r in enumerate(rels):
            if r["id"] == rel_id:
                return i, r
        raise HTTPException(404, "Relation non trouvée")

    @router.get("/{rel_id}/flux", response_model=list[FluxEntry])
    def list_flux(rel_id: str):
        _, r = _get_rel(rel_id, cfg.load_relations())
        return [FluxEntry(**f) for f in r.get("flux", [])]

    @router.post("/{rel_id}/flux", response_model=FluxEntry, status_code=201)
    def add_flux(rel_id: str, body: FluxEntry):
        rels = cfg.load_relations()
        i, r = _get_rel(rel_id, rels)
        if not body.id:
            body.id = "flux-" + str(uuid.uuid4())[:8]
        r.setdefault("flux", []).append(body.model_dump())
        rels[i] = r
        cfg.save_relations(rels)
        return body

    @router.delete("/{rel_id}/flux/{flux_id}", status_code=204)
    def delete_flux(rel_id: str, flux_id: str):
        rels = cfg.load_relations()
        i, r = _get_rel(rel_id, rels)
        flux = r.get("flux", [])
        new_flux = [f for f in flux if f["id"] != flux_id]
        if len(new_flux) == len(flux):
            raise HTTPException(404, "Flux non trouvé")
        r["flux"] = new_flux
        rels[i] = r
        cfg.save_relations(rels)

    return router
