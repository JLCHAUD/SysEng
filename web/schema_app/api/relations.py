"""CRUD Relations père/fils — schema_relations.json.

Chaque Relation peut contenir des `flux` : liens de tables entre la Classe
parente et la Classe enfant. Ces flux servent de patrons pour :
  - Le générateur MXL (instructions PULL/COLLECT commentées en template)
  - La pré-population du Tissage N2 (future feature)
"""
import uuid
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from typing import Optional
from pydantic import model_validator

from web.schema_app.services.config_service import load_relations, save_relations

router = APIRouter()


class FluxEntry(BaseModel):
    id: str = ""
    table: str = ""            # nom de la table (ou scalaire) échangée
    source: str = "child"      # "child" | "parent" — qui possède la donnée
    mode: str = "PULL"         # PULL | COLLECT | PUSH
    description: str = ""
    # Ancien format (migration automatique)
    from_table: str = ""
    to_table:   str = ""

    @model_validator(mode="after")
    def _migrate(self):
        """Migre l'ancien format {from_table, to_table} → {table}."""
        if not self.table and self.from_table:
            self.table = self.from_table
        return self


class RelationPF(BaseModel):
    id: str = ""
    parent_class: str
    child_class: str
    qualifier: str = "TYPICAL"    # PRESCRIBED | TYPICAL
    cardinality: str = "1..N"
    description: str = ""
    flux: list[FluxEntry] = []    # liens de tables (Tissage N1)


# ── CRUD Relations ─────────────────────────────────────────────────────────────

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
            # Préserver les flux existants si le body n'en envoie pas
            if not updated.get("flux"):
                updated["flux"] = r.get("flux", [])
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


# ── CRUD Flux d'une Relation ───────────────────────────────────────────────────

def _get_rel(rel_id: str, rels: list) -> tuple[int, dict]:
    for i, r in enumerate(rels):
        if r["id"] == rel_id:
            return i, r
    raise HTTPException(404, "Relation non trouvée")


@router.get("/{rel_id}/flux", response_model=list[FluxEntry])
def list_flux(rel_id: str):
    _, r = _get_rel(rel_id, load_relations())
    return [FluxEntry(**f) for f in r.get("flux", [])]


@router.post("/{rel_id}/flux", response_model=FluxEntry, status_code=201)
def add_flux(rel_id: str, body: FluxEntry):
    rels = load_relations()
    i, r = _get_rel(rel_id, rels)
    if not body.id:
        body.id = "flux-" + str(uuid.uuid4())[:8]
    r.setdefault("flux", []).append(body.model_dump())
    rels[i] = r
    save_relations(rels)
    return body


@router.delete("/{rel_id}/flux/{flux_id}", status_code=204)
def delete_flux(rel_id: str, flux_id: str):
    rels = load_relations()
    i, r = _get_rel(rel_id, rels)
    flux = r.get("flux", [])
    new_flux = [f for f in flux if f["id"] != flux_id]
    if len(new_flux) == len(flux):
        raise HTTPException(404, "Flux non trouvé")
    r["flux"] = new_flux
    rels[i] = r
    save_relations(rels)
