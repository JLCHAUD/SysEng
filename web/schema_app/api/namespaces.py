"""CRUD Namespaces — namespaces.json.

Un namespace = préfixe d'identifiant déclaré au niveau de l'écosystème,
puis assigné aux Classes via allowed_namespaces.
"""
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from web.schema_config import SchemaConfigService


class NamespaceSchema(BaseModel):
    id: str
    label: str = ""
    prefix: str = ""
    description: str = ""


def make_router(cfg: SchemaConfigService) -> APIRouter:
    router = APIRouter()

    @router.get("", response_model=list[NamespaceSchema])
    def list_namespaces():
        return [NamespaceSchema(**ns) for ns in cfg.load_namespaces()]

    @router.post("", response_model=NamespaceSchema, status_code=201)
    def create_namespace(body: NamespaceSchema):
        nss = cfg.load_namespaces()
        if any(ns["id"] == body.id for ns in nss):
            raise HTTPException(409, "ID déjà existant")
        nss.append(body.model_dump())
        cfg.save_namespaces(nss)
        return body

    @router.put("/{ns_id}", response_model=NamespaceSchema)
    def update_namespace(ns_id: str, body: NamespaceSchema):
        nss = cfg.load_namespaces()
        idx = next((i for i, ns in enumerate(nss) if ns["id"] == ns_id), None)
        if idx is None:
            raise HTTPException(404, "Namespace non trouvé")
        nss[idx] = body.model_dump()
        cfg.save_namespaces(nss)
        return body

    @router.delete("/{ns_id}", status_code=204)
    def delete_namespace(ns_id: str):
        nss = cfg.load_namespaces()
        new_list = [ns for ns in nss if ns["id"] != ns_id]
        if len(new_list) == len(nss):
            raise HTTPException(404, "Namespace non trouvé")
        cfg.save_namespaces(new_list)

    return router
