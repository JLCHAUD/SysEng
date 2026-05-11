import uuid
from fastapi import APIRouter, HTTPException
from web.schemas.models import ListDeclaration, CollectMapping
from web.services.config_service import load_hierarchy, save_hierarchy

router = APIRouter()


@router.get("")
def get_hierarchy():
    return load_hierarchy()


# ── Lists ──────────────────────────────────────────────────────────────────────

@router.get("/lists", response_model=list[ListDeclaration])
def list_lists():
    return [ListDeclaration(**l) for l in load_hierarchy().get("lists", [])]


@router.post("/lists", response_model=ListDeclaration, status_code=201)
def create_list(body: ListDeclaration):
    h = load_hierarchy()
    if any(l["id"] == body.id for l in h["lists"]):
        raise HTTPException(status_code=409, detail="ID déjà existant")
    entry = body.model_dump()
    if not entry.get("id"):
        entry["id"] = str(uuid.uuid4())[:8]
    h["lists"].append(entry)
    save_hierarchy(h)
    return ListDeclaration(**entry)


@router.put("/lists/{list_id}", response_model=ListDeclaration)
def update_list(list_id: str, body: ListDeclaration):
    h = load_hierarchy()
    for i, l in enumerate(h["lists"]):
        if l["id"] == list_id:
            updated = {**l, **body.model_dump()}
            h["lists"][i] = updated
            save_hierarchy(h)
            return ListDeclaration(**updated)
    raise HTTPException(status_code=404, detail="Liste non trouvée")


@router.delete("/lists/{list_id}", status_code=204)
def delete_list(list_id: str):
    h = load_hierarchy()
    new = [l for l in h["lists"] if l["id"] != list_id]
    if len(new) == len(h["lists"]):
        raise HTTPException(status_code=404, detail="Liste non trouvée")
    h["lists"] = new
    save_hierarchy(h)


# ── Collects ───────────────────────────────────────────────────────────────────

@router.get("/collects", response_model=list[CollectMapping])
def list_collects():
    return [CollectMapping(**c) for c in load_hierarchy().get("collects", [])]


@router.post("/collects", response_model=CollectMapping, status_code=201)
def create_collect(body: CollectMapping):
    h = load_hierarchy()
    if any(c["id"] == body.id for c in h["collects"]):
        raise HTTPException(status_code=409, detail="ID déjà existant")
    entry = body.model_dump()
    if not entry.get("id"):
        entry["id"] = str(uuid.uuid4())[:8]
    h["collects"].append(entry)
    save_hierarchy(h)
    return CollectMapping(**entry)


@router.put("/collects/{collect_id}", response_model=CollectMapping)
def update_collect(collect_id: str, body: CollectMapping):
    h = load_hierarchy()
    for i, c in enumerate(h["collects"]):
        if c["id"] == collect_id:
            updated = {**c, **body.model_dump()}
            h["collects"][i] = updated
            save_hierarchy(h)
            return CollectMapping(**updated)
    raise HTTPException(status_code=404, detail="COLLECT non trouvé")


@router.delete("/collects/{collect_id}", status_code=204)
def delete_collect(collect_id: str):
    h = load_hierarchy()
    new = [c for c in h["collects"] if c["id"] != collect_id]
    if len(new) == len(h["collects"]):
        raise HTTPException(status_code=404, detail="COLLECT non trouvé")
    h["collects"] = new
    save_hierarchy(h)
