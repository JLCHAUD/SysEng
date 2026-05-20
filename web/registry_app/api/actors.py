import re
from fastapi import APIRouter, HTTPException
from web.schemas.models import Actor, ActorCreate
from web.registry_app.services.config_service import load_acteurs, save_acteurs

router = APIRouter()


def _next_id(acteurs: list) -> str:
    """Calcule le prochain ID libre au format USRxxx (incrémental)."""
    pattern = re.compile(r'^USR(\d+)$', re.IGNORECASE)
    max_n = 0
    for a in acteurs:
        m = pattern.match(a.get("id", ""))
        if m:
            max_n = max(max_n, int(m.group(1)))
    return f"USR{max_n + 1:03d}"


@router.get("/next-id")
def get_next_id():
    """Retourne le prochain ID disponible (non utilisé)."""
    return {"id": _next_id(load_acteurs())}


@router.get("", response_model=list[Actor])
def list_actors():
    # On filtre les éventuels acteurs avec ID vide (données corrompues)
    return [Actor(**a) for a in load_acteurs() if a.get("id", "").strip()]


@router.get("/{actor_id}", response_model=Actor)
def get_actor(actor_id: str):
    for a in load_acteurs():
        if a["id"] == actor_id:
            return Actor(**a)
    raise HTTPException(status_code=404, detail="Acteur non trouvé")


@router.post("", response_model=Actor, status_code=201)
def create_actor(body: ActorCreate):
    if not body.id or not body.id.strip():
        raise HTTPException(status_code=422, detail="L'ID ne peut pas être vide")
    body.id = body.id.strip()
    acteurs = load_acteurs()
    if any(a["id"] == body.id for a in acteurs):
        raise HTTPException(status_code=409, detail=f"L'ID '{body.id}' existe déjà")
    # On profite du save pour purger les IDs vides
    acteurs = [a for a in acteurs if a.get("id", "").strip()]
    entry = body.model_dump()
    acteurs.append(entry)
    save_acteurs(acteurs)
    return Actor(**entry)


@router.put("/{actor_id}", response_model=Actor)
def update_actor(actor_id: str, body: ActorCreate):
    acteurs = load_acteurs()
    for i, a in enumerate(acteurs):
        if a["id"] == actor_id:
            updated = {**a, **body.model_dump()}
            updated["id"] = actor_id          # on ne laisse pas changer l'ID via PUT
            acteurs[i] = updated
            save_acteurs(acteurs)
            return Actor(**updated)
    raise HTTPException(status_code=404, detail="Acteur non trouvé")


@router.delete("/purge-empty", status_code=204)
def purge_empty_ids():
    """Supprime tous les acteurs avec un ID vide ou blanc."""
    acteurs = load_acteurs()
    cleaned = [a for a in acteurs if a.get("id", "").strip()]
    save_acteurs(cleaned)


@router.delete("/{actor_id}", status_code=204)
def delete_actor(actor_id: str):
    acteurs = load_acteurs()
    new = [a for a in acteurs if a["id"] != actor_id]
    if len(new) == len(acteurs):
        raise HTTPException(status_code=404, detail="Acteur non trouvé")
    save_acteurs(new)
