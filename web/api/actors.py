from fastapi import APIRouter, HTTPException
from web.schemas.models import Actor, ActorCreate
from web.services.config_service import load_acteurs, save_acteurs

router = APIRouter()


@router.get("", response_model=list[Actor])
def list_actors():
    return [Actor(**a) for a in load_acteurs()]


@router.get("/{actor_id}", response_model=Actor)
def get_actor(actor_id: str):
    for a in load_acteurs():
        if a["id"] == actor_id:
            return Actor(**a)
    raise HTTPException(status_code=404, detail="Acteur non trouvé")


@router.post("", response_model=Actor, status_code=201)
def create_actor(body: ActorCreate):
    acteurs = load_acteurs()
    if any(a["id"] == body.id for a in acteurs):
        raise HTTPException(status_code=409, detail="ID déjà existant")
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
            acteurs[i] = updated
            save_acteurs(acteurs)
            return Actor(**updated)
    raise HTTPException(status_code=404, detail="Acteur non trouvé")


@router.delete("/{actor_id}", status_code=204)
def delete_actor(actor_id: str):
    acteurs = load_acteurs()
    new = [a for a in acteurs if a["id"] != actor_id]
    if len(new) == len(acteurs):
        raise HTTPException(status_code=404, detail="Acteur non trouvé")
    save_acteurs(new)
