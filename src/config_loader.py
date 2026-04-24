import json
from copy import deepcopy
from datetime import date
from pathlib import Path
from typing import Dict, List

from src.models import (
    Activity, Actor, Deliverable, EntreeRegistre, ProfilActeur,
    Project, StatutUO, System, UOInstance, UOType,
    Role, TypeFiltre, NiveauAcces,
)

CONFIG_DIR = Path(__file__).parent.parent / "config"


def _load_json(filename: str) -> dict | list:
    with open(CONFIG_DIR / filename, encoding="utf-8") as f:
        return json.load(f)


# ─── Référentiels ─────────────────────────────────────────────────────────────

def load_uo_types() -> Dict[str, UOType]:
    raw = _load_json("types_uo.json")
    result = {}
    for type_id, data in raw.items():
        activities = [
            Activity(id=a["id"], name=a["name"], default_hours=a["default_hours"])
            for a in data.get("activities", [])
        ]
        deliverables = [
            Deliverable(id=d["id"], name=d["name"])
            for d in data.get("deliverables", [])
        ]
        result[type_id] = UOType(
            id=type_id, name=data["name"],
            activities=activities, deliverables=deliverables,
        )
    return result


def load_systems() -> Dict[str, System]:
    raw = _load_json("systemes.json")
    return {
        sys_id: System(
            id=sys_id, name=data["name"],
            rex_prefill=data.get("rex_prefill", []),
        )
        for sys_id, data in raw.items()
    }


def load_projects() -> Dict[str, Project]:
    raw = _load_json("projets.json")
    result = {}
    for proj_id, data in raw.items():
        actors = [
            Actor(name=a["name"], role=a["role"], email=a.get("email", ""))
            for a in data.get("actors", [])
        ]
        result[proj_id] = Project(id=proj_id, name=data["name"], actors=actors)
    return result


# ─── Instances UO ─────────────────────────────────────────────────────────────

def load_uo_instances() -> List[UOInstance]:
    raw = _load_json("uo_instances.json")
    uo_types = load_uo_types()
    systems = load_systems()
    projects = load_projects()

    instances = []
    for item in raw:
        uo_type = uo_types.get(item["uo_type_id"])
        system = systems.get(item["system_id"])
        project = projects.get(item["project_id"])

        resolved_type = None
        if uo_type:
            resolved_type = UOType(
                id=uo_type.id, name=uo_type.name,
                activities=deepcopy(uo_type.activities),
                deliverables=deepcopy(uo_type.deliverables),
            )

        instances.append(UOInstance(
            id=item["id"],
            uo_type_id=item["uo_type_id"],
            system_id=item["system_id"],
            project_id=item["project_id"],
            engineer_name=item["engineer_name"],
            total_hours=item["total_hours"],
            start_date=date.fromisoformat(item["start_date"]),
            end_date=date.fromisoformat(item["end_date"]),
            statut=StatutUO(item.get("statut", "BROUILLON")),
            degrade=item.get("degrade", False),
            degrade_note=item.get("degrade_note", ""),
            uo_type=resolved_type,
            system=system,
            project=project,
        ))
    return instances


# ─── Acteurs ──────────────────────────────────────────────────────────────────

def load_acteurs() -> List[ProfilActeur]:
    raw = _load_json("acteurs.json")
    return [
        ProfilActeur(
            id=a["id"],
            nom=a["nom"],
            role=Role(a["role"]),
            filtre_type=TypeFiltre(a["filtre_type"]),
            filtre_valeur=a["filtre_valeur"],
            acces=NiveauAcces(a["acces"]),
            email=a.get("email", ""),
        )
        for a in raw
    ]


# ─── Registre ─────────────────────────────────────────────────────────────────

def load_registre() -> List[EntreeRegistre]:
    raw = _load_json("registre.json")
    return [
        EntreeRegistre(
            id=f["id"],
            type_fichier=f["type_fichier"],
            chemin=f["chemin"],
            synchro_periodicite=f["synchro_periodicite"],
            derniere_synchro=f.get("derniere_synchro"),
            statut_dernier_synchro=f.get("statut_dernier_synchro"),
            genere_par_script=f.get("genere_par_script", True),
        )
        for f in raw.get("fichiers", [])
    ]


def save_registre(entrees: List[EntreeRegistre]) -> None:
    """Persiste les mises à jour du registre (dernière synchro, statut)."""
    raw = _load_json("registre.json")
    index = {e.id: e for e in entrees}
    for f in raw.get("fichiers", []):
        e = index.get(f["id"])
        if e:
            f["derniere_synchro"] = e.derniere_synchro
            f["statut_dernier_synchro"] = e.statut_dernier_synchro
    with open(CONFIG_DIR / "registre.json", "w", encoding="utf-8") as fh:
        json.dump(raw, fh, ensure_ascii=False, indent=2)
