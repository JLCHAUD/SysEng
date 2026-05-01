import json
from copy import deepcopy
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Dict, List, Optional, Union

import yaml

from src.models import (
    Activity, Actor, Deliverable, EntreeRegistre, ProfilActeur,
    Project, StatutUO, System, UOInstance, UOType,
    Role, TypeFiltre, NiveauAcces,
)

CONFIG_DIR = Path(__file__).parent.parent / "config"


# ─── Ownership ────────────────────────────────────────────────────────────────

@dataclass
class OwnerRoleViolation:
    """Incohérence entre le rôle attendu pour un type de fichier et celui de l'owner."""
    file_id: str
    type_fichier: str
    owner_id: str
    owner_nom: str
    owner_role: str
    expected_role: str

    def __str__(self) -> str:
        return (
            f"{self.file_id} ({self.type_fichier}) : owner {self.owner_nom} "
            f"a le rôle '{self.owner_role}' mais '{self.expected_role}' est attendu"
        )


def _load_json(filename: str) -> Union[dict, list]:
    with open(CONFIG_DIR / filename, encoding="utf-8") as f:
        return json.load(f)


def _load_yaml(filename: str) -> dict:
    with open(CONFIG_DIR / filename, encoding="utf-8") as f:
        return yaml.safe_load(f)


def load_file_types() -> Dict[str, dict]:
    """Charge file_types.yaml → dict {type_id: {label, owner_role, ...}}."""
    raw = _load_yaml("file_types.yaml")
    return raw.get("file_types", {})


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
            owner_id=item.get("owner_id"),
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
            owner_id=f.get("owner_id"),
        )
        for f in raw.get("fichiers", [])
    ]


def validate_owner_roles(
    registre: Optional[List[EntreeRegistre]] = None,
) -> List[OwnerRoleViolation]:
    """
    Vérifie que chaque fichier du registre a un owner dont le rôle correspond
    au rôle attendu par son type de fichier (owner_role dans file_types.yaml).

    Retourne la liste des violations détectées (liste vide = tout est cohérent).
    """
    if registre is None:
        registre = load_registre()

    acteurs_index: Dict[str, ProfilActeur] = {a.id: a for a in load_acteurs()}
    file_types = load_file_types()
    violations: List[OwnerRoleViolation] = []

    for entree in registre:
        type_def = file_types.get(entree.type_fichier)
        if not type_def:
            continue  # type inconnu → ignoré (pas de contrainte)

        expected_role: Optional[str] = type_def.get("owner_role")
        if not expected_role:
            continue  # pas de contrainte de rôle pour ce type

        if not entree.owner_id:
            # Owner absent mais rôle attendu → violation légère (pas bloquant)
            violations.append(OwnerRoleViolation(
                file_id=entree.id,
                type_fichier=entree.type_fichier,
                owner_id="",
                owner_nom="<non assigné>",
                owner_role="",
                expected_role=expected_role,
            ))
            continue

        acteur = acteurs_index.get(entree.owner_id)
        if not acteur:
            violations.append(OwnerRoleViolation(
                file_id=entree.id,
                type_fichier=entree.type_fichier,
                owner_id=entree.owner_id,
                owner_nom="<inconnu>",
                owner_role="",
                expected_role=expected_role,
            ))
            continue

        if acteur.role.value != expected_role:
            violations.append(OwnerRoleViolation(
                file_id=entree.id,
                type_fichier=entree.type_fichier,
                owner_id=entree.owner_id,
                owner_nom=acteur.nom,
                owner_role=acteur.role.value,
                expected_role=expected_role,
            ))

    return violations


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
