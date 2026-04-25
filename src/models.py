from dataclasses import dataclass, field
from datetime import date
from enum import Enum
from typing import List, Optional


# ─── Enums ────────────────────────────────────────────────────────────────────

class StatutUO(str, Enum):
    BROUILLON  = "BROUILLON"
    EN_COURS   = "EN_COURS"
    CLOTUREE   = "CLOTUREE"
    ARCHIVEE   = "ARCHIVEE"
    VISIBLE    = "VISIBLE"


class StatutActivite(str, Enum):
    EN_COURS        = "EN_COURS"
    ANNULEE         = "ANNULEE"
    TERMINEE        = "TERMINEE"
    NON_APPLICABLE  = "NON_APPLICABLE"


class StatutLivrable(str, Enum):
    A_FAIRE  = "A_FAIRE"
    EN_COURS = "EN_COURS"
    LIVRE    = "LIVRE"
    VALIDE   = "VALIDE"


class Role(str, Enum):
    INGENIEUR_SYS       = "ingenieur_sys"
    PILOTE_TECH         = "pilote_tech"
    ENGAGEMENT_MGR      = "engagement_mgr"
    IT_MANAGER          = "it_manager"
    DONNEUR_ORDRE       = "donneur_ordre"
    PILOTE_TECH_CLIENT  = "pilote_tech_client"
    RESP_PROJET_CLIENT  = "resp_projet_client"
    EXPERT_CLIENT       = "expert_client"
    FOURNISSEUR         = "fournisseur"


class TypeFiltre(str, Enum):
    UO              = "uo"
    INGENIEUR       = "ingenieur"
    PROJET          = "projet"
    SYSTEME         = "systeme"
    PROJET_SYSTEME  = "projet+systeme"
    ALL             = "ALL"


class NiveauAcces(str, Enum):
    READ_WRITE      = "read/write"
    READ            = "read"
    READ_FILTERED   = "read_filtered"
    READ_SUMMARY    = "read_summary"
    ADMIN           = "admin"


class DirectionPasserelle(str, Enum):
    PULL = "pull"
    PUSH = "push"


class TypePasserelle(str, Enum):
    CELL        = "CELL"
    CELL_NUM    = "CELL_NUM"
    CELL_DATE   = "CELL_DATE"
    CELL_PCT    = "CELL_PCT"
    TABLE_COL   = "TABLE_COL"
    TABLE_ROW   = "TABLE_ROW"
    TABLE_FULL  = "TABLE_FULL"
    COMPUTED    = "COMPUTED"
    REF         = "REF"


class ScopePasserelle(str, Enum):
    GLOBAL   = "GLOBAL"
    LOCAL    = "LOCAL"
    COMPUTED = "COMPUTED"


# ─── Acteurs ──────────────────────────────────────────────────────────────────

@dataclass
class Actor:
    name: str
    role: str
    email: str = ""


@dataclass
class ProfilActeur:
    """Profil d'accès d'un utilisateur dans l'écosystème."""
    id: str
    nom: str
    role: Role
    filtre_type: TypeFiltre
    filtre_valeur: str          # "UO-001,UO-002" ou "Alice Dubois" ou "MI20_RATP" ou "ALL"
    acces: NiveauAcces
    email: str = ""


# ─── Référentiels ─────────────────────────────────────────────────────────────

@dataclass
class Activity:
    id: str
    name: str
    default_hours: float
    start_date: Optional[date] = None
    end_date: Optional[date] = None
    progress_pct: float = 0.0
    statut: StatutActivite = StatutActivite.EN_COURS
    heures_realisees: float = 0.0
    allocated_hours: Optional[float] = None

    def effective_hours(self) -> float:
        return self.allocated_hours if self.allocated_hours is not None else self.default_hours


@dataclass
class Deliverable:
    id: str
    name: str
    due_date: Optional[date] = None
    date_reelle: Optional[date] = None
    status: StatutLivrable = StatutLivrable.A_FAIRE


@dataclass
class UOType:
    id: str
    name: str
    activities: List[Activity] = field(default_factory=list)
    deliverables: List[Deliverable] = field(default_factory=list)


@dataclass
class System:
    id: str
    name: str
    rex_prefill: List[str] = field(default_factory=list)


@dataclass
class Project:
    id: str
    name: str
    actors: List[Actor] = field(default_factory=list)


# ─── Instance UO ──────────────────────────────────────────────────────────────

@dataclass
class UOInstance:
    id: str
    uo_type_id: str
    system_id: str
    project_id: str
    engineer_name: str
    total_hours: float
    start_date: date
    end_date: date
    statut: StatutUO = StatutUO.BROUILLON
    degrade: bool = False           # marqueur UO★ (périmètre réduit / non standard)
    degrade_note: str = ""
    owner_id: Optional[str] = None  # ID acteur propriétaire (référence acteurs.json)

    # Références résolues (remplies par config_loader)
    uo_type: Optional[UOType] = None
    system: Optional[System] = None
    project: Optional[Project] = None


# ─── Passerelle ───────────────────────────────────────────────────────────────

@dataclass
class ReglePasserelle:
    type: TypePasserelle
    scope: ScopePasserelle
    nom_global: str = ""
    nom_local: str = ""
    feuille: str = ""
    tableau: str = ""
    cle: str = ""
    colonnes: str = ""          # liste séparée par virgules
    cellule: str = ""
    direction: DirectionPasserelle = DirectionPasserelle.PULL
    formule: str = ""


@dataclass
class Passerelle:
    version: str                            # ex. "1" ou "1-MOD"
    regles: List[ReglePasserelle] = field(default_factory=list)

    @property
    def est_modifiee(self) -> bool:
        return self.version.upper().endswith("-MOD")

    @property
    def version_num(self) -> str:
        return self.version.replace("-MOD", "").replace("-mod", "")


# ─── Registre ─────────────────────────────────────────────────────────────────

@dataclass
class EntreeRegistre:
    id: str
    type_fichier: str       # uo_instance | referentiel_uo | referentiel_projet | cockpit | pilote | consolidation | client
    chemin: str
    synchro_periodicite: str    # quotidien | hebdomadaire | manuel
    derniere_synchro: Optional[str] = None
    statut_dernier_synchro: Optional[str] = None
    genere_par_script: bool = True
    owner_id: Optional[str] = None   # ID acteur propriétaire (référence acteurs.json)
