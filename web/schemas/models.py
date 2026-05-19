from pydantic import BaseModel
from typing import Optional


# ── FileType ──────────────────────────────────────────────────────────────────

class FileTypeBase(BaseModel):
    label: str
    description: Optional[str] = ""
    template: Optional[str] = ""
    owner_role: str
    required_sheets: list[str] = []
    optional_sheets: list[str] = []
    allowed_namespaces: list[str] = []
    push_prefix: Optional[str] = ""


class FileTypeCreate(FileTypeBase):
    id: str


class FileType(FileTypeBase):
    id: str

    class Config:
        from_attributes = True


# ── FileInstance (Registre) ────────────────────────────────────────────────────

class FileInstanceBase(BaseModel):
    type_fichier: Optional[str] = None
    chemin: str
    synchro_periodicite: str = "quotidien"
    owner_role: str = ""
    genere_par_script: bool = False


class FileInstanceCreate(FileInstanceBase):
    id: str


class FileInstance(FileInstanceBase):
    id: str
    derniere_synchro: Optional[str] = None
    statut_dernier_synchro: Optional[str] = None

    class Config:
        from_attributes = True


# ── Actor ─────────────────────────────────────────────────────────────────────

class ActorBase(BaseModel):
    nom: str
    role: str
    filtre_type: str = "ALL"
    filtre_valeur: str = "ALL"
    acces: str = "read"
    email: Optional[str] = ""


class ActorCreate(ActorBase):
    id: str


class Actor(ActorBase):
    id: str

    class Config:
        from_attributes = True


# ── TableDef ──────────────────────────────────────────────────────────────────

class ColumnDef(BaseModel):
    name: str
    col_type: str = "string"
    header: str = ""
    write: str = ""
    is_key: bool = False
    description: str = ""


class TableDef(BaseModel):
    id: str
    file_id: str
    table_name: str
    sheet: str
    columns: list[ColumnDef] = []
    description: str = ""


# ── Hierarchy ─────────────────────────────────────────────────────────────────

class ListDeclaration(BaseModel):
    id: Optional[str] = None
    owner_file_id: str
    list_name: str
    form: str = "TABLE"
    source_table: Optional[str] = None
    filter_type: Optional[str] = None
    filter_where: Optional[str] = None


class CollectMapping(BaseModel):
    id: Optional[str] = None
    owner_file_id: str
    source_table: str
    list_name: str
    target_table: str
    where_clause: Optional[str] = None
    cols_filter: Optional[str] = None
    with_fields: Optional[str] = None


# ── Ecosystem Graph ────────────────────────────────────────────────────────────

class EcosystemNode(BaseModel):
    id: str
    label: str
    file_type: str
    status: str
    path: str


class EcosystemEdge(BaseModel):
    edge_type: str
    from_node: str
    to_node: str
    mode: Optional[str] = ""


class EcosystemGraph(BaseModel):
    version: str
    last_scan: Optional[str] = None
    nodes: list[EcosystemNode]
    edges: list[EcosystemEdge]
