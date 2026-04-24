"""
Ecosystem Schema Manager.

Catalogue vivant de toutes les tables et variables connues de l'écosystème.
Alimenté par le parser lors de la lecture des passerelles.
Persisté dans output/ecosystem.json.
"""
import json
from dataclasses import asdict, dataclass, field
from datetime import date
from pathlib import Path
from typing import Dict, List, Optional

ECOSYSTEM_PATH = Path(__file__).parent.parent / "output" / "ecosystem.json"


# ─── Schémas ──────────────────────────────────────────────────────────────────

@dataclass
class ColumnSchema:
    name: str               # identifiant passerelle  ex: "avancement"
    col_type: str           # KEY | string | float | date | pct | int
    header: str             # header Excel affiché    ex: "% Avancement"
    write: str              # engineer | creation | uo_generique | it_manager | ...
    description: str = ""


@dataclass
class TableSchema:
    id: str                 # ex: "uo.activites"
    source_file_id: str     # ex: "UO-001"
    source_sheet: str       # ex: "Activités"
    table_name: str         # nom du tableau Excel nommé  ex: "TabActivites"
    columns: Dict[str, ColumnSchema] = field(default_factory=dict)
    description: str = ""
    discovered_from: str = ""
    last_seen: str = ""


@dataclass
class VariableSchema:
    id: str                 # ex: "uo.avancement_global"
    var_type: str           # CELL | CELL_NUM | CELL_DATE | CELL_PCT | COMPUTED
    source_file_id: str     # ex: "UO-001"
    formula: str = ""       # pour COMPUTED
    description: str = ""
    discovered_from: str = ""
    last_seen: str = ""


@dataclass
class EcosystemSchema:
    version: str = "1"
    tables: Dict[str, TableSchema] = field(default_factory=dict)
    variables: Dict[str, VariableSchema] = field(default_factory=dict)


# ─── Sérialisation ────────────────────────────────────────────────────────────

def _schema_to_dict(schema: EcosystemSchema) -> dict:
    def _col(c: ColumnSchema) -> dict:
        return asdict(c)

    def _tbl(t: TableSchema) -> dict:
        d = asdict(t)
        d["columns"] = {k: _col(v) for k, v in t.columns.items()}
        return d

    def _var(v: VariableSchema) -> dict:
        return asdict(v)

    return {
        "version": schema.version,
        "tables": {k: _tbl(v) for k, v in schema.tables.items()},
        "variables": {k: _var(v) for k, v in schema.variables.items()},
    }


def _schema_from_dict(d: dict) -> EcosystemSchema:
    tables = {}
    for tid, tdata in d.get("tables", {}).items():
        cols = {
            cname: ColumnSchema(**cdata)
            for cname, cdata in tdata.get("columns", {}).items()
        }
        tdata_copy = {k: v for k, v in tdata.items() if k != "columns"}
        tables[tid] = TableSchema(**tdata_copy, columns=cols)

    variables = {
        vid: VariableSchema(**vdata)
        for vid, vdata in d.get("variables", {}).items()
    }

    return EcosystemSchema(
        version=d.get("version", "1"),
        tables=tables,
        variables=variables,
    )


# ─── Persistence ──────────────────────────────────────────────────────────────

def load() -> EcosystemSchema:
    if not ECOSYSTEM_PATH.exists():
        return EcosystemSchema()
    with open(ECOSYSTEM_PATH, encoding="utf-8") as f:
        return _schema_from_dict(json.load(f))


def save(schema: EcosystemSchema) -> None:
    ECOSYSTEM_PATH.parent.mkdir(parents=True, exist_ok=True)
    with open(ECOSYSTEM_PATH, "w", encoding="utf-8") as f:
        json.dump(_schema_to_dict(schema), f, ensure_ascii=False, indent=2)


# ─── API publique ─────────────────────────────────────────────────────────────

def register_table(table: TableSchema) -> None:
    """Enregistre ou met à jour une table dans l'écosystème."""
    schema = load()
    table.last_seen = str(date.today())
    existing = schema.tables.get(table.id)
    if existing:
        # Fusion : on enrichit les colonnes existantes sans écraser
        for col_name, col in table.columns.items():
            if col_name not in existing.columns:
                existing.columns[col_name] = col
        existing.last_seen = table.last_seen
        if table.source_file_id:
            existing.source_file_id = table.source_file_id
    else:
        schema.tables[table.id] = table
    save(schema)


def register_variable(variable: VariableSchema) -> None:
    """Enregistre ou met à jour une variable dans l'écosystème."""
    schema = load()
    variable.last_seen = str(date.today())
    schema.variables[variable.id] = variable
    save(schema)


def register_many(tables: List[TableSchema], variables: List[VariableSchema]) -> None:
    """Enregistre un lot de tables et variables en une seule opération."""
    schema = load()
    today = str(date.today())

    for table in tables:
        table.last_seen = today
        existing = schema.tables.get(table.id)
        if existing:
            for col_name, col in table.columns.items():
                if col_name not in existing.columns:
                    existing.columns[col_name] = col
            existing.last_seen = today
        else:
            schema.tables[table.id] = table

    for variable in variables:
        variable.last_seen = today
        schema.variables[variable.id] = variable

    save(schema)


def get_table(table_id: str) -> Optional[TableSchema]:
    return load().tables.get(table_id)


def get_variable(variable_id: str) -> Optional[VariableSchema]:
    return load().variables.get(variable_id)


def list_tables() -> List[TableSchema]:
    return list(load().tables.values())


def list_variables() -> List[VariableSchema]:
    return list(load().variables.values())


def summary() -> dict:
    schema = load()
    return {
        "nb_tables": len(schema.tables),
        "nb_variables": len(schema.variables),
        "tables": list(schema.tables.keys()),
        "variables": list(schema.variables.keys()),
    }
