"""
ExoSync — Ecosystem Schema Manager (Exomap v2)
===============================================
Catalogue vivant de l'exostructure : fichiers, tables, variables, dépendances.

L'Exomap est le reflet de l'exostructure — jamais sa définition.
Elle est construite dynamiquement à partir des Manifestes lus lors des syncs.

Persisté dans output/ecosystem.json.

Contenu :
  files     — registre des fichiers connus (path, type, dernier sync)
  tables    — tables Excel découvertes via GET_TABLE
  variables — variables COMPUTE/CELL découvertes
  edges     — arcs PULL/PUSH entre fichiers et store (graphe de dépendances)
"""
import json
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

ECOSYSTEM_PATH = Path(__file__).parent.parent / "output" / "ecosystem.json"


# ─── Schémas de données ───────────────────────────────────────────────────────

@dataclass
class ColumnSchema:
    name: str               # identifiant passerelle  ex: "avancement"
    col_type: str           # KEY | string | float | date | pct | int
    header: str             # header Excel affiché    ex: "% Avancement"
    write: str              # engineer | creation | uo_generique | it_manager
    description: str = ""


@dataclass
class TableSchema:
    id: str                 # ex: "UO-001::TabActivites"
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
    var_type: str           # CELL | CELL_NUM | COMPUTED | ...
    source_file_id: str     # ex: "UO-001"
    formula: str = ""       # pour COMPUTED
    description: str = ""
    discovered_from: str = ""
    last_seen: str = ""


@dataclass
class FileRecord:
    """Représente un fichier Excel connu dans l'écosystème."""
    file_id: str            # ex: "UO-001"
    path: str               # chemin relatif  ex: "UOs/UO-001_climat.xlsx"
    file_type: str          # ex: "uo_instance"
    manifest_version: str = "1"
    last_sync: Optional[str] = None   # ISO datetime
    status: str = "unknown"           # "ok" | "error" | "unknown"
    manifest_metadata: Dict[str, str] = field(default_factory=dict)  # champs en-tête libre (owner, projet…)


@dataclass
class EdgeRecord:
    """
    Arc dirigé dans le graphe de dépendances.

    Exemples :
      PULL  from="store::projet.acteurs"       to="UO-001::TabActeurs"
      PUSH  from="UO-001::avancement_global"   to="store::uo.UO-001.avancement"
    """
    edge_type: str   # "PULL" | "PUSH"
    from_node: str   # "store::<clé>" ou "<file_id>::<var>"
    to_node: str     # "<file_id>::<table>" ou "store::<clé>"
    mode: str = ""   # "OVERWRITE" | "APPEND_NEW" | "UPDATE" | "READ_ONLY"


@dataclass
class ConsistencyWarning:
    code: str        # "PULL_NEVER_PUSHED" | "PUSH_CONFLICT" | "STALE_FILE"
                     # "COLLECT_FILE_NOT_FOUND" | "COLLECT_TABLE_MISSING"
                     # "COLLECT_CIRCULAR_DEP" | "COLLECT_TYPE_CAST"
    message: str
    details: str = ""


@dataclass
class ListRecord:
    """
    Déclaration d'une liste de fichiers fils dans le Manifeste d'un père.

    Exemples :
      LIST UOs_actifs FROM TABLE liste_uo          → form="TABLE"
      LIST uo_mi20 TYPE=uo_instance WHERE projet=MI20  → form="DYNAMIC"
    """
    list_name: str          # ex: "UOs_actifs"
    owner_file_id: str      # fichier père qui déclare la liste
    form: str               # "TABLE" | "DYNAMIC"
    # Forme TABLE
    source_table: str = ""          # nom de la table Excel dans le père
    context_columns: List[str] = field(default_factory=list)  # colonnes contextuelles détectées
    # Forme DYNAMIC
    filter_type: str = ""           # "uo_instance"
    filter_where: List[str] = field(default_factory=list)     # ["projet=MI20", "owner=Jean"]


@dataclass
class CollectEdge:
    """
    Arc de collecte cross-fichiers : père agrège une table depuis ses fils.

    Exemple :
      COLLECT Planning FROM UOs_actifs INTO vue_planning
      → owner_file_id="synthese_mi20", list_name="UOs_actifs",
        source_table="Planning", target_table="vue_planning"
    """
    owner_file_id: str      # fichier père
    list_name: str          # "UOs_actifs"
    source_table: str       # table extraite chez chaque fils  ex: "Planning"
    target_table: str       # table résultat dans le père       ex: "vue_planning"
    context_columns: List[str] = field(default_factory=list)  # colonnes injectées depuis la liste
    where_clause: str = ""  # filtre sur les lignes  ex: "criticite >= 3"
    cols_filter: List[str] = field(default_factory=list)      # sélection de colonnes
    with_fields: List[str] = field(default_factory=list)      # champs WITH (liste dynamique)


@dataclass
class EcosystemSchema:
    version: str = "2.0"
    last_scan: Optional[str] = None
    files: Dict[str, FileRecord] = field(default_factory=dict)
    tables: Dict[str, TableSchema] = field(default_factory=dict)
    variables: Dict[str, VariableSchema] = field(default_factory=dict)
    edges: List[EdgeRecord] = field(default_factory=list)
    lists: List[ListRecord] = field(default_factory=list)
    collect_edges: List[CollectEdge] = field(default_factory=list)


# ─── Sérialisation ────────────────────────────────────────────────────────────

def _to_dict(schema: EcosystemSchema) -> dict:
    return {
        "version":       schema.version,
        "last_scan":     schema.last_scan,
        "files":         {k: asdict(v) for k, v in schema.files.items()},
        "tables":        {
            k: {**asdict(v), "columns": {cn: asdict(cv) for cn, cv in v.columns.items()}}
            for k, v in schema.tables.items()
        },
        "variables":     {k: asdict(v) for k, v in schema.variables.items()},
        "edges":         [asdict(e) for e in schema.edges],
        "lists":         [asdict(l) for l in schema.lists],
        "collect_edges": [asdict(c) for c in schema.collect_edges],
    }


def _from_dict(d: dict) -> EcosystemSchema:
    files = {
        fid: FileRecord(**fdata)
        for fid, fdata in d.get("files", {}).items()
    }
    tables = {}
    for tid, tdata in d.get("tables", {}).items():
        cols = {
            cname: ColumnSchema(**cdata)
            for cname, cdata in tdata.get("columns", {}).items()
        }
        base = {k: v for k, v in tdata.items() if k != "columns"}
        tables[tid] = TableSchema(**base, columns=cols)

    variables = {
        vid: VariableSchema(**vdata)
        for vid, vdata in d.get("variables", {}).items()
    }
    edges = [EdgeRecord(**e) for e in d.get("edges", [])]
    lists = [ListRecord(**l) for l in d.get("lists", [])]
    collect_edges = [CollectEdge(**c) for c in d.get("collect_edges", [])]

    # Migration : FileRecord sans manifest_metadata (ecosystem.json antérieur)
    for frec in files.values():
        if not hasattr(frec, "manifest_metadata"):
            frec.manifest_metadata = {}

    return EcosystemSchema(
        version=d.get("version", "2.0"),
        last_scan=d.get("last_scan"),
        files=files,
        tables=tables,
        variables=variables,
        edges=edges,
        lists=lists,
        collect_edges=collect_edges,
    )


# ─── Persistence ──────────────────────────────────────────────────────────────

def load(path: Path = ECOSYSTEM_PATH) -> EcosystemSchema:
    if not path.exists():
        return EcosystemSchema()
    with open(path, encoding="utf-8") as f:
        raw = json.load(f)
    # Migration v1 → v2
    if raw.get("version", "1") == "1":
        raw["version"] = "2.0"
        raw.setdefault("files", {})
        raw.setdefault("edges", [])
        raw.setdefault("last_scan", None)
    return _from_dict(raw)


def save(schema: EcosystemSchema, path: Path = ECOSYSTEM_PATH) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    schema.last_scan = datetime.now().isoformat(timespec="seconds")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(_to_dict(schema), f, ensure_ascii=False, indent=2)


# ─── API publique — enregistrement ───────────────────────────────────────────

def register_table(table: TableSchema, path: Path = ECOSYSTEM_PATH) -> None:
    """Enregistre ou met à jour une table dans l'écosystème."""
    schema = load(path)
    today = datetime.now().isoformat(timespec="seconds")
    table.last_seen = today
    existing = schema.tables.get(table.id)
    if existing:
        for col_name, col in table.columns.items():
            if col_name not in existing.columns:
                existing.columns[col_name] = col
        existing.last_seen = today
        if table.source_file_id:
            existing.source_file_id = table.source_file_id
    else:
        schema.tables[table.id] = table
    save(schema, path)


def register_variable(variable: VariableSchema, path: Path = ECOSYSTEM_PATH) -> None:
    """Enregistre ou met à jour une variable dans l'écosystème."""
    schema = load(path)
    variable.last_seen = datetime.now().isoformat(timespec="seconds")
    schema.variables[variable.id] = variable
    save(schema, path)


def register_many(
    tables: List[TableSchema],
    variables: List[VariableSchema],
    path: Path = ECOSYSTEM_PATH,
) -> None:
    """Enregistre un lot de tables et variables en une seule opération."""
    schema = load(path)
    today = datetime.now().isoformat(timespec="seconds")

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

    save(schema, path)


def record_file_sync(
    file_id: str,
    path_str: str,
    file_type: str,
    status: str,
    manifest_version: str = "1",
    manifest_metadata: Optional[Dict[str, str]] = None,
    ecosystem_path: Path = ECOSYSTEM_PATH,
) -> None:
    """
    Enregistre (ou met à jour) la trace d'un sync de fichier dans l'Exomap.
    Appelé par sync.py après chaque execute_ast().
    manifest_metadata : champs libres de l'en-tête (owner, projet, has_risques…)
    """
    schema = load(ecosystem_path)
    schema.files[file_id] = FileRecord(
        file_id=file_id,
        path=path_str,
        file_type=file_type,
        manifest_version=manifest_version,
        last_sync=datetime.now().isoformat(timespec="seconds"),
        status=status,
        manifest_metadata=manifest_metadata or {},
    )
    save(schema, ecosystem_path)


def record_edges_from_ast(
    ast,           # PasserelleAST — import circulaire évité
    file_id: str,
    ecosystem_path: Path = ECOSYSTEM_PATH,
) -> None:
    """
    Extrait les arcs PULL/PUSH du Manifeste et les ajoute à l'Exomap.
    Remplace les anciens arcs du même fichier (re-sync = re-découverte).
    Appelé par sync.py après execute_ast().
    """
    schema = load(ecosystem_path)

    # Retirer les arcs existants de ce fichier
    schema.edges = [
        e for e in schema.edges
        if not (e.from_node.startswith(f"{file_id}::") or
                e.to_node.startswith(f"{file_id}::"))
    ]

    # Arcs PULL : store → fichier
    for pull in ast.pulls:
        schema.edges.append(EdgeRecord(
            edge_type="PULL",
            from_node=f"store::{pull.global_name}",
            to_node=f"{file_id}::{pull.table}",
            mode=pull.mode,
        ))

    # Arcs PUSH : fichier → store
    for push in ast.pushes:
        schema.edges.append(EdgeRecord(
            edge_type="PUSH",
            from_node=f"{file_id}::{push.var_name}",
            to_node=f"store::{push.global_name}",
        ))

    # Listes déclarées (LIST)
    schema.lists = [l for l in schema.lists if l.owner_file_id != file_id]
    for list_node in getattr(ast, "lists", []):
        schema.lists.append(ListRecord(
            list_name=list_node.name,
            owner_file_id=file_id,
            form=list_node.form,
            source_table=getattr(list_node, "source_table", ""),
            filter_type=getattr(list_node, "filter_type", ""),
            filter_where=[f"{f}={v}" for f, _, v in getattr(list_node, "filter_where", [])],
        ))

    # Arcs COLLECT : père → fils (via liste)
    schema.collect_edges = [c for c in schema.collect_edges if c.owner_file_id != file_id]
    for collect in getattr(ast, "collects", []):
        schema.collect_edges.append(CollectEdge(
            owner_file_id=file_id,
            list_name=collect.list_name,
            source_table=collect.source_table,
            target_table=collect.target_table,
            where_clause=getattr(collect, "where_clause", ""),
            cols_filter=getattr(collect, "cols_filter", []),
            with_fields=getattr(collect, "with_fields", []),
        ))

    save(schema, ecosystem_path)


# ─── API publique — requêtes hiérarchie ──────────────────────────────────────

def get_files_by_type(
    file_type: str,
    filters: Optional[Dict[str, str]] = None,
    path: Path = ECOSYSTEM_PATH,
) -> List[FileRecord]:
    """
    Retourne les fichiers connus d'un type donné, avec filtres optionnels
    sur manifest_metadata. Utilisé par resolve_lists() pour les listes DYNAMIC.

    Exemple :
      get_files_by_type("uo_instance", {"projet": "MI20", "owner": "Jean"})
    """
    schema = load(path)
    results = [f for f in schema.files.values() if f.file_type == file_type]
    if filters:
        for key, value in filters.items():
            results = [f for f in results if f.manifest_metadata.get(key) == value]
    return results


def get_collect_children(
    owner_file_id: str,
    path: Path = ECOSYSTEM_PATH,
) -> List[CollectEdge]:
    """Retourne les arcs COLLECT dont ce fichier est le père."""
    schema = load(path)
    return [c for c in schema.collect_edges if c.owner_file_id == owner_file_id]


def get_collect_parents(
    child_file_id: str,
    path: Path = ECOSYSTEM_PATH,
) -> List[str]:
    """
    Retourne les FILE_IDs des pères qui collectent depuis ce fichier
    (via une liste qui l'inclut).
    """
    schema = load(path)
    parents = []
    for lst in schema.lists:
        # Seules les listes TABLE peuvent être vérifiées statiquement
        # (les listes DYNAMIC sont résolues à l'exécution)
        if lst.form == "DYNAMIC":
            continue
        # Vérifier si child_file_id est dans la table du père
        # (non vérifiable ici sans ouvrir l'Excel → on retourne les pères candidats)
    return parents  # complété à l'exécution par l'executor


# ─── Détection d'incohérences ─────────────────────────────────────────────────

def check_consistency(path: Path = ECOSYSTEM_PATH) -> List[ConsistencyWarning]:
    """
    Analyse l'Exomap et retourne la liste des incohérences détectées.

    Règles vérifiées :
      1. PULL_NEVER_PUSHED  — une clé store est consommée mais jamais produite
      2. PUSH_CONFLICT      — deux fichiers pushent vers la même clé store
      3. STALE_FILE         — fichier connu mais jamais synchronisé
    """
    schema = load(path)
    warnings: List[ConsistencyWarning] = []

    pushed_keys  = {e.to_node for e in schema.edges if e.edge_type == "PUSH"}
    pulled_keys  = {e.from_node for e in schema.edges if e.edge_type == "PULL"}

    # 1. PULL de clés jamais PUSHées
    for key in pulled_keys:
        if key not in pushed_keys:
            warnings.append(ConsistencyWarning(
                code="PULL_NEVER_PUSHED",
                message=f"{key} est consommée (PULL) mais jamais produite (PUSH)",
                details=f"Clé store '{key.replace('store::', '')}' introuvable dans les PUSH",
            ))

    # 2. Conflits PUSH — plusieurs fichiers pushent vers la même clé
    push_edges = [e for e in schema.edges if e.edge_type == "PUSH"]
    target_count: Dict[str, List[str]] = {}
    for e in push_edges:
        target_count.setdefault(e.to_node, []).append(e.from_node)
    for target, sources in target_count.items():
        if len(sources) > 1:
            warnings.append(ConsistencyWarning(
                code="PUSH_CONFLICT",
                message=f"Conflit : {target} est pushé par {len(sources)} fichiers",
                details=f"Sources : {', '.join(sources)}",
            ))

    # 3. Fichiers jamais synchronisés
    for fid, frec in schema.files.items():
        if frec.last_sync is None:
            warnings.append(ConsistencyWarning(
                code="STALE_FILE",
                message=f"{fid} ({frec.file_type}) n'a jamais été synchronisé",
                details=f"Chemin : {frec.path}",
            ))

    # 4. COLLECT référençant une liste inconnue
    known_lists = {(l.owner_file_id, l.list_name) for l in schema.lists}
    for ce in schema.collect_edges:
        if (ce.owner_file_id, ce.list_name) not in known_lists:
            warnings.append(ConsistencyWarning(
                code="COLLECT_LIST_UNKNOWN",
                message=f"{ce.owner_file_id} : COLLECT référence la liste '{ce.list_name}' non déclarée",
                details=f"Cible : {ce.target_table}",
            ))

    # 5. COLLECT_CIRCULAR_DEP — détection de cycles dans le graphe père/fils
    # Graphe statique : père → [pères des fils connus via listes TABLE]
    # (les listes DYNAMIC ne sont pas vérifiables statiquement)
    parent_map: Dict[str, str] = {}  # list_name+owner → owner (pour cycles simples)
    for ce in schema.collect_edges:
        for lst in schema.lists:
            if lst.list_name == ce.list_name and lst.owner_file_id == ce.owner_file_id:
                if lst.form == "TABLE":
                    # Vérifier si le père est référencé dans ses propres fils
                    # (nécessite d'ouvrir les Excel → non fait ici, check superficiel)
                    if ce.owner_file_id == ce.source_table:
                        warnings.append(ConsistencyWarning(
                            code="COLLECT_CIRCULAR_DEP",
                            message=f"Cycle possible : {ce.owner_file_id} se référence lui-même",
                            details=f"Liste '{ce.list_name}', COLLECT {ce.source_table} → {ce.target_table}",
                        ))

    return warnings


# ─── Lineage — représentation textuelle du graphe ────────────────────────────

def lineage_text(
    file_id: Optional[str] = None,
    path: Path = ECOSYSTEM_PATH,
) -> str:
    """
    Retourne une représentation textuelle du graphe de dépendances.

    Si file_id est fourni, n'affiche que les arcs de ce fichier.
    """
    schema = load(path)
    edges = schema.edges

    if file_id:
        edges = [
            e for e in edges
            if e.from_node.startswith(f"{file_id}::") or
               e.to_node.startswith(f"{file_id}::")
        ]

    if not edges and not schema.files:
        return "  (Exomap vide — lancez au moins un sync)"

    lines = []

    # Grouper les fichiers connus
    files_with_edges = set()
    for e in schema.edges:
        for part in (e.from_node, e.to_node):
            fid = part.split("::")[0]
            if fid != "store":
                files_with_edges.add(fid)

    # Afficher les PUSH par fichier source
    pushes_by_file: Dict[str, List[EdgeRecord]] = {}
    pulls_by_file:  Dict[str, List[EdgeRecord]] = {}
    for e in edges:
        if e.edge_type == "PUSH":
            fid = e.from_node.split("::")[0]
            pushes_by_file.setdefault(fid, []).append(e)
        else:
            fid = e.to_node.split("::")[0]
            pulls_by_file.setdefault(fid, []).append(e)

    all_fids = sorted(
        files_with_edges | set(schema.files.keys()),
        key=lambda f: (schema.files.get(f, FileRecord(f, "", "")).file_type, f)
    )

    if file_id:
        all_fids = [f for f in all_fids if f == file_id]

    for fid in all_fids:
        frec = schema.files.get(fid)
        ftype = frec.file_type if frec else "?"
        status = frec.status if frec else "unknown"
        status_tag = "[OK]" if status == "ok" else f"[{status.upper()}]"
        lines.append(f"\n{fid} ({ftype}) {status_tag}")

        for e in pulls_by_file.get(fid, []):
            store_key = e.from_node.replace("store::", "")
            table = e.to_node.split("::")[-1]
            mode = f"  MODE={e.mode}" if e.mode else ""
            lines.append(f"  <-- PULL  store::{store_key} -> {table}{mode}")

        for e in pushes_by_file.get(fid, []):
            store_key = e.to_node.replace("store::", "")
            var = e.from_node.split("::")[-1]
            lines.append(f"  --> PUSH  {var} -> store::{store_key}")

        for ce in schema.collect_edges:
            if ce.owner_file_id == fid:
                where = f" WHERE {ce.where_clause}" if ce.where_clause else ""
                lines.append(
                    f"  <<< COLLECT  [{ce.list_name}].{ce.source_table}"
                    f" -> {ce.target_table}{where}"
                )

        for lst in schema.lists:
            if lst.owner_file_id == fid:
                if lst.form == "TABLE":
                    lines.append(f"  --- LIST  {lst.list_name} FROM TABLE {lst.source_table}")
                else:
                    where = " WHERE " + " AND ".join(lst.filter_where) if lst.filter_where else ""
                    lines.append(f"  --- LIST  {lst.list_name} TYPE={lst.filter_type}{where}")

    return "\n".join(lines) if lines else "  (aucun arc enregistré)"


def lineage_dict(path: Path = ECOSYSTEM_PATH) -> dict:
    """Retourne l'Exomap complète sous forme de dict (pour JSON export ou CLI)."""
    schema = load(path)
    return {
        "version":       schema.version,
        "last_scan":     schema.last_scan,
        "files":         {k: asdict(v) for k, v in schema.files.items()},
        "edges":         [asdict(e) for e in schema.edges],
        "lists":         [asdict(l) for l in schema.lists],
        "collect_edges": [asdict(c) for c in schema.collect_edges],
        "stats": {
            "nb_files":          len(schema.files),
            "nb_tables":         len(schema.tables),
            "nb_variables":      len(schema.variables),
            "nb_edges":          len(schema.edges),
            "nb_pull_edges":     sum(1 for e in schema.edges if e.edge_type == "PULL"),
            "nb_push_edges":     sum(1 for e in schema.edges if e.edge_type == "PUSH"),
            "nb_lists":          len(schema.lists),
            "nb_collect_edges":  len(schema.collect_edges),
        },
    }


# ─── API publique historique (rétro-compat v1) ───────────────────────────────

def get_table(table_id: str) -> Optional[TableSchema]:
    return load().tables.get(table_id)


def get_variable(variable_id: str) -> Optional[VariableSchema]:
    return load().variables.get(variable_id)


def list_tables() -> List[TableSchema]:
    return list(load().tables.values())


def list_variables() -> List[VariableSchema]:
    return list(load().variables.values())


def summary(ecosystem_path: Optional[Path] = None) -> dict:
    schema = load(ecosystem_path or ECOSYSTEM_PATH)
    return {
        "nb_tables":    len(schema.tables),
        "nb_variables": len(schema.variables),
        "nb_files":     len(schema.files),
        "nb_edges":     len(schema.edges),
        "tables":       list(schema.tables.keys()),
        "variables":    list(schema.variables.keys()),
    }
