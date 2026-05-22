"""Service partagé de génération MXL depuis une définition de Classe.

Utilisé par :
  - web/schema_app/api/mxl.py        → template MXL N1 (par Classe)
  - web/registry_app/api/xlsx_generator.py → feuille _Manifeste du fichier Excel N2

La génération MXL d'instance (depuis tables.json + hierarchy.json + ecosystem.json)
reste dans web/registry_app/api/mxl.py — source de vérité distincte.
"""
from __future__ import annotations


# ── Helpers ────────────────────────────────────────────────────────────────────

def _var_from_sheet(tbl: dict) -> str:
    """Dérive '$activites' depuis le nom de feuille ('Activites')."""
    sheet = tbl.get("sheet", tbl.get("table_name", "table"))
    return "$" + sheet.lower().replace(" ", "_").replace("-", "_")


def _col_line(full_ref: str, is_key: bool, header: str, write: str) -> str:
    """
    Génère une ligne COL ou chaîne vide si rien à déclarer.
    full_ref : '$activites.id' (déjà construit par l'appelant)
    Syntaxe : COL $table.col : KEY [WRITE=...] [HEADER="..."]
    """
    parts: list[str] = []
    if is_key:
        parts.append("KEY")
    if write:
        parts.append(f"WRITE={write}")
    if header:
        parts.append(f'HEADER="{header}"')
    if not parts:
        return ""
    return f"  COL {full_ref} : {'  '.join(parts)}"


# ── Générateur principal ───────────────────────────────────────────────────────

def build_class_mxl_lines(
    class_id: str,
    file_id: str,
    ft: dict,
    std_tables: list[dict],
    relations: list[dict] | None = None,
) -> list[str]:
    """
    Génère les lignes MXL depuis une définition de Classe. Pure function — aucun I/O.

    Args:
        class_id   : identifiant de la Classe (ex: "uo_instance")
        file_id    : identifiant réel du Post ou placeholder (ex: "UO-001" ou "UO_INSTANCE_ID")
        ft         : entrée de file_types.yaml pour cette Classe
        std_tables : tables __class__{class_id} depuis tables.json
        relations  : relations schema_relations.json (N1 uniquement, None si non disponible)

    Returns:
        liste de lignes MXL — adaptée à parse_sheet (rows 3+) ou à un fichier .mxl texte
    """
    lines: list[str] = []
    push_prefix = ft.get("push_prefix", "")
    base_ns     = push_prefix.rstrip(".").replace("{id}", file_id).replace("{FILE_ID}", file_id)
    style       = ft.get("default_style", "")

    # ── Bloc 0 : En-tête ──────────────────────────────────────────────────────
    lines += [
        "# -- EN-TETE --------------------------------------------------",
        f"FILE_TYPE: {class_id}",
        f"FILE_ID:   {file_id}",
        "VERSION:   1",
    ]
    if style:
        lines.append(f"STYLE:     {style}")
    lines.append("DOC:       ")

    # Champs identitaires (is_key=True) → section IDENT (autodéclaration)
    ident_fields = [f for f in ft.get("min_fields", []) if f.get("is_key")]
    if ident_fields:
        lines.append("# -- IDENTIFICATION -------------------------------------------")
        for f in ident_fields:
            label = f.get("label", f["name"])
            lines.append(f'IDENT {f["name"]} : LABEL="{label}"')
        lines.append("")

    # Champs metadata non-clés (is_key=False, user_input) → en-tête classique
    # Stockés dans manifest_metadata, utilisables dans LIST DYNAMIC WHERE.
    meta_fields = [
        f for f in ft.get("min_fields", [])
        if not f.get("is_key") and f.get("source") in ("user_input", "reference", "parametre")
    ]
    for f in meta_fields:
        label = f.get("label", f["name"])
        lines.append(f"{f['name']}:   # {label}")
    if meta_fields:
        lines.append("")

    # ── Bloc 2 : Imports PULL (suggestions commentées) ────────────────────────
    allowed_ns = ft.get("allowed_namespaces", [])
    if allowed_ns:
        lines.append("# -- IMPORTS ---------------------------------------------------")
        for ns in allowed_ns:
            lines.append(
                f"# PULL {ns}<source> -> FILL_TABLE(<Feuille>, <Table>)  MODE=APPEND_NEW  KEY=id"
            )
        lines.append("")

    # ── Blocs 3+7 : Lecture locale + Colonnes ─────────────────────────────────
    if std_tables:
        lines.append("# -- LECTURE LOCALE --------------------------------------------")
        for tbl in std_tables:
            var   = _var_from_sheet(tbl)
            sheet = tbl.get("sheet", tbl["table_name"])
            lines.append(f"DEF {var} = GET_TABLE({sheet}, {tbl['table_name']})")
        lines.append("")

        lines.append("# -- COLONNES --------------------------------------------------")
        for tbl in std_tables:
            var = _var_from_sheet(tbl)
            for col in tbl.get("columns", []):
                is_key = col.get("is_key", False)
                header = col.get("header") or col.get("name", "")
                write  = col.get("write", "") if not is_key else ""
                if not is_key:
                    write = write or "creation"
                cl = _col_line(f"{var}.{col['name']}", is_key, header, write)
                if cl:
                    lines.append(cl)
        lines.append("")

        # ── Bloc 9 : Exports PUSH (tables) ────────────────────────────────────
        if base_ns:
            lines.append("# -- EXPORTS VERS L'ECOSYSTEME --------------------------------")
            for tbl in std_tables:
                var = _var_from_sheet(tbl)
                key = f"{base_ns}.{tbl['table_name'].lower()}"
                lines.append(f"PUSH {var} -> {key}")
            lines.append("")

    # ── PUSH scalaires pushables (en commentaire) ─────────────────────────────
    pushable = [f for f in ft.get("min_fields", []) if f.get("pushable") and base_ns]
    if pushable:
        lines.append("# -- PUSH SCALAIRES (décommenter après avoir DEF les variables) --")
        for f in pushable:
            lines.append(f"# PUSH ${f['name']} -> {base_ns}.{f['name']}")
        lines.append("")

    # ── Flux N1 (Tissage schéma) — commentaires uniquement ────────────────────
    if relations:
        flux_sortants = [
            r for r in relations
            if r.get("parent_class") == class_id and r.get("flux")
        ]
        if flux_sortants:
            lines.append("# ── Flux N1 (patrons définis en Schéma N1) ───────────────────")
            for rel in flux_sortants:
                child     = rel["child_class"]
                qualifier = rel.get("qualifier", "TYPICAL")
                lines.append(f"# Relation {qualifier} : {class_id} ↔ {child}")
                for fx in rel.get("flux", []):
                    table  = fx.get("table", "")
                    mode   = fx.get("mode", "PULL")
                    source = fx.get("source", "child")
                    if source == "child":
                        if mode == "COLLECT":
                            lines.append(
                                f"# COLLECT {child}.*.{table}"
                                f" -> FILL_TABLE({{sheet}}, {table}) MODE=APPEND"
                            )
                        else:
                            lines.append(
                                f"# PULL {child}.{{source_id}}.{table}"
                                f" -> FILL_TABLE({{sheet}}, {table}) MODE=OVERWRITE"
                            )
                    else:
                        lines.append(f"# PUSH {table} -> {child}.{{target_id}}.{table}")
                lines.append("#")

    return lines
