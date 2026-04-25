"""
ExoSync CLI
===========
Interface en ligne de commande pour piloter ExoSync.

Usage :
    python -m src [commande] [options]

Commandes disponibles :
    sync      Synchronise un ou plusieurs fichiers Excel
    status    Affiche le contenu du store central
    store     Opérations bas-niveau sur le store (get / set / clear)
    generate  Génère les fichiers Excel depuis le registre (uo_instances.json)
"""
import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).parent.parent


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _ok(msg: str) -> None:
    print(f"  [OK] {msg}")


def _warn(msg: str) -> None:
    print(f"  [!]  {msg}")


def _err(msg: str) -> None:
    print(f"  [ERR] {msg}", file=sys.stderr)


def _header(title: str) -> None:
    width = 60
    print(f"\n{'-' * width}")
    print(f"  {title}")
    print(f"{'-' * width}")


# ─── Commande : sync ──────────────────────────────────────────────────────────

def cmd_sync(args: argparse.Namespace) -> int:
    """Lance la synchronisation via sync.synchroniser()."""
    from src.sync import synchroniser

    ids   = args.id   or None
    types = args.type or None
    force = args.force

    _header("ExoSync — Synchronisation")
    if ids:
        print(f"  Fichiers : {', '.join(ids)}")
    elif types:
        print(f"  Types    : {', '.join(types)}")
    else:
        print("  Périmètre : tous les fichiers du registre")

    try:
        rapport_path = synchroniser(ids=ids, types=types, force=force)
        _ok(f"Rapport sauvegardé : {rapport_path.relative_to(ROOT)}")
        return 0
    except Exception as exc:
        _err(f"Erreur inattendue : {exc}")
        return 1


# ─── Commande : status ────────────────────────────────────────────────────────

def cmd_status(args: argparse.Namespace) -> int:
    """Affiche le contenu du store central."""
    from src.store import JsonStore, DEFAULT_STORE_PATH

    store = JsonStore(DEFAULT_STORE_PATH)
    snap  = store.snapshot()
    variables = snap.get("variables", {})

    prefix = args.prefix or ""
    if prefix:
        variables = {k: v for k, v in variables.items() if k.startswith(prefix)}

    _header(f"Store central — {DEFAULT_STORE_PATH.relative_to(ROOT)}")
    print(f"  Dernière MAJ : {snap.get('derniere_maj', 'jamais')}")
    print(f"  Variables    : {len(variables)}")
    print()

    if not variables:
        print("  (store vide)")
        return 0

    # Grouper par préfixe (ex: "uo.UO-001.*", "projet.*")
    max_key = max(len(k) for k in variables)
    for key, val in sorted(variables.items()):
        if isinstance(val, list):
            display = f"[table: {len(val)} lignes]"
        elif isinstance(val, dict):
            display = f"{{dict: {len(val)} clés}}"
        else:
            display = str(val)
        print(f"  {key:<{max_key}}  =  {display}")

    return 0


# ─── Commande : store ─────────────────────────────────────────────────────────

def cmd_store_get(args: argparse.Namespace) -> int:
    from src.store import JsonStore, DEFAULT_STORE_PATH
    store = JsonStore(DEFAULT_STORE_PATH)
    val = store.get(args.key)
    if val is None:
        _warn(f"Clé '{args.key}' absente du store")
        return 1
    print(json.dumps(val, ensure_ascii=False, indent=2, default=str))
    return 0


def cmd_store_set(args: argparse.Namespace) -> int:
    from src.store import JsonStore, DEFAULT_STORE_PATH
    store = JsonStore(DEFAULT_STORE_PATH)
    # Tenter de parser la valeur comme JSON, sinon conserver comme str
    try:
        val = json.loads(args.value)
    except json.JSONDecodeError:
        val = args.value
    store.set(args.key, val)
    _ok(f"{args.key} = {val!r}")
    return 0


def cmd_store_delete(args: argparse.Namespace) -> int:
    from src.store import JsonStore, DEFAULT_STORE_PATH
    store = JsonStore(DEFAULT_STORE_PATH)
    store.delete(args.key)
    _ok(f"Clé '{args.key}' supprimée (ou absente)")
    return 0


def cmd_store_clear(args: argparse.Namespace) -> int:
    from src.store import JsonStore, DEFAULT_STORE_PATH
    store = JsonStore(DEFAULT_STORE_PATH)
    n = len(store.get_all())
    store.clear()
    _ok(f"Store vidé ({n} variable(s) supprimée(s))")
    return 0


# ─── Commande : lineage ───────────────────────────────────────────────────────

def cmd_lineage(args: argparse.Namespace) -> int:
    """Affiche le graphe de dépendances (Exomap) et les incohérences."""
    from src.ecosystem import lineage_text, lineage_dict, check_consistency
    from src.config_loader import load_registre, load_acteurs, load_file_types
    import json as _json

    file_id = getattr(args, "id", None)

    if getattr(args, "json", False):
        print(_json.dumps(lineage_dict(), ensure_ascii=False, indent=2))
        return 0

    _header(f"ExoSync — Lineage (Exomap v2)")

    text = lineage_text(file_id=file_id)
    print(text)

    # Vérification de cohérence
    warnings = check_consistency()
    if warnings:
        print(f"\n  --- {len(warnings)} avertissement(s) ---")
        for w in warnings:
            _warn(f"[{w.code}] {w.message}")
            if w.details:
                print(f"       {w.details}")
    else:
        print("\n  Aucune incoherence detectee.")

    # Ownership summary
    try:
        entrees = load_registre()
        acteurs_idx = {a.id: a for a in load_acteurs()}
        file_types = load_file_types()
        # Filtrer si --id
        if file_id:
            entrees = [e for e in entrees if e.id == file_id]
        if entrees:
            print(f"\n  --- Owners ---")
            for e in entrees:
                owner_str = "<non assigné>"
                if e.owner_id:
                    a = acteurs_idx.get(e.owner_id)
                    if a:
                        owner_str = f"{a.nom} ({a.role.value})  [{e.owner_id}]"
                    else:
                        owner_str = f"<inconnu: {e.owner_id}>"
                expected = file_types.get(e.type_fichier, {}).get("owner_role", "—")
                print(f"    {e.id:25s}  owner={owner_str}  role_attendu={expected}")
    except Exception:
        pass  # pas bloquant

    return 0


# ─── Commande : generate ──────────────────────────────────────────────────────

def cmd_generate(args: argparse.Namespace) -> int:
    """Génère les fichiers Excel des UO depuis uo_instances.json."""
    from src.config_loader import load_uo_instances
    from src.generators.uo_generator import generate_uo_file, OUTPUT_DIR

    _header("ExoSync — Génération des fichiers UO")

    instances = load_uo_instances()
    if args.id:
        instances = [uo for uo in instances if uo.id in args.id]
        if not instances:
            _err(f"Aucune UO trouvée pour : {args.id}")
            return 1

    output_dir = Path(args.output) if args.output else OUTPUT_DIR

    generated = []
    errors     = []
    for uo in instances:
        try:
            path = generate_uo_file(uo, output_dir=output_dir)
            _ok(f"{uo.id} → {path.relative_to(ROOT)}")
            generated.append(path)
        except Exception as exc:
            _err(f"{uo.id} : {exc}")
            errors.append(uo.id)

    print()
    print(f"  {len(generated)} fichier(s) genere(s), {len(errors)} erreur(s)")
    return 0 if not errors else 1


# ─── Commande : doctor ────────────────────────────────────────────────────────

def cmd_doctor(args: argparse.Namespace) -> int:
    """Diagnostique la sante de l'ecosysteme ExoSync."""
    from src.store import DEFAULT_STORE_PATH, JsonStore
    from src.ecosystem import ECOSYSTEM_PATH, load as eco_load, check_consistency
    from src.config_loader import load_registre, validate_owner_roles
    from src.history import list_runs, list_snapshots
    from pathlib import Path

    _header("ExoSync — Doctor (diagnostic)")
    ok_count = 0
    warn_count = 0
    err_count = 0

    # ── Store ────────────────────────────────────────────────────────────────
    print("\n  [Store]")
    if DEFAULT_STORE_PATH.exists():
        store = JsonStore(DEFAULT_STORE_PATH)
        nb = len(store.get_all())
        _ok(f"store.json present ({nb} variables)")
        ok_count += 1
    else:
        _warn("store.json absent (aucune sync effectuee ?)")
        warn_count += 1

    # ── Ecosystem ────────────────────────────────────────────────────────────
    print("\n  [Ecosysteme]")
    if ECOSYSTEM_PATH.exists():
        schema = eco_load()
        _ok(f"ecosystem.json present ({len(schema.files)} fichiers, {len(schema.edges)} arcs)")
        ok_count += 1
        warnings = check_consistency()
        if warnings:
            for w in warnings:
                _warn(f"[{w.code}] {w.message}")
                warn_count += 1
        else:
            _ok("Aucune incoherence detectee")
            ok_count += 1
    else:
        _warn("ecosystem.json absent")
        warn_count += 1

    # ── Registre ─────────────────────────────────────────────────────────────
    print("\n  [Registre]")
    entrees = []
    try:
        entrees = load_registre()
        _ok(f"registre.json : {len(entrees)} entree(s)")
        ok_count += 1
        for e in entrees:
            chemin = ROOT / e.chemin
            if not chemin.exists():
                _err(f"Fichier manquant : {e.chemin} ({e.id})")
                err_count += 1
            elif e.statut_dernier_synchro == "erreur":
                _warn(f"{e.id} : derniere sync en erreur")
                warn_count += 1
    except Exception as exc:
        _err(f"Impossible de lire le registre : {exc}")
        err_count += 1

    # ── Ownership ────────────────────────────────────────────────────────────
    print("\n  [Ownership]")
    try:
        violations = validate_owner_roles(entrees if entrees else None)
        if not violations:
            _ok("Tous les owners ont le bon role")
            ok_count += 1
        else:
            for v in violations:
                _warn(str(v))
                warn_count += 1
    except Exception as exc:
        _warn(f"Validation ownership impossible : {exc}")
        warn_count += 1

    # ── Historique ───────────────────────────────────────────────────────────
    print("\n  [Historique]")
    runs = list_runs()
    snaps = list_snapshots()
    _ok(f"{len(runs)} run(s) en historique, {len(snaps)} snapshot(s)")
    ok_count += 1

    # ── Résumé ───────────────────────────────────────────────────────────────
    print(f"\n  {'─'*40}")
    print(f"  OK={ok_count}  WARN={warn_count}  ERR={err_count}")

    if err_count > 0:
        return 2
    if warn_count > 0:
        return 1
    return 0


# ─── Commande : history ───────────────────────────────────────────────────────

def cmd_history(args: argparse.Namespace) -> int:
    """Affiche l'historique des runs et snapshots."""
    from src.history import list_runs, list_snapshots, compare_snapshots, history_of_key
    import json as _json

    if args.key:
        # Historique d'une cle specifique
        _header(f"Historique de la cle : {args.key}")
        vals = history_of_key(args.key)
        if not vals:
            print("  Aucun snapshot trouve.")
            return 0
        for ts, v in vals:
            print(f"  {ts}  =  {v!r}")
        return 0

    if args.compare:
        # Comparer deux snapshots
        snaps = list_snapshots()
        if len(snaps) < 2:
            _err("Il faut au moins 2 snapshots pour comparer.")
            return 1
        a, b = snaps[1], snaps[0]
        _header(f"Comparaison snapshots")
        print(f"  Avant  : {a.name}")
        print(f"  Apres  : {b.name}")
        diff = compare_snapshots(a, b)
        if diff["ajouts"]:
            print(f"\n  Ajouts ({len(diff['ajouts'])}) :")
            for k, v in diff["ajouts"].items():
                print(f"    + {k} = {v!r}")
        if diff["suppressions"]:
            print(f"\n  Suppressions ({len(diff['suppressions'])}) :")
            for k, v in diff["suppressions"].items():
                print(f"    - {k} = {v!r}")
        if diff["modifications"]:
            print(f"\n  Modifications ({len(diff['modifications'])}) :")
            for k, d in diff["modifications"].items():
                print(f"    ~ {k} : {d['avant']!r} -> {d['apres']!r}")
        print(f"\n  Inchanges : {diff['inchanges']}")
        return 0

    # Liste des runs
    _header("ExoSync — Historique des runs")
    runs = list_runs()
    snaps = list_snapshots()
    print(f"  Runs      : {len(runs)}")
    print(f"  Snapshots : {len(snaps)}")

    n = min(args.last or 10, len(runs))
    if runs:
        print(f"\n  Derniers {n} run(s) :")
        for r in runs[:n]:
            print(f"    {r.name}")

    return 0


# ─── Commande : doc ──────────────────────────────────────────────────────────

def cmd_doc(args: argparse.Namespace) -> int:
    """Genere la documentation HTML de l ecosysteme."""
    from src.doc_generator import generate_html_doc

    _header("ExoSync — Generation documentation HTML")
    output_dir = Path(args.output) if args.output else None

    try:
        out = generate_html_doc(output_dir=output_dir)
        try:
            display = out.relative_to(ROOT)
        except ValueError:
            display = out
        _ok(f"Documentation generee : {display}")
        return 0
    except Exception as exc:
        _err(f"Erreur : {exc}")
        return 1


# ─── Parseur argparse ─────────────────────────────────────────────────────────

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="python -m src",
        description="ExoSync — synchronisation de données via fichiers Excel MXL",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
exemples :
  python -m src sync                        # synchronise tout
  python -m src sync --id UO-001 UO-002    # synchronise 2 fichiers
  python -m src sync --type uo_instance    # synchronise par type
  python -m src status                     # affiche le store
  python -m src status --prefix uo.        # filtre par préfixe
  python -m src store get uo.UO-001.avancement
  python -m src store set uo.UO-001.avancement 75.0
  python -m src store clear
  python -m src generate                   # génère tous les fichiers UO
  python -m src generate --id UO-001      # génère un seul fichier
""",
    )

    sub = parser.add_subparsers(dest="command", metavar="commande")
    sub.required = True

    # ── sync ──────────────────────────────────────────────────────────────────
    p_sync = sub.add_parser("sync", help="Synchronise les fichiers Excel")
    p_sync.add_argument(
        "--id", nargs="+", metavar="ID",
        help="IDs des fichiers à synchroniser (ex: UO-001 UO-002)"
    )
    p_sync.add_argument(
        "--type", nargs="+", metavar="TYPE",
        help="Types à synchroniser (ex: uo_instance referentiel_uo)"
    )
    p_sync.add_argument(
        "--force", action="store_true",
        help="Ignorer la vérification de verrouillage"
    )

    # ── status ────────────────────────────────────────────────────────────────
    p_status = sub.add_parser("status", help="Affiche le contenu du store central")
    p_status.add_argument(
        "--prefix", metavar="PREFIX",
        help="Filtrer les clés par préfixe (ex: uo.)"
    )

    # ── store ─────────────────────────────────────────────────────────────────
    p_store = sub.add_parser("store", help="Opérations bas-niveau sur le store")
    store_sub = p_store.add_subparsers(dest="store_command", metavar="sous-commande")
    store_sub.required = True

    p_get = store_sub.add_parser("get", help="Lire une clé du store")
    p_get.add_argument("key", help="Nom de la variable (ex: uo.UO-001.avancement)")

    p_set = store_sub.add_parser("set", help="Écrire une clé dans le store")
    p_set.add_argument("key", help="Nom de la variable")
    p_set.add_argument("value", help="Valeur (JSON ou chaîne)")

    p_del = store_sub.add_parser("delete", help="Supprimer une clé du store")
    p_del.add_argument("key", help="Nom de la variable")

    store_sub.add_parser("clear", help="Vider toutes les variables du store")

    # ── lineage ───────────────────────────────────────────────────────────────
    p_lin = sub.add_parser("lineage", help="Affiche le graphe de dependances (Exomap)")
    p_lin.add_argument(
        "--id", metavar="FILE_ID",
        help="Filtrer sur un fichier (ex: UO-001)"
    )
    p_lin.add_argument(
        "--json", action="store_true",
        help="Sortie JSON brute"
    )

    # ── generate ──────────────────────────────────────────────────────────────
    p_gen = sub.add_parser("generate", help="Genere les fichiers Excel UO")
    p_gen.add_argument(
        "--id", nargs="+", metavar="ID",
        help="IDs des UO a generer (genere tout si absent)"
    )
    p_gen.add_argument(
        "--output", metavar="DIR",
        help="Repertoire de sortie (defaut: output/UOs/)"
    )

    # ── doctor ────────────────────────────────────────────────────────────────
    sub.add_parser("doctor", help="Diagnostique la sante de l ecosysteme")

    # ── doc ───────────────────────────────────────────────────────────────────
    p_doc = sub.add_parser("doc", help="Genere la documentation HTML de l ecosysteme")
    p_doc.add_argument(
        "--output", metavar="DIR",
        help="Repertoire de sortie (defaut: output/doc/)"
    )

    # ── history ───────────────────────────────────────────────────────────────
    p_hist = sub.add_parser("history", help="Historique des runs et snapshots")
    p_hist.add_argument(
        "--last", type=int, default=10, metavar="N",
        help="Nombre de runs a afficher (defaut: 10)"
    )
    p_hist.add_argument(
        "--compare", action="store_true",
        help="Compare les 2 derniers snapshots"
    )
    p_hist.add_argument(
        "--key", metavar="CLE",
        help="Historique d une cle specifique dans les snapshots"
    )

    return parser


# ─── Dispatch ─────────────────────────────────────────────────────────────────

def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if args.command == "sync":
        return cmd_sync(args)

    if args.command == "status":
        return cmd_status(args)

    if args.command == "store":
        if args.store_command == "get":
            return cmd_store_get(args)
        if args.store_command == "set":
            return cmd_store_set(args)
        if args.store_command == "delete":
            return cmd_store_delete(args)
        if args.store_command == "clear":
            return cmd_store_clear(args)

    if args.command == "lineage":
        return cmd_lineage(args)

    if args.command == "generate":
        return cmd_generate(args)

    if args.command == "doctor":
        return cmd_doctor(args)

    if args.command == "history":
        return cmd_history(args)

    if args.command == "doc":
        return cmd_doc(args)

    parser.print_help()
    return 0


if __name__ == "__main__":
    sys.exit(main())
