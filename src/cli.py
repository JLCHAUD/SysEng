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
    print(f"  {len(generated)} fichier(s) généré(s), {len(errors)} erreur(s)")
    return 0 if not errors else 1


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
    p_gen = sub.add_parser("generate", help="Génère les fichiers Excel UO")
    p_gen.add_argument(
        "--id", nargs="+", metavar="ID",
        help="IDs des UO à générer (génère tout si absent)"
    )
    p_gen.add_argument(
        "--output", metavar="DIR",
        help="Répertoire de sortie (défaut: output/UOs/)"
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

    parser.print_help()
    return 0


if __name__ == "__main__":
    sys.exit(main())
