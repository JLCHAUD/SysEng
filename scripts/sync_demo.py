"""
sync_demo.py -- Synchronise l'ecosysteme de demonstration ExoSync.

Execute la chaine PUSH -> PULL -> BIND dans le bon ordre :
  1. UOs (PUSH activites -> store)
  2. Cockpits (PULL activites -> COMPUTE -> BIND Dashboard)
  3. Pilote (PULL toutes activites -> COMPUTE -> BIND Dashboard)

Resultats attendus :
  COCKPIT-ALICE  : av_uo1=70%  av_uo2=40%  av_global=55%
  COCKPIT-BRUNO  : av_uo1=80%  av_uo2=15%  av_global=47.5%
  PILOTE         : av_alice=55%  av_bruno=47.5%

Usage :
    python scripts/sync_demo.py
"""

import sys
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from src import store as Store
from src.executor import execute_ast
from src.parser import parse_file

DEMO_DIR = ROOT / "output" / "demo"

# Fichiers a synchroniser (dans l'ordre de dependance)

DEMO_FILES = [
    # Phase 1 -- UOs : PUSH vers le store
    ("DEMO-UO-A1",         "uo_instance", "UO-A1_analyse_systeme.xlsx"),
    ("DEMO-UO-A2",         "uo_instance", "UO-A2_validation_fonct.xlsx"),
    ("DEMO-UO-B1",         "uo_instance", "UO-B1_architecture_sys.xlsx"),
    ("DEMO-UO-B2",         "uo_instance", "UO-B2_essais_terrain.xlsx"),
    # Phase 2 -- Cockpits : PULL + COMPUTE + BIND
    ("DEMO-COCKPIT-ALICE", "cockpit",     "COCKPIT-ALICE.xlsx"),
    ("DEMO-COCKPIT-BRUNO", "cockpit",     "COCKPIT-BRUNO.xlsx"),
    # Phase 3 -- Pilote : PULL + COMPUTE + BIND
    ("DEMO-PILOTE",        "pilote",      "PILOTE.xlsx"),
]


def _fmt_result(file_id, result):
    ok = len(result.errors) == 0
    status = "[OK] " if ok else "[ERR]"
    return (
        "  %s %-25s PULL=%2d  PUSH=%2d  BIND=%2d  err=%d" % (
            status, file_id,
            len(result.pulled), len(result.pushed),
            len(result.bound), len(result.errors),
        )
    )


def main():
    print("=" * 65)
    print("  ExoSync - Synchronisation demo")
    print("=" * 65)

    erreurs_globales = 0

    for file_id, file_type, filename in DEMO_FILES:
        path = DEMO_DIR / filename

        # Separateur de phase
        if file_id == "DEMO-UO-A1":
            print("\n-- Phase 1 : UOs (PUSH -> store) -----------------------------")
        elif file_id == "DEMO-COCKPIT-ALICE":
            print("\n-- Phase 2 : Cockpits (PULL -> COMPUTE -> BIND) --------------")
        elif file_id == "DEMO-PILOTE":
            print("\n-- Phase 3 : Pilote (PULL -> COMPUTE -> BIND) ----------------")

        if not path.exists():
            print("  [ERR]  %-25s Fichier manquant" % file_id)
            print("         -> Lancez d'abord : python scripts/generate_demo.py")
            erreurs_globales += 1
            continue

        ast = parse_file(path)
        if ast is None:
            print("  [ERR]  %-25s Pas de feuille _Manifeste" % file_id)
            erreurs_globales += 1
            continue

        if ast.errors:
            for e in ast.errors:
                print("  [WARN] Parse %s L%d: %s" % (file_id, e.line_num, e.message))

        result = execute_ast(ast, path, Store)
        print(_fmt_result(file_id, result))

        for err in result.errors:
            print("         [ERR] %s" % err)
        for warn in result.warnings:
            print("         [WARN] %s" % warn)

        erreurs_globales += len(result.errors)

    # Verification du store
    print("\n-- Store -- cles demo ------------------------------------------------")
    all_vars = Store.get_all()
    demo_keys = {k: v for k, v in all_vars.items() if k.startswith("demo.")}
    for k, v in sorted(demo_keys.items()):
        if isinstance(v, list):
            print("  %-42s -> [%d lignes]" % (k, len(v)))
        elif isinstance(v, float):
            print("  %-42s -> %.1f%%" % (k, v * 100))
        else:
            print("  %-42s -> %s" % (k, v))

    # Resume
    print("\n" + "=" * 65)
    if erreurs_globales == 0:
        print("  [OK] Synchronisation demo terminee SANS erreur.")
        print()
        print("  Resultats attendus :")
        print("    COCKPIT-ALICE  : UO-A1=70%  UO-A2=40%  Global=55%")
        print("    COCKPIT-BRUNO  : UO-B1=80%  UO-B2=15%  Global=47.5%")
        print("    PILOTE         : Alice=55%   Bruno=47.5%")
        print()
        print("  Ouvrez les fichiers Excel dans output/demo/ pour verifier")
        print("  les valeurs dans l'onglet Dashboard (cellule F3/F4/F5).")
    else:
        print("  [ERR] %d erreur(s) detectee(s)." % erreurs_globales)
    print("=" * 65)


if __name__ == "__main__":
    main()
