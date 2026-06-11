"""
valider_un.py -- Synchronise UN SEUL fichier Excel, par chemin.
================================================================

Outil du Parcours de Validation du moteur ExoSync.
Permet de tester un fichier fait main SANS l'inscrire au registre :
on parse sa feuille _Manifeste, on l'exécute, on affiche ce qui s'est
passé (PULL / PUSH / BIND) et l'état du store.

Usage :
    python scripts/valider_un.py <chemin_du_fichier.xlsx>
    python scripts/valider_un.py validation/marche1_uo.xlsx

Astuce : pour repartir d'un store vide entre deux essais :
    python -m src store clear
"""

import sys
from pathlib import Path

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from src import store as Store
from src.executor import execute_ast
from src.parser import parse_file


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    if not argv:
        print("Usage : python scripts/valider_un.py <fichier.xlsx>")
        return 1

    path = Path(argv[0])
    if not path.is_absolute():
        path = ROOT / path

    print("=" * 65)
    print("  ExoSync - Validation d'un fichier")
    print("  Fichier : %s" % path)
    print("=" * 65)

    if not path.exists():
        print("  [ERR] Fichier introuvable.")
        return 1

    # 1) Lecture du Manifeste -> AST
    ast = parse_file(path)
    if ast is None:
        print("  [ERR] Aucune feuille _Manifeste trouvee dans ce fichier.")
        return 1

    print("\n-- Manifeste lu -----------------------------------------------")
    print("  FILE_TYPE : %s" % (ast.header.file_type or "<vide>"))
    print("  FILE_ID   : %s" % (ast.header.file_id or "<vide>"))

    if ast.errors:
        print("\n-- Erreurs de syntaxe (parsing) -------------------------------")
        for e in ast.errors:
            print("  [WARN] L%s : %s" % (getattr(e, "line_num", "?"), e.message))

    # 2) Execution sur le fichier Excel
    result = execute_ast(ast, path, Store)

    print("\n-- Resultat de l'execution ------------------------------------")
    print("  PULL  (tables remplies depuis le store) : %d" % len(result.pulled))
    for x in result.pulled:
        print("        <- %s" % x)
    print("  PUSH  (cles ecrites dans le store)      : %d" % len(result.pushed))
    for x in result.pushed:
        print("        -> %s" % x)
    print("  BIND  (cellules Dashboard mises a jour) : %d" % len(result.bound))
    for x in result.bound:
        print("        = %s" % x)

    if result.errors:
        print("\n-- ERREURS ----------------------------------------------------")
        for err in result.errors:
            print("  [ERR] %s" % err)
    if result.warnings:
        print("\n-- Avertissements ---------------------------------------------")
        for warn in result.warnings:
            print("  [WARN] %s" % warn)

    # 3) Etat du store
    print("\n-- Store (toutes les cles) ------------------------------------")
    all_vars = Store.get_all()
    if not all_vars:
        print("  (store vide)")
    for k, v in sorted(all_vars.items()):
        if isinstance(v, list):
            print("  %-45s -> [%d lignes]" % (k, len(v)))
        elif isinstance(v, float):
            print("  %-45s -> %s" % (k, v))
        else:
            print("  %-45s -> %s" % (k, v))

    print("\n" + "=" * 65)
    ok = len(result.errors) == 0
    if ok:
        print("  [OK] Execution terminee sans erreur.")
        print("  -> Ouvre le fichier Excel pour verifier l'onglet Dashboard.")
    else:
        print("  [ERR] %d erreur(s). Corrige le Manifeste et relance." % len(result.errors))
    print("=" * 65)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
