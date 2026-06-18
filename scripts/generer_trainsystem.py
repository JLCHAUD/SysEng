"""
Génère cockpits ingénieurs + dashboard pilote dans projet_TrainSystem/.

Tout est produit dans le même répertoire que les fichiers UO pour
pouvoir tester le cycle complet sans changer de dossier.

Usage :
    python scripts/generer_trainsystem.py

Sortie :
    projet_TrainSystem/Cockpit_Alice_Dubois.xlsx
    projet_TrainSystem/Cockpit_Bruno_Lecomte.xlsx
    projet_TrainSystem/Cockpit_Camille_Vidal.xlsx
    projet_TrainSystem/Dashboard_Jean-Luc_Bernard.xlsx
    projet_TrainSystem/store.json
"""

import io
import sys
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from src.config_loader import load_uo_instances, load_acteurs
from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
from src.generators.dashboard_metier_generator import generate_dashboard_metier
from src.store import JsonStore

PROJET_DIR = ROOT / "projet_TrainSystem"
STORE_PATH = PROJET_DIR / "store.json"

# Données simulées : ce qu'un ingénieur aurait saisi dans son cockpit
SAISIES_SIMULEES = {
    "uo.UO-001.avancement": 0.65,
    "uo.UO-001.heures_realisees": 20.0,
    "uo.UO-002.avancement": 0.40,
    "uo.UO-002.heures_realisees": 52.0,   # > 48h allouées → alerte dérive
    "uo.UO-003.avancement": 0.20,
    "uo.UO-003.heures_realisees": 8.0,
    "uo.UO-004.avancement": 0.50,
    "uo.UO-004.heures_realisees": 12.0,
    "uo.UO-005.avancement": 0.75,
    "uo.UO-005.heures_realisees": 27.0,
}


def banner(title: str):
    print(f"\n{'-' * 60}")
    print(f"  {title}")
    print(f"{'-' * 60}")


def main():
    PROJET_DIR.mkdir(parents=True, exist_ok=True)

    banner("Chargement de la configuration")
    all_uos = load_uo_instances()
    acteurs = load_acteurs()
    pilotes_metier = [a for a in acteurs if a.role.value == "pilote_metier"]

    engineers = sorted(set(u.engineer_name for u in all_uos))
    print(f"  UOs chargées   : {len(all_uos)}")
    print(f"  Ingénieurs     : {', '.join(engineers)}")
    print(f"  Pilotes métier : {', '.join(a.nom for a in pilotes_metier)}")
    print(f"  Dossier sortie : {PROJET_DIR}")

    # Étape 1 : Cockpits ingénieurs
    banner("Étape 1 — Génération des cockpits ingénieurs")
    for eng in engineers:
        path = generate_cockpit_ingenieur(
            eng, all_uos,
            output_dir=PROJET_DIR,
            uo_dir=PROJET_DIR,
        )
        nb_uos = len([u for u in all_uos if u.engineer_name == eng])
        print(f"  ✅  {path.name}  ({nb_uos} UOs)")

    # Étape 2 : Injection des données simulées
    banner("Étape 2 — Injection des données dans le store")
    store = JsonStore(STORE_PATH)
    store.set_many(SAISIES_SIMULEES)
    print(f"  {len(SAISIES_SIMULEES)} valeurs → {STORE_PATH.relative_to(ROOT)}")

    # Étape 3 : Dashboards pilotes métier
    banner("Étape 3 — Génération des dashboards pilotes métier")
    for pilote in pilotes_metier:
        path = generate_dashboard_metier(pilote, all_uos, store, output_dir=PROJET_DIR)
        uo_filtrees = len([u for u in all_uos
                           if u.engineer_name in pilote.filtre_valeur.split(",")])
        print(f"  ✅  {path.name}  ({uo_filtrees} UOs dans périmètre)")

    # Résumé
    banner("Résumé — Fichiers dans projet_TrainSystem/")
    generated = []
    for f in sorted(PROJET_DIR.glob("*.xlsx")):
        size_kb = f.stat().st_size // 1024
        generated.append(f.name)
        print(f"  {f.name:<50} {size_kb:>4} Ko")

    print(f"\n  ✅  {len(generated)} fichiers Excel dans {PROJET_DIR.name}/")
    print("  Ouvrez-les dans Excel pour vérification manuelle.\n")


if __name__ == "__main__":
    main()
