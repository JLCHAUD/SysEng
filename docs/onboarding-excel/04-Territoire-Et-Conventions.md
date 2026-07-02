# Territoire et conventions

## Frontière de fichiers

| Territoire | Fichiers | Règle |
|---|---|---|
| Toi — libre | `src/generators/*.py`, `src/xl_design.py`, `projet_TrainSystem/creer_*.py`, les `_Manifeste` (via le code des générateurs) | Modifie librement, teste, commit |
| Le binôme cœur — escalade requise | `src/parser.py`, `src/executor.py`, `src/models.py` | Ne modifie pas directement — passe par `ESCALADES.md` (voir `05-Escalade.md`) |

`src/models.py` est côté "cœur" même s'il ressemble à une structure de
données neutre : il définit les objets partagés (UO, activités, acteurs...)
entre le moteur et les générateurs — une modification y a un impact que le
binôme cœur doit valider.

Tout fichier non listé explicitement ici (par exemple `src/store.py`,
`src/sync.py`, `src/ecosystem.py`, le dossier `web/`, ou un script
`projet_TrainSystem/` qui n'est pas un `creer_*.py`) est par défaut côté
"cœur" — pas de modification directe, escalade si besoin via `ESCALADES.md`.

## Deux niveaux d'intervention

1. **Modification manuelle d'un `_Manifeste`** dans un fichier Excel déjà
   généré : utile pour un test rapide, mais **volatile** — les `.xlsx` sont
   gitignorés, rien n'est versionné, et la prochaine régénération du fichier
   par son générateur écrasera la modification.
2. **Modification du code du générateur Python** correspondant : c'est la
   voie normale pour tout changement qui doit survivre — le générateur
   redevient la source de vérité, reproductible à chaque régénération.

En pratique : valide une idée rapidement au niveau 1 si besoin, mais reporte
toujours le résultat au niveau 2 avant de considérer le travail terminé.

## Convention de tests (TDD)

Le dépôt suit un principe strict : un test qui échoue avant le code, jamais
l'inverse. Outils : `pytest` + fixture `tmp_path` (répertoire temporaire
isolé par test) + `openpyxl.load_workbook` (relecture du fichier généré pour
inspection).

```python
def test_avancement_visible_dans_synthese(tmp_path):
    path = generate_dashboard_metier(acteur, uos, store, output_dir=tmp_path)
    wb = load_workbook(path, data_only=True)
    ws = wb["Synthèse"]
    # ... assertions sur le contenu de la feuille générée
```

Avant toute modification : lancer `pytest` pour confirmer que tout passe.
Après toute modification : écrire le test qui décrit le comportement attendu,
le voir échouer, implémenter, le voir passer.

Exception : pour tester des primitives de style isolées (comme dans
`tests/test_xl_design.py`), un `openpyxl.Workbook()` construit en mémoire
suffit — pas besoin de `tmp_path`/`load_workbook` si le test n'exerce pas un
générateur complet de bout en bout.

## Workflow Git

- Une branche dédiée par sujet de travail (`git checkout -b feature/<nom>`),
  jamais de commit direct sur `master`
- Commits fréquents et atomiques — un commit = un changement cohérent et
  testé, pas un gros commit en fin de journée
- Pull request vers `master` quand une brique de travail est terminée et que
  `pytest` passe intégralement
- Message de commit au format `type(scope): description` (ex.
  `feat(cockpit): ajoute colonne priorite`, `fix(dashboard): corrige filtre pilote_id`)
