# Escalade — quand et comment

Le mécanisme complet (format, exemple) vit dans `ESCALADES.md` à la racine du
dépôt — ce fichier explique juste **quand** l'utiliser.

## Décider si c'est une escalade

Pose-toi cette question : *"Est-ce que je peux résoudre ça en modifiant
uniquement mon territoire (générateurs, `xl_design.py`, `_Manifeste`) ?"*

- **Oui** → pas d'escalade, fais-le directement
- **Non, ça touche `src/parser.py`, `src/executor.py` ou `src/models.py`** →
  escalade

## Exemples de décision

| Besoin | Escalade ? | Pourquoi |
|---|---|---|
| Ajouter une colonne "priorité" dans le cockpit | Non | Modification du générateur, ton territoire |
| Changer la couleur des badges de statut | Non | `src/xl_design.py`, ton territoire |
| Utiliser `GROUP_BY` dans un `_Manifeste` alors que ça n'a jamais été fait avant | Non | La fonction existe déjà dans le moteur (voir `02-Langage-MXL.md`) — utilise-la, pas besoin d'attendre |
| Une fonction `COMPUTE` qui n'existe pas du tout dans le moteur (ex. médiane) | Oui | Ça touche `src/executor.py` |
| Une nouvelle instruction MXL qui n'a pas d'équivalent actuel | Oui | Ça touche `src/parser.py` |
| Une question sur la structure d'un objet partagé (`UOInstance`, `Activity`...) | Oui | Ça touche `src/models.py` |

## Écrire l'entrée

Ouvre `ESCALADES.md` à la racine du dépôt, ajoute une entrée en tête de la
section "Entrées" en suivant le format documenté dans ce même fichier. Commit
et push comme n'importe quelle autre modification — c'est un fichier normal
du dépôt, versionné, pas un outil externe.

```bash
git add ESCALADES.md
git commit -m "docs(escalade): besoin fonction COMPUTE=MEDIANE"
git push
```

Pas de notification automatique : le binôme qui développe le cœur relit ce
fichier périodiquement. Si un besoin est urgent, un message direct reste
pertinent en complément de l'entrée écrite — l'entrée garde une trace, le
message accélère la prise en compte.
