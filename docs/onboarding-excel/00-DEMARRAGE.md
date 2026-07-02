# Démarrage — package de contexte Excel ExoSync

Tu es un LLM qui va aider à développer la partie **Excel** du projet ExoSync :
fichiers UO, cockpits ingénieur, dashboards métier — la génération de ces
fichiers (code Python) et leur contenu (`_Manifeste` MXL).

## Comment utiliser ce dossier

Ne colle pas tout ce dossier d'un coup. Selon la tâche du jour, ouvre le(s)
fichier(s) pertinent(s) et donne-les à ton LLM en contexte :

| Tâche | Fichiers à donner en contexte |
|---|---|
| Comprendre le projet dans son ensemble | `01-Vue-Ensemble-Projet.md` |
| Écrire ou modifier un `_Manifeste` | `02-Langage-MXL.md` |
| Modifier la mise en forme d'un Excel généré | `03-Design-System-XD.md` |
| Modifier un générateur Python (`creer_uo.py`, `cockpit_ingenieur_generator.py`, `dashboard_metier_generator.py`) | `02-Langage-MXL.md` + `03-Design-System-XD.md` + `04-Territoire-Et-Conventions.md` |
| Committer / ouvrir une PR | `04-Territoire-Et-Conventions.md` |
| Signaler un besoin qui dépasse ton périmètre | `05-Escalade.md` |

## Résumé express

ExoSync est un moteur Python qui synchronise des données à travers un
écosystème de fichiers Excel. Chaque fichier porte sa propre feuille
`_Manifeste`, écrite dans un petit langage (MXL), qui décrit ce que le fichier
donne et reçoit des autres fichiers — pas de base de données centrale qui
décrit la structure, elle vit dans les fichiers eux-mêmes.

Le projet est réparti en deux territoires (voir
`04-Territoire-Et-Conventions.md` pour le détail) :
- **Le cœur du moteur** (`src/parser.py`, `src/executor.py`, `src/models.py`)
  — pas ton périmètre, en cas de besoin voir `05-Escalade.md`
- **La partie Excel** (générateurs, design system, `_Manifeste`) — ton
  périmètre, libre

## Setup technique

```bash
cd chemin/vers/SysEng
pip install -r requirements.txt
pytest          # doit passer à 0 échec avant de commencer à modifier quoi que ce soit
python -m src sync --dir projet_TrainSystem   # synchronise après une modification
```

Python 3.11+ requis (voir README.md et la matrice CI du dépôt). Dépendances
principales : `openpyxl` (lecture/écriture Excel), `pytest` (tests), `click`
(CLI).

**Important** : tous les fichiers `.xlsx` du dépôt sont dans `.gitignore`.
Une modification manuelle d'un `_Manifeste` directement dans Excel n'est
**jamais** versionnée — seul le code Python qui génère ce `_Manifeste` l'est.
Pour qu'une modification survive, il faut soit la reporter dans le générateur
Python correspondant, soit accepter qu'elle reste locale (test ponctuel).
