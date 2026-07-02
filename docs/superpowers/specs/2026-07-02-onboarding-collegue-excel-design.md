# Onboarding collègue — développement Excel autonome (LLM tiers)

- **Date** : 2026-07-02
- **Conversation** : CONV-14 — Onboarding Collègue Excel
- **Branche cible** : `feature/onboarding-excel`
- **Statut** : spec validée, prête pour plan d'implémentation
- **Dépôt** : `SysEng` (github.com/JLCHAUD/SysEng), même dépôt que le cœur

---

## 1. Objectif & périmètre

Une collègue va développer une partie d'ExoSync avec un **LLM autre que Claude
Code** (donc sans accès aux skills, à la mémoire, ni à ce vault Obsidian). Elle
doit pouvoir travailler de façon **autonome** sur les fichiers Excel du projet,
sans dépendre en continu du binôme user+Claude qui développe le cœur du moteur.

Ce projet livre un **package de documentation embarqué dans le dépôt Git**
(`docs/onboarding-excel/`) qui donne à n'importe quel LLM assez de contexte
pour :
- comprendre le projet (but, architecture, état d'avancement)
- écrire/modifier des `_Manifeste` MXL valides
- modifier les générateurs Python qui produisent les Excel
- respecter la charte graphique (`XD`) et les conventions du dépôt
- savoir quand et comment remonter un besoin qui dépasse son périmètre

Périmètre : uniquement la rédaction des documents d'onboarding et du fichier
`ESCALADES.md`. Aucune modification du code existant (moteur ou générateurs)
n'est nécessaire pour ce projet.

**Hors périmètre (explicitement reporté)** : l'interface web de lancement de
création d'UO, mentionnée par l'utilisateur comme une amélioration future —
non traitée ici, à reprendre dans une conversation dédiée une fois cette
collaboration rodée.

---

## 2. Contexte — répartition des rôles

| Qui | Quoi |
|---|---|
| User + Claude (ici) | Cœur du moteur : `src/parser.py`, `src/executor.py`, `src/models.py` — le langage MXL lui-même et son exécution |
| Collègue + son LLM | Tout ce qui produit/façonne les Excel : `src/generators/*.py`, `src/xl_design.py`, `projet_TrainSystem/creer_*.py`, les `_Manifeste` des fichiers UO/cockpit/dashboard |

Elle a accès au même dépôt Git, travaille sur sa propre branche, et
synchronise par pull/push + pull request. Elle est habilitée à voir les
données réelles du projet (contenu Alstom) : le package de contexte n'a pas
besoin d'être générique ou anonymisé.

**Deux niveaux d'intervention prévus** :
1. Modification manuelle d'un `_Manifeste` existant (ajustement ponctuel)
2. Modification du code Python des générateurs, pour que les changements
   soient reproductibles à chaque régénération (pas seulement une rustine à la main)

---

## 3. Architecture du package

Tout vit dans le dépôt Git `SysEng` sous `docs/onboarding-excel/`, en plus
d'un fichier `ESCALADES.md` à la racine du dépôt. Ce choix est structurant :
son LLM ne voit que ce qui est versionné dans `SysEng` — jamais ce vault
Obsidian, jamais la mémoire de Claude Code. Tout contexte nécessaire doit donc
être dupliqué/condensé **dans le dépôt**, même s'il existe déjà ailleurs
(vault, specs superpowers) sous une autre forme.

Structure modulaire plutôt qu'un document unique : chaque fichier cible une
tâche, elle colle celui qui correspond à ce qu'elle fait ce jour-là plutôt que
de recoller un pavé complet à chaque session. Un fichier stable (le langage
MXL) n'est pas mélangé avec un fichier qui évolue vite (l'état du projet).

```
SysEng/
├── ESCALADES.md                          ← nouveau, racine du dépôt
└── docs/onboarding-excel/
    ├── 00-DEMARRAGE.md
    ├── 01-Vue-Ensemble-Projet.md
    ├── 02-Langage-MXL.md
    ├── 03-Design-System-XD.md
    ├── 04-Territoire-Et-Conventions.md
    └── 05-Escalade.md
```

---

## 4. Contenu détaillé par fichier

### `00-DEMARRAGE.md`
Point d'entrée. Explique comment utiliser le dossier avec un LLM tiers :
quel(s) fichier(s) coller selon la tâche du jour (ex. « créer un cockpit » →
coller 02 + 03 + 04). Contient un résumé express du projet en quelques lignes
et le setup technique minimal : Python 3.14, `pip install -r requirements.txt`,
`pytest` pour vérifier que tout passe avant de commencer, `python -m src sync
--dir projet_TrainSystem` pour synchroniser après modification.

### `01-Vue-Ensemble-Projet.md`
Le but d'ExoSync et le concept d'**exostructure** (la vérité vit dans les
`_Manifeste` de chaque fichier, pas en base centrale). La pyramide des vues
(UO instance → cockpit ingénieur → dashboard métier → dashboard client). État
d'avancement condensé : ce qui est livré (C1 modèle UO, C2 pyramide interne,
design system), ce qui reste (C3 vues externes, C4 assembleur d'instanciation).

### `02-Langage-MXL.md`
Référence complète des instructions `_Manifeste` utilisées dans le projet
(`LIST`, `COLLECT`, `PULL`, `PUSH`, `BIND`, `COMPUTE`, `VALIDATE`, `FILE_TYPE`/
`FILE_ID`), avec des exemples réels tirés des fichiers du projet. Convention de
nommage des UO (`L{NN}U{NN}-{PPPP}-{SSSS}`) et des clés du store
(`uo.{id}.champ`, `cockpit.{nom}.mes_uos`, `dashboard.{id}.synthese`).
Distingue explicitement le sous-ensemble **stable/garanti** de ce qui reste en
discussion côté moteur (brique B de la roadmap, gelée) — pour qu'elle sache sur
quoi elle peut s'appuyer sans risque de changement de syntaxe.

### `03-Design-System-XD.md`
Comment utiliser `src/xl_design.py` : les méthodes principales (`XD.banner()`,
`XD.table_header()`, `XD.named_table()`, `XD.health_spine()`, badges de statut/
criticité) et la palette des 11 familles d'onglets. Deux règles non-négociables
mises en avant : zéro style Excel inline (tout passe par `XD`), et la colonne
spine santé ne doit jamais faire partie du range d'une table nommée (sinon
`GET_TABLE`/`COLLECT` lit une colonne en trop côté moteur).

### `04-Territoire-Et-Conventions.md`
Rappelle la frontière de fichiers (tableau §2 ci-dessus). Détaille la
convention TDD du dépôt : un test qui échoue avant le code, `pytest` +
`tmp_path` + `load_workbook` pour inspecter les fichiers générés. Décrit le
workflow Git attendu : une branche dédiée par sujet, des commits fréquents et
atomiques, une pull request vers `master` une fois une brique de travail
terminée et testée.

### `05-Escalade.md`
Explique le fichier `ESCALADES.md` : format d'entrée attendu (date, contexte,
besoin précis, exemple concret de ce qu'elle voudrait pouvoir écrire dans un
`_Manifeste` mais qui n'est pas encore supporté), et la distinction entre ce
qui relève d'une escalade (instruction ou fonction MXL manquante côté moteur)
et ce qu'elle peut résoudre seule (mise en forme, structure de feuille,
logique dans un générateur existant).

---

## 5. Mécanisme d'escalade — `ESCALADES.md`

Fichier unique à la racine du dépôt, tenu par la collègue, relu périodiquement
par l'utilisateur. Chaque entrée suit un format simple :

```markdown
## 2026-07-15 — Besoin d'une fonction COMPUTE=MEDIANE

**Contexte** : dashboard métier, calcul du délai médian de clôture des UO.
**Ce qui manque** : COMPUTE ne supporte que MEAN_WEIGHTED et SUM actuellement.
**Ce que je voudrais écrire** :
`COMPUTE $delai_median = MEDIANE(tbl_uos, col=delai_jours)`
**Contournement actuel** : calcul en dur côté générateur Python (pas idéal,
pas synchronisé si le store change).
```

Pas d'automatisation de notification (pas de webhook, pas de CI) — l'utilisateur
relit le fichier manuellement. Cohérent avec le mode de collaboration actuel
(sync manuelle, pas d'outillage lourd tant que le rodage n'est pas terminé).

---

## 6. Maintenance du package

Les documents `01` (état du projet) et potentiellement `02` (si le langage
MXL évolue côté moteur) devront être mis à jour par l'utilisateur ou Claude
au fil des CONV suivantes, comme n'importe quelle doc de spec du dépôt. Les
documents `03` (design system) et `04` (conventions) sont plus stables et ne
changent que si les conventions elles-mêmes changent.

Pas de mécanisme de synchronisation automatique prévu dans ce projet — la
mise à jour reste manuelle, à la charge du binôme user+Claude quand une
évolution du cœur ou des conventions impacte ce que la collègue doit savoir.

---

## 7. Critères de succès

- Les 6 fichiers de `docs/onboarding-excel/` existent et sont commités sur le
  dépôt `SysEng`
- `ESCALADES.md` existe à la racine avec le format d'entrée documenté et un
  exemple
- Un LLM sans aucun contexte préalable, en lisant uniquement `00-DEMARRAGE.md`
  puis les fichiers qu'il pointe, peut expliquer correctement : le but du
  projet, où se trouve sa frontière de travail, comment écrire un `_Manifeste`
  MXL valide, et comment escalader un besoin
- Aucune modification du code existant (moteur, générateurs) — projet
  purement documentaire
