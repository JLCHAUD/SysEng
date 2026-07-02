# Onboarding Collègue Excel — Plan d'implémentation

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Livrer un package de documentation embarqué dans le dépôt Git (`docs/onboarding-excel/` + `ESCALADES.md`) permettant à une collègue de développer les Excel d'ExoSync (UO, cockpits, dashboards) avec un LLM tiers, de façon autonome et sans accès à ce vault Obsidian ni à Claude Code.

**Architecture:** 7 fichiers markdown autonomes, chacun ciblant une tâche précise plutôt qu'un seul pavé monolithique. Le contenu MXL (`02-Langage-MXL.md`) est vérifié ligne à ligne contre le code source actuel (`src/parser.py`, `src/executor.py`) — pas recopié depuis une doc externe. Projet purement documentaire : aucune modification de code.

**Tech Stack:** Markdown. Dépôt Git `SysEng` (github.com/JLCHAUD/SysEng). Pas de dépendance technique nouvelle.

**Spec de référence :** `docs/superpowers/specs/2026-07-02-onboarding-collegue-excel-design.md`

---

## Contexte technique à lire avant de commencer

Ce plan ne suit pas le cycle TDD classique (pas de code à tester) : chaque tâche
écrit un fichier markdown avec un contenu complet et vérifié, puis relit le
fichier pour s'assurer qu'il ne contient aucune contradiction avec le code
source réel du dépôt.

**Faits vérifiés dans le code source** (base de toutes les tâches ci-dessous,
extraits confirmés par lecture complète de `src/parser.py`, `src/executor.py`,
`src/generators/cockpit_ingenieur_generator.py`,
`src/generators/dashboard_metier_generator.py`,
`projet_TrainSystem/creer_uo.py`, `src/store.py`) :

- Format physique `_Manifeste` : mono-colonne, ligne 1 = `MANIFESTE_V=1`,
  ligne 2 vide, instructions à partir de la ligne 3 (`parser.py:749`)
- Instructions supportées par le parser actuel : `FILE_TYPE`, `FILE_ID`,
  `VERSION`, `DOC` (en-tête), métadonnées libres (`clé: valeur`), `DEF`, `COL`,
  `BIND`, `PUSH`, `VALIDATE`, `NOTIFY`, `PULL` (`FILL_TABLE`/`UPDATE_CELLS`),
  `LIST` (formes `TABLE` et `DYNAMIC`), `COLLECT`, `IDENT`, `EXTENDS`
- Fonctions `COMPUTE`/agrégation : `FILTER`, `MEAN_WEIGHTED`, `SUM`, `COUNT`,
  `COUNT_IF`, `AVG`, `MIN`, `MAX`, `DIV`, `SWITCH_RANGE`, `IF`, `IF_NULL`,
  `GROUP_BY`, `SORT`, `TOP_N`
- Règles `VALIDATE` : `NOT_NULL`, `POSITIVE`, `NON_NEGATIVE`, `UNIQUE`,
  `RANGE(min,max)`, `IN(...)`, `NOT_EMPTY`, `MAX_LENGTH(n)`, `MIN_LENGTH(n)`,
  `MATCHES(regex)`, `MAX(n)`, `MIN(n)`
- **Sous-ensemble éprouvé** (utilisé par les générateurs actuels, testé en
  production) : `FILE_TYPE`/`FILE_ID`, métadonnées libres (`ingenieur:`,
  `pilote_id:`), `DEF` (`GET_TABLE`, `COMPUTE(...)`), `COL ... WRITE=...`,
  `LIST TYPE=... WHERE ...`, `COLLECT ... FROM ... INTO ...`, `PUSH`, `BIND`,
  `VALIDATE`
- **Sous-ensemble disponible mais pas encore exploité** par les générateurs
  actuels (existe et fonctionne dans le moteur, mais aucun exemple généré en
  production) : `EXTENDS`, `NOTIFY`, `SWITCH_RANGE`, `GROUP_BY`, `SORT`,
  `TOP_N`, `PULL`, `IDENT`
- Conventions de clés du store observées : `cockpit.<nom_ingenieur>.mes_uos`,
  `dashboard.<pilote_id>.synthese`, `uo.<code>.<champ>` (ex.
  `uo.L09U1-CFL2400-CLIM.avancement`)
- Tous les fichiers `.xlsx` sont gitignorés (`.gitignore` racine, règle
  `*.xlsx`) — un `_Manifeste` modifié à la main dans Excel n'est **jamais**
  versionné ; seul le code Python qui le génère l'est

---

## Fichiers concernés

| Action | Fichier |
|--------|---------|
| Créer | `ESCALADES.md` (racine du dépôt) |
| Créer | `docs/onboarding-excel/00-DEMARRAGE.md` |
| Créer | `docs/onboarding-excel/01-Vue-Ensemble-Projet.md` |
| Créer | `docs/onboarding-excel/02-Langage-MXL.md` |
| Créer | `docs/onboarding-excel/03-Design-System-XD.md` |
| Créer | `docs/onboarding-excel/04-Territoire-Et-Conventions.md` |
| Créer | `docs/onboarding-excel/05-Escalade.md` |

---

### Task 0 : Créer la branche de travail

**Files:** aucun

- [ ] **Step 1 : Créer et basculer sur la branche**

```bash
cd "C:\Users\fabie\Documents\JLC\Python\SysEng"
git checkout master
git pull
git checkout -b feature/onboarding-excel
```

- [ ] **Step 2 : Créer le dossier de destination**

```bash
mkdir -p docs/onboarding-excel
```

---

### Task 1 : `ESCALADES.md` (racine du dépôt)

**Files:**
- Create: `ESCALADES.md`

- [ ] **Step 1 : Écrire le fichier**

```markdown
# Escalades — besoins moteur MXL

Ce fichier sert à faire remonter un besoin qui dépasse le périmètre "Excel"
(générateurs, `_Manifeste`, design system) et qui nécessiterait une évolution
du **cœur du moteur** (`src/parser.py`, `src/executor.py`, `src/models.py`).

**Qui écrit ici** : la personne qui développe la partie Excel du projet.
**Qui relit** : le binôme qui développe le cœur (relecture périodique
manuelle, pas de notification automatique).

## Ce qui relève d'une escalade

- Une instruction MXL qui n'existe pas dans `_Manifeste` et dont tu aurais
  besoin (nouveau mot-clé, nouvelle fonction `COMPUTE`, nouvelle règle
  `VALIDATE`)
- Un comportement du moteur qui semble incohérent ou bloquant lors d'un
  `python -m src sync`
- Une question de conception sur `src/models.py` (structure de données
  partagée entre le cœur et les générateurs)

## Ce qui NE relève PAS d'une escalade (tu peux le faire toi-même)

- Mise en forme Excel (tout passe par `src/xl_design.py` — voir
  `docs/onboarding-excel/03-Design-System-XD.md`)
- Structure d'une feuille, ajout d'une colonne dans un générateur existant
- Utilisation d'une instruction MXL qui existe déjà dans le moteur mais n'est
  pas encore utilisée par les générateurs (voir la liste "disponible mais pas
  exploité" dans `docs/onboarding-excel/02-Langage-MXL.md`) — dans ce cas,
  utilise-la directement, pas besoin d'escalade

## Format d'une entrée

```markdown
## AAAA-MM-JJ — Titre court du besoin

**Contexte** : quel fichier/générateur, quel objectif métier.
**Ce qui manque** : ce que le moteur ne permet pas de faire aujourd'hui.
**Ce que je voudrais écrire** : la syntaxe MXL que tu voudrais pouvoir
utiliser (même approximative — l'idée compte plus que la syntaxe exacte).
**Contournement actuel** : comment tu fais en attendant (souvent : logique
codée en dur côté générateur Python, pas idéal, pas synchronisé si le store
change).
```

## Exemple

## 2026-07-15 — Besoin d'une fonction COMPUTE=MEDIANE

**Contexte** : dashboard métier, calcul du délai médian de clôture des UO.
**Ce qui manque** : `COMPUTE` supporte `MEAN_WEIGHTED`, `SUM`, `AVG`, `MIN`,
`MAX` mais pas de médiane.
**Ce que je voudrais écrire** :
`DEF $delai_median = COMPUTE(MEDIANE($uos_closes.delai_jours))`
**Contournement actuel** : calcul en dur côté générateur Python (pas idéal,
pas synchronisé si le store change).

---

## Entrées

<!-- Ajoute tes entrées ci-dessous, la plus récente en premier -->
```

- [ ] **Step 2 : Vérifier le fichier**

Relire et confirmer : le format d'entrée est concret (pas de placeholder),
l'exemple est réaliste (`MEDIANE` n'existe effectivement pas dans la liste des
fonctions `COMPUTE` du moteur — vérifié dans `src/executor.py`), la distinction
escalade / non-escalade est claire.

- [ ] **Step 3 : Commit**

```bash
git add ESCALADES.md
git commit -m "docs(onboarding): ESCALADES.md - mecanisme de remontee de besoins moteur"
```

---

### Task 2 : `docs/onboarding-excel/00-DEMARRAGE.md`

**Files:**
- Create: `docs/onboarding-excel/00-DEMARRAGE.md`

- [ ] **Step 1 : Écrire le fichier**

```markdown
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

Python 3.14 requis. Dépendances principales : `openpyxl` (lecture/écriture
Excel), `pytest` (tests), `click` (CLI).

**Important** : tous les fichiers `.xlsx` du dépôt sont dans `.gitignore`.
Une modification manuelle d'un `_Manifeste` directement dans Excel n'est
**jamais** versionnée — seul le code Python qui génère ce `_Manifeste` l'est.
Pour qu'une modification survive, il faut soit la reporter dans le générateur
Python correspondant, soit accepter qu'elle reste locale (test ponctuel).
```

- [ ] **Step 2 : Vérifier le fichier**

Confirmer que le tableau de routage pointe vers des noms de fichiers qui
existeront bien après les tâches suivantes (`01` à `05`), que la commande
`python -m src sync --dir projet_TrainSystem` correspond à celle documentée
dans `CLAUDE.md` du dépôt, et que la remarque sur `.gitignore` est cohérente
avec la règle `*.xlsx` du fichier `.gitignore` racine.

- [ ] **Step 3 : Commit**

```bash
git add docs/onboarding-excel/00-DEMARRAGE.md
git commit -m "docs(onboarding): 00-DEMARRAGE - point d'entree du package de contexte"
```

---

### Task 3 : `docs/onboarding-excel/01-Vue-Ensemble-Projet.md`

**Files:**
- Create: `docs/onboarding-excel/01-Vue-Ensemble-Projet.md`

- [ ] **Step 1 : Écrire le fichier**

```markdown
# Vue d'ensemble du projet ExoSync

## Le problème que ça résout

Un ingénieur travaille sur un fichier Excel (une "UO", Unité d'Œuvre). Son
avancement doit remonter à son cockpit personnel, qui agrège toutes ses UO.
Ce cockpit doit à son tour remonter vers le dashboard de son responsable
métier, qui voit la synthèse de toute son équipe. Sans outil, ça veut dire
resaisir la même donnée à plusieurs endroits, ou construire des liaisons
Excel fragiles.

ExoSync automatise cette remontée : chaque fichier Excel porte une feuille
`_Manifeste` qui décrit ce qu'il donne (`PUSH`) et ce qu'il reçoit (`PULL`,
`LIST`/`COLLECT`) des autres fichiers de l'écosystème. Un script Python
(`python -m src sync`) lit tous les `_Manifeste`, exécute les instructions, et
met à jour les fichiers en conséquence.

## Le concept d'exostructure

Il n'existe **pas** de base centrale qui décrit "voici tous les fichiers du
projet et comment ils sont reliés". Cette structure émerge dynamiquement à
partir des `_Manifeste` : c'est le fichier lui-même qui porte sa propre
identité (`FILE_TYPE`, `FILE_ID`) et ses propres règles de synchronisation.
C'est pour ça qu'on parle d'**exostructure** — la structure vit à l'extérieur
d'une base centrale, distribuée dans les fichiers.

## La pyramide des vues

```
DASHBOARD CLIENT (Alstom, filtré par le leader métier)      ← pas encore fait
      ▲ re-PUSH namespace client.*
DASHBOARD MÉTIER (leader Train System)                       ← fait
      ▲ LIST mes_cockpits TYPE=cockpit_ingenieur WHERE pilote_id=...
      ▲ COLLECT tbl_mes_uos FROM mes_cockpits INTO tbl_vue_synthese
COCKPIT INGÉNIEUR (×20+, un par ingénieur) — hub du matin     ← fait
      ▲ PUSH cockpit.<nom>.mes_uos
UO INSTANCES (50-150 fichiers vivants)                        ← fait
      L{NN}U{NN}-{PPPP}-{SSSS}  (ex: L09U1-CFL2400-CLIM)
```

- Une **UO instance** est le fichier de travail quotidien d'un ingénieur :
  activités, points ouverts, livrables.
- Un **cockpit ingénieur** agrège toutes les UO d'un même ingénieur — c'est
  son hub du matin : avancement, alertes, liens vers ses fichiers.
- Un **dashboard métier** agrège tous les cockpits des ingénieurs d'un même
  responsable (filtré par `pilote_id`) — synthèse d'équipe.
- Le **dashboard client** (pas encore construit) re-publiera une vue filtrée
  vers le client final (Alstom), avec curation de ce qui est montré.

## Projet réel : catalogue Train System

Le projet concret est un catalogue d'Unités d'Œuvre pour le client Alstom,
domaine Train System (ingénierie système). Structure : 5 lots (phases projet)
× 2 UO (système / sous-système), ~20+ ingénieurs, 50-150 fichiers UO vivants
à terme. Convention de nommage : `L{NN}U{NN}-{PPPP}-{SSSS}` où `{PPPP}` est
le code projet et `{SSSS}` le code système (ex : `L09U1-CFL2400-CLIM` = lot 9,
UO1, projet CFL2400, système climatisation).

## État d'avancement (2026-07-02)

| Brique | Contenu | État |
|---|---|---|
| Noyau moteur | Parser + executor MXL | ✅ fait, 415 tests passants |
| Modèle UO réel | UO type Train System, catalogue 5 lots × 2 UO | ✅ fait |
| Pyramide interne | Cockpit ingénieur + dashboard métier fonctionnels | ✅ fait |
| Design system Excel | Charte graphique centralisée (`src/xl_design.py`) | ✅ fait |
| Vues externes | Dashboard client, vue par projet | ⬜ à faire |
| Chaîne d'instanciation | Assembleur : commande client → UO prête en < 5 min | 🔄 partiel (`creer_uo.py` existe, orchestration manuelle) |
| Contrat hybride | Contrat minimal UO figé (imposé vs libre) | ⬜ après les vues externes |

C'est précisément la partie "Excel" de ce tableau — modèle UO, pyramide
interne, design system, vues externes, chaîne d'instanciation — qui constitue
ton périmètre de travail.
```

- [ ] **Step 2 : Vérifier le fichier**

Confirmer la cohérence avec l'état réel du dépôt : 415 tests passants (vérifié
par `pytest` en session précédente), le design system est bien fini (branche
`feature/design-system` mergée, commit `db416bd`), les vues externes et
l'assembleur restent bien à faire (aucun code trouvé dans le dépôt pour un
dashboard client).

- [ ] **Step 3 : Commit**

```bash
git add docs/onboarding-excel/01-Vue-Ensemble-Projet.md
git commit -m "docs(onboarding): 01-Vue-Ensemble-Projet - but, pyramide, etat d'avancement"
```

---

### Task 4 : `docs/onboarding-excel/02-Langage-MXL.md`

**Files:**
- Create: `docs/onboarding-excel/02-Langage-MXL.md`

- [ ] **Step 1 : Écrire le fichier**

```markdown
# Le langage MXL — référence

MXL (Manifest Exchange Language) est le petit langage écrit dans la feuille
`_Manifeste` de chaque fichier Excel ExoSync. Cette référence est vérifiée
directement contre le code du moteur (`src/parser.py`, `src/executor.py`) —
pas une doc à part qui pourrait diverger du comportement réel.

## Format physique de la feuille `_Manifeste`

| Colonne A | Colonne B | Colonne C |
|---|---|---|
| Instruction MXL | Ancre Excel (optionnelle, rarement utilisée) | Commentaire en français |

- **Ligne 1** : toujours `MANIFESTE_V=1`
- **Ligne 2** : vide
- **Lignes 3+** : les instructions, une par ligne
- Les lignes vides sont ignorées, un mot-clé non reconnu lève une erreur de
  parsing explicite au sync

## En-tête

```
FILE_TYPE: uo_instance
FILE_ID: L09U1-CFL2400-CLIM
ingenieur: Alice Dubois
```

- `FILE_TYPE` (obligatoire) : type du fichier dans l'écosystème
  (`uo_instance`, `cockpit_ingenieur`, `dashboard_pilote`, ...)
- `FILE_ID` (obligatoire) : identifiant unique de ce fichier
- Toute autre ligne `clé: valeur` est une **métadonnée libre**, stockée et
  utilisable comme filtre par `LIST ... WHERE`. Deux métadonnées sont utilisées
  en pratique aujourd'hui : `ingenieur:` (nom de l'ingénieur propriétaire d'une
  UO ou d'un cockpit) et `pilote_id:` (identifiant du responsable métier,
  utilisé pour que le dashboard retrouve les cockpits de son équipe)

## DEF — capturer ou calculer une donnée

Toute donnée manipulée doit d'abord être nommée avec `$` via `DEF`.

**Lire un tableau Excel structuré :**
```
DEF $act = GET_TABLE(Activites, tbl_activites)
```

**Lire une cellule/plage nommée :**
```
DEF $seuil = GET_CELL(Parametres, seuil_rouge)
```

**Calculer une valeur** — deux syntaxes strictement équivalentes :
```
DEF $avancement = COMPUTE(MEAN_WEIGHTED($actifs.avancement, $actifs.poids))
DEF $avancement = MEAN_WEIGHTED($actifs.avancement, $actifs.poids)
```
Le moteur réécrit automatiquement la forme courte en `COMPUTE(...)` — les deux
sont interchangeables, la forme explicite `COMPUTE(...)` est celle utilisée
dans `creer_uo.py`, la forme courte n'apparaît pas encore dans le code des
générateurs mais fonctionne de façon identique.

### Fonctions disponibles dans `COMPUTE`

| Fonction | Rôle | Exemple réel |
|---|---|---|
| `FILTER(table, condition)` | Sous-ensemble de lignes | `FILTER($act, applicable = "OUI")` |
| `MEAN_WEIGHTED(val, poids)` | Moyenne pondérée | `MEAN_WEIGHTED($actifs.avancement, $actifs.poids)` |
| `SUM(col)` | Somme | `SUM($actifs.heures_consommees)` |
| `COUNT(col)` | Nombre de lignes non-null | — |
| `COUNT_IF(col, "val")` | Comptage conditionnel | `COUNT_IF($oil.statut, "OUVERT")` |
| `AVG(col)` | Moyenne simple | — |
| `MIN(col)` / `MAX(col)` | Min / max | — |
| `DIV(a, b)` | Division (0 si b=0) | — |
| `IF(cond, si_vrai, si_faux)` | Condition | — |
| `IF_NULL(val, defaut)` | Valeur par défaut si null | — |

**Disponibles dans le moteur mais aucun exemple généré en production
aujourd'hui** — utilisables sans escalade, juste pas encore éprouvées côté
Excel réel :

| Fonction | Rôle |
|---|---|
| `SWITCH_RANGE(val, [lo,hi]:"label", ...)` | Statut selon une plage numérique (ex. feu tricolore) |
| `GROUP_BY(table, col, sortie=AGG(col), ...)` | Agrégation par catégorie |
| `SORT(table, col, ASC\|DESC)` | Tri |
| `TOP_N(table, n, col, ASC\|DESC)` | Les n premières lignes après tri |

Opérateurs de condition dans `FILTER`/`IF` : `=`, `!=`, `>`, `>=`, `<`, `<=`,
`AND`, `OR`.

Toute formule non reconnue fait échouer le sync avec une erreur explicite
(`Fonction COMPUTE inconnue`) — pas de mode dégradé silencieux.

## COL — décrire une colonne

```
COL $mes_uos.avancement : WRITE=engineer
COL $mes_uos.heures_realisees : WRITE=engineer
```

`WRITE=` indique qui a le droit d'écrire dans la colonne (convention, pas
vérifié par le moteur à l'écriture Excel elle-même). Ajouter le mot `KEY`
dans les attributs marque la colonne comme clé primaire de la table.

## PUSH — envoyer vers le store central

```
PUSH $variable -> nom.namespace.cle
```

Exemples réels :
```
PUSH $mes_uos -> cockpit.Alice_Dubois.mes_uos
PUSH $synthese -> dashboard.USR004.synthese
PUSH $avancement -> uo.L09U1-CFL2400-CLIM.avancement
```

Optionnel : `ONLY_IF condition` pour ne pousser que si une condition est
vraie (sinon le PUSH est ignoré silencieusement, pas d'erreur). Existe dans le
moteur, pas encore utilisé dans les générateurs actuels.

**Conventions de nommage des clés observées dans le code** :
- `uo.<code>.<champ>` — données d'une UO (ex. `uo.L09U1-CFL2400-CLIM.avancement`)
- `cockpit.<nom_ingenieur>.mes_uos` — table consolidée d'un cockpit
- `dashboard.<pilote_id>.synthese` — table consolidée d'un dashboard

## PULL — recevoir depuis le store central

```
PULL nom.namespace.cle -> FILL_TABLE(feuille, tableau) MODE=mode [KEY=col]
PULL nom.namespace.cle -> UPDATE_CELLS(feuille, tableau, KEY=col, COLS=col1;col2)
```

Modes de `FILL_TABLE` : `READ_ONLY`, `OVERWRITE`, `APPEND_NEW` (ajoute
seulement les nouvelles lignes selon `KEY`), `UPDATE` (met à jour les lignes
existantes sans en ajouter). Existe dans le moteur, aucun exemple généré vu
dans les fichiers actuels (les générateurs actuels utilisent plutôt
`LIST`/`COLLECT` pour agréger, voir plus bas).

## BIND — écrire une valeur dans une plage nommée

```
BIND $avancement -> KPI.kpi_avancement
```

Écrit la valeur au moment du sync (pas une formule Excel live — un nouveau
sync est nécessaire pour rafraîchir).

## VALIDATE — règles de qualité

```
VALIDATE $actifs.avancement : RANGE(0, 100)
VALIDATE $act.applicable : IN("OUI", "NON")
```

Règles disponibles : `NOT_NULL`, `NOT_EMPTY`, `POSITIVE`, `NON_NEGATIVE`,
`UNIQUE`, `RANGE(min,max)`, `IN("a","b",...)`, `MIN_LENGTH(n)`,
`MAX_LENGTH(n)`, `MATCHES(regex)`, `MIN(n)`, `MAX(n)`. Ajouter
`SEVERITY=warning` pour ne pas bloquer le sync (par défaut : bloquant).

## LIST — découvrir des fichiers

```
LIST mes_cockpits TYPE=cockpit_ingenieur WHERE pilote_id=USR004
```

Forme utilisée en production. Cherche tous les fichiers dont `FILE_TYPE`
correspond et dont une métadonnée libre de l'en-tête satisfait la condition
`WHERE`. Il existe aussi une forme `LIST nom FROM TABLE mon_tableau` (lit une
liste de fichiers depuis un tableau Excel local) — existe dans le moteur, pas
utilisée en pratique aujourd'hui.

## COLLECT — agréger des tables cross-fichiers

```
COLLECT tbl_mes_uos FROM mes_cockpits INTO tbl_vue_synthese
```

Aspire la table `tbl_mes_uos` depuis chaque fichier de la liste `mes_cockpits`
et consolide le tout dans `tbl_vue_synthese`. Options disponibles dans le
moteur (`WHERE condition`, `COLS=[...]`, `WITH champ1, champ2` pour ajouter des
colonnes issues des métadonnées du fichier source) — aucun exemple généré vu
en production, mais utilisables directement si besoin.

## EXTENDS — héritage de template (disponible, pas utilisé en pratique)

```
EXTENDS uo_generique
```

Hérite des `DEF`/`COL`/`PUSH`/`PULL`/`VALIDATE` d'un template `.mxl` dans
`config/templates/`. Complètement implémenté et testé côté moteur ; les
générateurs actuels (`creer_uo.py`, cockpit, dashboard) écrivent leurs
instructions en clair plutôt que d'hériter d'un template — mécanisme
disponible si un besoin de factorisation apparaît.

## NOTIFY — alerte (disponible, pas utilisé en pratique)

```
NOTIFY log "ALERTE : statut rouge" IF $statut_global = "ROUGE"
```

Canaux : `log`, `email` (`TO adresse`), `webhook` (`TO url`). Implémenté côté
moteur (SMTP et HTTP réels), aucun générateur actuel ne l'utilise.

## Trois `_Manifeste` réels et complets, tels que générés par le code actuel

### Cockpit ingénieur (`src/generators/cockpit_ingenieur_generator.py`)

```
MANIFESTE_V=1

FILE_TYPE: cockpit_ingenieur
FILE_ID: Cockpit_Alice_Dubois
ingenieur: Alice Dubois
pilote_id: USR004

DEF $mes_uos = GET_TABLE(Mes UOs, tbl_mes_uos)
COL $mes_uos.avancement : WRITE=engineer
COL $mes_uos.heures_realisees : WRITE=engineer

PUSH $mes_uos -> cockpit.Alice_Dubois.mes_uos
```

### Dashboard métier (`src/generators/dashboard_metier_generator.py`)

```
MANIFESTE_V=1

FILE_TYPE: dashboard_pilote
FILE_ID: Dashboard_USR004
pilote_id: USR004

LIST mes_cockpits TYPE=cockpit_ingenieur WHERE pilote_id=USR004
COLLECT tbl_mes_uos FROM mes_cockpits INTO tbl_vue_synthese

DEF $synthese = GET_TABLE(Vue Synthèse, tbl_vue_synthese)
PUSH $synthese -> dashboard.USR004.synthese
```

### UO instance (`projet_TrainSystem/creer_uo.py`)

```
MANIFESTE_V=1

FILE_TYPE: uo_instance
FILE_ID: L09U1-CFL2400-CLIM
ingenieur: Alice Dubois

DEF $act = GET_TABLE(Activites, tbl_activites)
DEF $oil = GET_TABLE(OIL, tbl_oil)
DEF $liv = GET_TABLE(Livrables, tbl_livrables)

DEF $actifs = COMPUTE(FILTER($act, applicable = "OUI"))
DEF $po_ouv = COMPUTE(FILTER($oil, statut = "OUVERT"))

DEF $avancement = COMPUTE(MEAN_WEIGHTED($actifs.avancement, $actifs.poids))
DEF $h_conso = COMPUTE(SUM($actifs.heures_consommees))
DEF $po_ouverts = COMPUTE(COUNT_IF($oil.statut, "OUVERT"))
DEF $po_fermes = COMPUTE(COUNT_IF($oil.statut, "CLOS"))
DEF $po_critiques = COMPUTE(COUNT_IF($po_ouv.criticite, "HAUTE"))

VALIDATE $actifs.avancement : RANGE(0, 100)
VALIDATE $act.applicable : IN("OUI", "NON")
VALIDATE $oil.statut : IN("OUVERT", "CLOS")

BIND $avancement -> KPI.kpi_avancement
BIND $h_conso -> KPI.kpi_h_conso
BIND $po_ouverts -> KPI.kpi_po_ouverts

PUSH $avancement -> uo.L09U1-CFL2400-CLIM.avancement
PUSH $h_conso -> uo.L09U1-CFL2400-CLIM.heures_consommees
PUSH $actifs -> uo.L09U1-CFL2400-CLIM.activites
```

## Erreurs courantes

| Erreur | Solution |
|---|---|
| `PUSH` sur une variable sans `DEF` préalable | Toujours `DEF` avant de `PUSH` |
| `COL` écrit avant le `DEF GET_TABLE` correspondant | `DEF` en premier, `COL` ensuite |
| Penser que `BIND` est une formule Excel live | `BIND` écrit une valeur figée au moment du sync — relancer un sync pour rafraîchir |
| Modifier un `_Manifeste` à la main en pensant que c'est versionné | Les `.xlsx` sont gitignorés — seul le générateur Python qui l'écrit est versionné |
| Utiliser une fonction `COMPUTE` inexistante | Le sync échoue avec une erreur explicite listant la fonction inconnue — vérifier la liste ci-dessus avant d'escalader |
```

- [ ] **Step 2 : Vérifier le fichier**

Relire chaque exemple et le comparer mot pour mot avec les extraits confirmés
par lecture du code source (`cockpit_ingenieur_generator.py:315-367`,
`dashboard_metier_generator.py:355-398`, `creer_uo.py:449-507`). Vérifier que
la distinction "éprouvé en production" vs "disponible mais pas exploité" est
appliquée de façon cohérente à `EXTENDS`, `NOTIFY`, `SWITCH_RANGE`,
`GROUP_BY`, `SORT`, `TOP_N`, `PULL` partout dans le document.

- [ ] **Step 3 : Commit**

```bash
git add docs/onboarding-excel/02-Langage-MXL.md
git commit -m "docs(onboarding): 02-Langage-MXL - reference verifiee contre le code source"
```

---

### Task 5 : `docs/onboarding-excel/03-Design-System-XD.md`

**Files:**
- Create: `docs/onboarding-excel/03-Design-System-XD.md`

- [ ] **Step 1 : Écrire le fichier**

```markdown
# Design System Excel — `src/xl_design.py`

Tout ce qui touche à la mise en forme d'un fichier Excel généré par ExoSync
passe par la classe `XD` de `src/xl_design.py`. **Règle non négociable : zéro
style Excel écrit en inline dans un générateur.** Si tu as besoin d'un style
qui n'existe pas encore dans `XD`, ajoute-le dans `xl_design.py`, ne le code
pas en dur dans ton générateur.

## Pourquoi cette règle

Elle garantit que tout fichier généré (UO, cockpit, dashboard, et toute vue
future comme le dashboard client) partage la même palette, la même
typographie, les mêmes composants visuels — sans effort de synchronisation
manuelle entre générateurs.

## Import et usage

```python
from src.xl_design import XD

XD.banner(ws, "activites", "Mes Activités", subtitle="UO L09U1-CFL2400-CLIM", n_cols=10)
XD.table_header(ws, row=5, headers=["ID", "Libellé", "Heures"], key="activites", col_start=2)
XD.named_table(ws, display_name="tbl_activites", ref="B5:D20", key="activites")
```

## Les 11 familles d'onglets

Chaque onglet appartient à une famille (`key`), qui détermine sa couleur
(3 tons dérivés : bannière foncée, en-tête moyen, lignes alternées claires) et
son glyphe. Familles disponibles : `general`, `dashboard`, `description`,
`planning`, `donnees_entree`, `activites`, `livrables`, `oil`, `kpi`, `orga`,
`manifeste`. Voir `XD.SHEETS` dans `src/xl_design.py` pour la palette exacte
de chaque famille — ne pas inventer de nouvelle couleur ailleurs, choisir la
famille la plus proche du sens de l'onglet.

## Composants principaux

| Méthode | Rôle |
|---|---|
| `XD.banner(ws, key, title, subtitle, se, n_cols)` | Bannière 1 ligne en tête d'onglet, pose aussi `tabColor` |
| `XD.table_header(ws, row, headers, key, col_start)` | En-tête de tableau coloré au ton moyen de la famille |
| `XD.data_row(ws, row, i, col_start, col_end, key)` | Ligne de données avec alternance de fond |
| `XD.named_table(ws, display_name, ref, key)` | Table Excel nommée (nécessaire pour `GET_TABLE`/`COLLECT` côté MXL) avec en-tête coloré |
| `XD.statut_cf(ws, rng)` | Mise en forme conditionnelle badges de statut (TERMINEE, EN_COURS, A_FAIRE...) |
| `XD.criticite_cf(ws, rng)` | Mise en forme conditionnelle badges de criticité OIL |
| `XD.traffic_light(ws, row, col, value, thresholds)` | Cellule feu rouge/ambre/vert selon un seuil |
| `XD.section_box(ws, title, r1, c1, r2, c2, key)` | Bande de titre + cadre, pour les onglets clé-valeur |
| `XD.health_spine(ws, key, header_row, row_start, row_end, status_col, spine_col)` | Colonne santé (voir ci-dessous) |

## La colonne "spine santé"

Composant contextuel pour les onglets à lignes "vivantes" (Activités du côté
UO, Mes UOs du côté cockpit, Synthèse du côté dashboard) : une colonne A fine
(largeur 2.5), colorée en direct par mise en forme conditionnelle selon une
colonne statut, qui donne un repère visuel immédiat sans avoir à lire le
détail de chaque ligne.

**Règle absolue : la colonne spine ne doit jamais faire partie du `ref` d'une
table Excel nommée.** Si elle l'est, `GET_TABLE`/`COLLECT` côté MXL lit une
colonne en trop et casse l'agrégation. Conséquence pratique : quand une table
a une spine, le `ref` de la table nommée commence en colonne B, pas en
colonne A.

```python
# Bon : la table commence en B, la spine (col A) est hors du ref
XD.named_table(ws, "tbl_activites", "B4:J20", "activites")
XD.health_spine(ws, "activites", header_row=4, row_start=5, row_end=20, status_col=7)
```

## Règles universelles à respecter sur tout nouvel onglet

1. Bannière en tête (`XD.banner`)
2. Police Segoe UI partout (gérée automatiquement par `XD.fnt`)
3. En-tête de tableau au ton de la famille (`XD.table_header`)
4. Lignes alternées (`XD.data_row` ou géré par le style de table nommée)
5. Jaune de saisie `XD.INPUT` (`FFF2CC`) là où l'utilisateur doit saisir
6. Couleur d'onglet (`tabColor`) posée automatiquement par `XD.banner`
```

- [ ] **Step 2 : Vérifier le fichier**

Comparer chaque signature de méthode citée avec `src/xl_design.py` réel
(classe `XD`, méthodes `banner`, `table_header`, `data_row`, `named_table`,
`statut_cf`, `criticite_cf`, `traffic_light`, `section_box`, `health_spine`) —
confirmer que les noms de paramètres correspondent exactement à la signature
actuelle du fichier.

- [ ] **Step 3 : Commit**

```bash
git add docs/onboarding-excel/03-Design-System-XD.md
git commit -m "docs(onboarding): 03-Design-System-XD - usage de la classe XD"
```

---

### Task 6 : `docs/onboarding-excel/04-Territoire-Et-Conventions.md`

**Files:**
- Create: `docs/onboarding-excel/04-Territoire-Et-Conventions.md`

- [ ] **Step 1 : Écrire le fichier**

```markdown
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

## Workflow Git

- Une branche dédiée par sujet de travail (`git checkout -b feature/<nom>`),
  jamais de commit direct sur `master`
- Commits fréquents et atomiques — un commit = un changement cohérent et
  testé, pas un gros commit en fin de journée
- Pull request vers `master` quand une brique de travail est terminée et que
  `pytest` passe intégralement
- Message de commit au format `type(scope): description` (ex.
  `feat(cockpit): ajoute colonne priorite`, `fix(dashboard): corrige filtre pilote_id`)
```

- [ ] **Step 2 : Vérifier le fichier**

Confirmer la cohérence avec les conventions déjà observées dans l'historique
Git du dépôt (`git log --oneline` montre le format `type(scope): description`
utilisé de façon constante) et avec la frontière validée dans la spec.

- [ ] **Step 3 : Commit**

```bash
git add docs/onboarding-excel/04-Territoire-Et-Conventions.md
git commit -m "docs(onboarding): 04-Territoire-Et-Conventions - frontiere, TDD, workflow Git"
```

---

### Task 7 : `docs/onboarding-excel/05-Escalade.md`

**Files:**
- Create: `docs/onboarding-excel/05-Escalade.md`

- [ ] **Step 1 : Écrire le fichier**

```markdown
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
```

- [ ] **Step 2 : Vérifier le fichier**

Confirmer que les exemples de décision (colonne priorité, couleur badges,
`GROUP_BY`, fonction médiane, nouvelle instruction, `models.py`) sont
cohérents avec la frontière définie dans `04-Territoire-Et-Conventions.md` et
avec la liste "disponible mais pas exploité" de `02-Langage-MXL.md`.

- [ ] **Step 3 : Commit**

```bash
git add docs/onboarding-excel/05-Escalade.md
git commit -m "docs(onboarding): 05-Escalade - quand et comment escalader"
```

---

### Task 8 : Vérification finale

**Files:** aucun (vérification uniquement)

- [ ] **Step 1 : Lister les fichiers livrés**

```bash
cd "C:\Users\fabie\Documents\JLC\Python\SysEng"
ls ESCALADES.md docs/onboarding-excel/
```

Attendu : `ESCALADES.md` à la racine, et dans `docs/onboarding-excel/` :
`00-DEMARRAGE.md`, `01-Vue-Ensemble-Projet.md`, `02-Langage-MXL.md`,
`03-Design-System-XD.md`, `04-Territoire-Et-Conventions.md`,
`05-Escalade.md`.

- [ ] **Step 2 : Vérifier les critères de succès de la spec**

Relire `docs/superpowers/specs/2026-07-02-onboarding-collegue-excel-design.md`
section 7 et cocher chaque critère :
- [ ] Les 6 fichiers de `docs/onboarding-excel/` existent et sont commités
- [ ] `ESCALADES.md` existe à la racine avec format + exemple
- [ ] Lecture de `00-DEMARRAGE.md` seul permet de comprendre où aller ensuite
      selon la tâche
- [ ] Aucune modification du code existant (vérifier avec `git diff master --stat`
      que seuls des fichiers `.md` apparaissent)

- [ ] **Step 3 : Vérifier qu'aucun test n'est cassé**

```bash
pytest --tb=short -q
```

Attendu : toujours 415 passed (ce projet ne touche aucun fichier `.py`, donc
aucune régression possible, mais on confirme).

- [ ] **Step 4 : Commit final si des ajustements ont eu lieu pendant la vérification, sinon rien à committer**

```bash
git status
git log --oneline master..feature/onboarding-excel
```

Attendu : 8 commits (Task 1 à 7 + éventuels ajustements), branche prête pour
review/merge.
```

- [ ] **Step 5 : Rapport final**

Résumer dans le chat : liste des fichiers créés, résultat des critères de
succès, résultat de `pytest`, et proposer la suite (merge vers `master`, ou
attente de relecture par l'utilisateur).
