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
