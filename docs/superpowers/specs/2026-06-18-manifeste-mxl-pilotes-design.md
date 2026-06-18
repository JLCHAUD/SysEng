# Design — Refonte `_Manifeste` MXL + Pilotes n:m

**Date :** 2026-06-18
**Projet :** ExoSync / SysEng
**Statut :** Approuvé

---

## Contexte

Les générateurs `cockpit_ingenieur_generator.py` et `dashboard_metier_generator.py` produisaient
un `_Manifeste` au format tabulaire multi-colonnes (format `passerelle.py` legacy).
`sync.py` essaie d'abord de parser la feuille avec `parser.py` (MXL mono-colonne) ; si la feuille
existe mais contient le format tabulaire, le parser produit des erreurs "Mot-clé inconnu" et
l'exécuteur ne fait rien. Le cycle push/pull est silencieusement cassé.

Ce design corrige le format et introduit une relation n:m entre UOs et pilotes.

---

## Décisions d'architecture

### 1. Format `_Manifeste` — MXL mono-colonne

Toutes les feuilles `_Manifeste` générées utilisent le format lu par `parser.py` :

| Colonne | Rôle | Exemple |
|---------|------|---------|
| A | Instruction MXL | `PUSH $activites -> uo.activites` |
| B | Ancre (cellule cible pour BIND/PULL) | `Synthèse.G5` |
| C | Commentaire français (ignoré par le moteur) | `Export vers store central` |

La colonne B reste réservée à l'ancre MXL. La colonne C est le seul endroit pour les commentaires.

### 2. Tables, jamais de cellules individuelles

Les générateurs utilisent `GET_TABLE` + `COL` pour référencer les données.
Interdit : `GET_CELL(Mes UOs, F6)`, `GET_CELL(Mes UOs, F7)` — fragile, ne scale pas.

### 3. Relation n:m UO ↔ Pilotes

Une UO peut avoir plusieurs pilotes selon leur rôle (métier TS, métier projet, sécurité…).
Chaque pilote est déclaré dans le `_Manifeste` de l'UO en tant que métadonnée libre.
Le parser stocke ces champs dans `ast.header.manifest_metadata` automatiquement.

### 4. Auto-découverte dans le dashboard

Le dashboard ne liste plus les UOs en dur.
Il utilise `LIST DYNAMIC WHERE pilote_<role>=<id>` pour découvrir ses UOs à chaque sync.

---

## Modèle de données

### `UOInstance` — nouveau champ `pilotes`

```python
class UOInstance:
    # champs existants…
    pilotes: Dict[str, str] = {}
    # ex: {"metier_ts": "USR004", "metier_projet": "USR007"}
```

### `config/uo_instances.json` — exemple

```json
{
  "id": "UO-001",
  "engineer_name": "Alice Dubois",
  "charge_allouee": 32,
  "pilotes": {
    "metier_ts":     "USR004",
    "metier_projet": "USR007"
  }
}
```

Un `{}` vide est valide — UO non encore assignée à un pilote.

---

## Format `_Manifeste` UO instance

```
Ligne 1, col A : MANIFESTE_V=1
Ligne 2         : (vide — skippée par parse_sheet)
Ligne 3, col A : FILE_TYPE: uo_instance          col C : Type de fichier ExoSync
Ligne 4, col A : FILE_ID: UO-001                 col C : Identifiant unique de l'UO
Ligne 5, col A : pilote_metier_ts: USR004        col C : Pilote métier TS responsable
Ligne 6, col A : pilote_metier_projet: USR007    col C : Pilote métier Projet
Ligne 7         : (vide)
Ligne 8, col A : DEF $activites = GET_TABLE(Activités, tbl_activites)      col C : Table des activités
Ligne 9, col A : COL $activites.avancement : WRITE=engineer                col C : % avancement saisi ingénieur
Ligne 10, col A: COL $activites.heures_realisees : WRITE=engineer          col C : Heures réalisées saisies
Ligne 11        : (vide)
Ligne 12, col A: PUSH $activites -> uo.activites                           col C : Export vers store central
```

---

## Format `_Manifeste` cockpit ingénieur

```
Ligne 1 : MANIFESTE_V=1
Ligne 3 : FILE_TYPE: cockpit_ingenieur
Ligne 4 : FILE_ID: Cockpit_Alice_Dubois
Ligne 5 : ingenieur: Alice Dubois                   col C : Propriétaire du cockpit
Ligne 7 : DEF $mes_uos = GET_TABLE(Mes UOs, tbl_mes_uos)
Ligne 8 : COL $mes_uos.avancement : WRITE=engineer  col C : % avancement — zone de saisie
Ligne 9 : COL $mes_uos.heures_realisees : WRITE=engineer   col C : H réalisées — zone de saisie
Ligne 11: PUSH $mes_uos -> cockpit.mes_uos          col C : Remonte les saisies vers le store
```

Le tableau "Mes UOs" dans la feuille Excel doit être un tableau nommé (`tbl_mes_uos`)
pour que `GET_TABLE` le retrouve via le mécanisme `_trouver_tableau_nomme` de `passerelle.py`.

---

## Format `_Manifeste` dashboard pilote

```
Ligne 1 : MANIFESTE_V=1
Ligne 3 : FILE_TYPE: dashboard_pilote
Ligne 4 : FILE_ID: Dashboard_USR004
Ligne 5 : pilote_id: USR004                         col C : Identifiant du pilote propriétaire
Ligne 7 : LIST mes_uos TYPE=uo_instance WHERE pilote_metier_ts=USR004   col C : Découverte auto des UOs
Ligne 8 : COLLECT Activites FROM mes_uos INTO vue_synthese              col C : Agrégation toutes UOs
```

Le dashboard **ne contient pas de liste d'UOs en dur**.
Quand une nouvelle UO est créée avec `pilote_metier_ts: USR004`, elle apparaît
automatiquement au prochain sync sans modifier le dashboard.

---

## Fichiers à modifier ou créer

| Fichier | Action | Raison |
|---------|--------|--------|
| `src/models.py` | Modifier | Ajouter `pilotes: Dict[str, str]` sur `UOInstance` |
| `config/uo_instances.json` | Modifier | Ajouter le champ `pilotes` sur les 5 UOs existantes |
| `src/generators/cockpit_ingenieur_generator.py` | Modifier | Réécrire `_sheet_manifeste()` en MXL mono-colonne + nommer le tableau "Mes UOs" |
| `src/generators/dashboard_metier_generator.py` | Modifier | Réécrire `_sheet_manifeste_dashboard()` en MXL mono-colonne avec LIST DYNAMIC |
| `src/executor.py` | Vérifier / Modifier | S'assurer que LIST DYNAMIC + COLLECT sont exécutés |
| `tests/test_cockpit_ingenieur.py` | Modifier | Adapter les assertions sur le format `_Manifeste` |
| `tests/test_dashboard_metier.py` | Modifier | Adapter les assertions + ajouter test auto-découverte |

---

## Risque : exécution LIST + COLLECT dans `executor.py`

`parser.py` parse `LIST DYNAMIC` et `COLLECT` mais leur exécution dans `executor.py`
doit être vérifiée. Si non implémentée, c'est une tâche du plan d'implémentation.
Ce spec couvre le **format** ; l'exécution est la couche suivante.

---

## Non couvert par ce spec (scope suivant)

- Branchement des nouveaux cockpits dans le CLI (`main.py generate-cockpit`)
- Mise à jour de `config/registre.json` pour les nouveaux fichiers
- Interface de création/édition des UOs (admin)
