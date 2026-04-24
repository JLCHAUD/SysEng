
# Système de Gestion des UO — État du Projet
**Date :** 2026-04-20  
**Répertoire :** `C:\Users\fabie\Documents\JLC\Python\SysEng`

---

## Vision du projet

Plateforme Python de pilotage technique des **Unités d'Œuvre (UO)** d'ingénierie système ferroviaire (RATP MI20, RER-NG, etc.).

Chaque UO est une tranche de travail (type × système × projet × charge) portée par un ingénieur. La plateforme :
- génère des fichiers Excel structurés par UO
- synchronise les données entre fichiers via des **feuilles _Passerelle** (méta-langage)
- centralise les indicateurs dans un **store JSON** (source de vérité)
- produit des **cockpits par ingénieur** et une **consolidation centrale**
- s'administre via **Creator.xlsx** (console IT Manager)

---

## Architecture

```
SysEng/
├── main.py                          # CLI click (15 commandes)
├── config/
│   ├── uo_instances.json            # 5 UO instances
│   ├── types_uo.json                # 4 types UO (SF, SS, ST, DJ)
│   ├── systemes.json                # référentiels systèmes
│   ├── projets.json                 # MI20_RATP, RER_NG
│   ├── acteurs.json                 # 8 profils utilisateurs
│   └── registre.json                # 14 fichiers déclarés
├── src/
│   ├── models.py                    # dataclasses + enums (v2)
│   ├── config_loader.py             # chargement config JSON
│   ├── store.py                     # store central output/store.json
│   ├── ecosystem.py                 # catalogue vivant output/ecosystem.json
│   ├── parser.py                    # parser méta-langage _Passerelle → AST
│   ├── passerelle.py                # moteur d'exécution passerelle (v1, à migrer)
│   ├── sync.py                      # synchronisation + rapports
│   ├── styles.py                    # styles Excel partagés
│   └── generators/
│       ├── uo_generator.py          # génération fichiers UO
│       ├── cockpit_generator.py     # génération cockpits ingénieurs
│       ├── consolidation_generator.py
│       └── creator_generator.py    # Creator.xlsx (console admin)
├── output/
│   ├── UOs/                         # 5 fichiers UO générés
│   ├── cockpits/                    # 3 cockpits générés
│   ├── consolidation.xlsx
│   ├── store.json                   # store central (vide pour l'instant)
│   ├── ecosystem.json               # 3 tables, 3 variables découvertes
│   └── rapports/                    # rapports sync JSON
└── referentiels/                    # (à générer — non encore créés)
```

---

## Couche données — Modèles (src/models.py)

### Enums
| Enum | Valeurs |
|------|---------|
| `StatutUO` | BROUILLON, EN_COURS, CLOTUREE, ARCHIVEE, VISIBLE |
| `StatutActivite` | EN_COURS, ANNULEE, TERMINEE, NON_APPLICABLE |
| `StatutLivrable` | A_FAIRE, EN_COURS, LIVRE, VALIDE |
| `Role` | ingenieur_sys, pilote_tech, engagement_mgr, it_manager, donneur_ordre, pilote_tech_client, resp_projet_client, expert_client, fournisseur |
| `TypeFiltre` | uo, ingenieur, projet, systeme, projet+systeme, ALL |
| `NiveauAcces` | read/write, read, read_filtered, read_summary, admin |
| `DirectionPasserelle` | pull, push |
| `TypePasserelle` | CELL, CELL_NUM, CELL_DATE, CELL_PCT, TABLE_COL, TABLE_ROW, TABLE_FULL, COMPUTED, REF |

### Dataclasses principales
- `UOInstance` : id, uo_type_id, system_id, project_id, engineer_name, total_hours, start_date, end_date, **statut, degrade, degrade_note** + refs résolues
- `UOType` : id, name, activities[], deliverables[]
- `Activity` : id, name, default_hours, statut, heures_realisees, allocated_hours
- `Deliverable` : id, name, due_date, date_reelle, status
- `ProfilActeur` : id, nom, role, filtre_type, filtre_valeur, acces
- `EntreeRegistre` : id, type_fichier, chemin, synchro_periodicite, derniere_synchro, statut_dernier_synchro
- `ReglePasserelle` / `Passerelle` : modèles v1 (utilisés par passerelle.py)

---

## Données de configuration

### 5 UO Instances
| ID | Type | Système | Projet | Ingénieur | Statut |
|----|------|---------|--------|-----------|--------|
| UO-001 | spec_fonctionnelle | climatisation | MI20_RATP | Alice Dubois | EN_COURS |
| UO-002 | spec_systeme | frein | MI20_RATP | Alice Dubois | BROUILLON |
| UO-003 | spec_technique | portes | MI20_RATP | Bruno Lecomte | EN_COURS |
| UO-004 | dossier_justification | eclairage | RER_NG | Bruno Lecomte | BROUILLON |
| UO-005 | spec_fonctionnelle | traction | RER_NG | Camille Vidal | EN_COURS |

### 8 Profils acteurs (config/acteurs.json)
USR001–USR008, couvrant tous les rôles (ingénieurs, pilotes, IT manager, clients)

### Registre (14 fichiers)
- 3 référentiels UO (SF, SS, ST)
- 2 référentiels projet (MI20_RATP, RER_NG)
- 5 instances UO
- 3 cockpits ingénieurs
- 1 consolidation

---

## Méta-langage _Passerelle

La feuille `_Passerelle` (col A = instruction, col B = ancre) pilote toutes les synchronisations.

### Syntaxe complète
```
# Commentaire
FILE_TYPE: uo_instance
FILE_ID:   UO-001
VERSION:   1

DEF $var = GET_CELL(Feuille, plage_nommee)
DEF $var = GET_TABLE(Feuille, NomTableau)
DEF $var = COMPUTE(formule_python)

COL $table.col : KEY [HEADER="Label Excel"]
COL $table.col : WRITE=engineer|creation|uo_generique  [HEADER="..."] [LOCKED]

BIND $var -> Feuille.plage_nommee

PUSH $var -> global.variable.name

PULL global.var -> FILL_TABLE(Feuille, Tableau)  MODE=READ_ONLY|APPEND_NEW|UPDATE|OVERWRITE  [KEY=col]
PULL global.var -> UPDATE_CELLS(Feuille, Tableau, KEY=col, COLS=c1;c2)
```

### Règles de résolution des plages nommées (col B)
1. Plage nommée Excel native → utilisation directe
2. Référence cellule `Feuille.C2` → lecture de la valeur dans cette cellule
3. Col B vide → Python crée la plage nommée ET écrit la référence en col B (auto-documenté)
4. Tableau nommé Excel natif → résolution via `ws.tables`

### Propriété WRITE= (ownership des colonnes)
- `creation` : écrit uniquement à la création du fichier, jamais écrasé ensuite
- `engineer` : modifiable par l'ingénieur, jamais écrasé par sync
- `uo_generique` : vient du référentiel UO, sync peut écraser
- `it_manager` : réservé IT manager

---

## Parser (src/parser.py) — NOUVEAU, v2

### AST produit
```python
PasserelleAST:
    header:  FileHeader           # FILE_TYPE, FILE_ID, VERSION
    defs:    List[DefNode]        # GET_CELL | GET_TABLE | COMPUTE
    cols:    List[ColNode]        # colonnes avec KEY/WRITE/HEADER/LOCKED
    binds:   List[BindNode]       # BIND $var -> Sheet.range
    pushes:  List[PushNode]       # PUSH $var -> global.name
    pulls:   List[PullNode]       # PULL global -> FILL_TABLE|UPDATE_CELLS
    errors:  List[ParseError]     # lignes non parsées
    _defs_index: Dict[str, DefNode]
    _cols_index: Dict[str, List[ColNode]]
```

### API publique
```python
parse_file(filepath: Path) -> Optional[PasserelleAST]
parse_sheet(ws) -> PasserelleAST
parse_lines(lines: List[Tuple[str, str]]) -> PasserelleAST
enrich_ecosystem(ast: PasserelleAST) -> Tuple[int, int]  # (nb_tables, nb_vars)
ast_summary(ast: PasserelleAST) -> str
```

### Résultat sur UO-001 (testé)
- 0 erreurs
- 6 DEF (3 GET_TABLE + 3 COMPUTE)
- 20 COL (8 activites + 5 livrables + 7 po)
- 3 BIND Dashboard
- 6 PUSH store
- 2 PULL import

---

## Ecosystem Schema (src/ecosystem.py)

Catalogue vivant persisté dans `output/ecosystem.json`.

```python
ColumnSchema:   name, col_type, header, write, description
TableSchema:    id, source_file_id, source_sheet, table_name, columns{}, discovered_from, last_seen
VariableSchema: id, var_type, source_file_id, formula, discovered_from, last_seen
EcosystemSchema: version, tables{}, variables{}
```

### État actuel (après parse UO-001)
- 3 tables : `uo.activites` (8 cols), `uo.livrables` (5 cols), `uo.points_ouverts` (7 cols)
- 3 variables : `uo.avancement_global`, `uo.heures_realisees`, `uo.nb_po_ouverts`

---

## Générateurs de fichiers Excel (src/generators/)

### uo_generator.py → 8 feuilles par UO
1. **Organisation Projet** — tableau acteurs (TabActeurs)
2. **Livrables** — tableau livrables (TabLivrables) + mise en forme conditionnelle
3. **Planning** — planning des livrables avec formule écart
4. **Activités** — tableau activités (TabActivites) + footer pondéré
5. **REX** — retour d'expérience
6. **Points Ouverts** — tableau PO (TabPO) + MFC statut
7. **Dashboard** — KPIs + barre de progression + avancement formule
8. **_Passerelle** — méta-langage (nouveau format v2)

Tableaux Excel nommés créés nativement (TabActivites, TabLivrables, TabPO, TabActeurs).

### cockpit_generator.py
Cockpit par ingénieur — agrégation des UO dont il est responsable.

### consolidation_generator.py
Vue consolidée de toutes les UO.

### creator_generator.py → Creator.xlsx (10 feuilles)
Console d'administration IT Manager :
- Acteurs, Registre, Types UO, Systèmes, Projets (config)
- Catalogue Tables, Catalogue Variables (référence)
- Ecosystem Summary (catalogue)
- Créer Fichier (formulaire de création)
- Guide (aide)

---

## Store central (src/store.py)

`output/store.json` — source de vérité inter-fichiers.

```python
store.get("uo.activites")          # lecture
store.set("uo.avancement_global", 0.75)
store.set_many({"k1": v1, "k2": v2})
store.get_all()
store.snapshot()
```

Clés utilisées : `uo.activites`, `uo.livrables`, `uo.points_ouverts`, `uo.avancement_global`, `uo.heures_realisees`, `uo.nb_po_ouverts`, `projet.acteurs`, `referentiel.uo_types`...

---

## Moteur de synchronisation (src/sync.py)

Ordre de traitement : referentiel_uo → referentiel_projet → uo_instance → cockpit → pilote → consolidation → client

- Détection de verrouillage (fichier ouvert dans Excel) via `open(path, "r+b")`
- Rapport JSON par session dans `output/rapports/`
- Mise à jour du registre (dernière_synchro, statut)

**Limitation actuelle :** `sync.py` appelle `passerelle.py` (moteur v1), pas encore le nouveau parser.

---

## CLI (main.py) — 15 commandes

| Commande | Description |
|----------|-------------|
| `generate-uo --uo-id UO-001` | Génère 1 fichier UO |
| `generate-all-uo` | Génère tous les fichiers UO |
| `generate-cockpit --engineer "Alice Dubois"` | Cockpit 1 ingénieur |
| `generate-all-cockpits` | Tous les cockpits |
| `generate-consolidation` | Fichier de consolidation |
| `generate-all` | Tout générer en une commande |
| `sync [--id X] [--type Y] [--force]` | Synchronisation |
| `sync-uo UO-001` | Sync 1 UO |
| `create-creator` | Génère Creator.xlsx |
| `list-registre` | Affiche le registre |
| `onboard <chemin>` | Audit + template passerelle |
| `parse-file <chemin> [--enrich]` | Parse _Passerelle, affiche AST |
| `enrich-ecosystem [--dir]` | Scan xlsx → ecosystem |

---

## Ce qui est FAIT ✓

- [x] Modèles v2 complets (UOInstance, ProfilActeur, EntreeRegistre, Passerelle...)
- [x] Config JSON complète (5 UO, 4 types, 2 projets, 2 systèmes, 8 acteurs, 14 registre)
- [x] Générateur UO (8 feuilles, tableaux nommés, MFC, Dashboard avec formules)
- [x] Générateur cockpit ingénieur
- [x] Générateur consolidation
- [x] Creator.xlsx (console admin 10 feuilles)
- [x] Store central JSON
- [x] Ecosystem schema (catalogue vivant)
- [x] **Parser méta-langage v2** (AST complet, 0 erreurs sur UO-001)
- [x] `_Passerelle` générée en nouveau méta-langage dans les UO
- [x] CLI `parse-file` et `enrich-ecosystem`
- [x] Moteur sync v1 (passerelle.py — fonctionne sur l'ancien format)

## Ce qui RESTE à faire

### Priorité 1 — Connecter parser → moteur de sync
- Écrire `src/executor.py` : exécute un `PasserelleAST` (résolution des DEF, PULL, PUSH, BIND, COMPUTE)
- Remplacer l'appel à `executer_passerelle()` dans `sync.py` par le nouvel exécuteur
- Gérer les 4 modes PULL : READ_ONLY, APPEND_NEW, UPDATE, OVERWRITE
- Implémenter les formules COMPUTE : MEAN_WEIGHTED, SUM, COUNT_IF

### Priorité 2 — Fichiers référentiels
- Générer `referentiels/UO_Generique_spec_fonctionnelle.xlsx` (etc.) avec leur `_Passerelle`
- Générer `referentiels/Projet_MI20_RATP.xlsx`, `Projet_RER_NG.xlsx`
- Ces fichiers alimentent le store lors du premier bootstrap

### Priorité 3 — BIND (Dashboard autonome)
- Implémenter l'écriture de formules Excel dans `executor.py` via BIND
- Ex : `BIND $avancement_global -> Dashboard.avancement_global` → écrit `=Activités!F12` dans la plage nommée Dashboard

### Priorité 4 — Résolution des plages nommées (règle 3)
- Quand col B est vide pour un BIND/PUSH : créer la plage nommée Excel ET écrire la référence en col B

### Priorité 5 — Creator reader
- `src/creator_reader.py` : lit Creator.xlsx → met à jour ecosystem.json et registre.json
- Déclenchement sur flag "Créer Fichier"

### Priorité 6 — Bootstrap sync
- Premier run pour peupler store.json depuis les référentiels
- Validation end-to-end : référentiel → store → UO instance

---

## Environnement technique

- Python 3.11+, Windows 11
- openpyxl (génération/lecture Excel)
- click (CLI)
- dataclasses, enum, json (stdlib)
- Pas de base de données — tout en fichiers (JSON + Excel)
- Pas de dépendance réseau — 100% local
