# Design — Édition du Schéma N1 dans une Affaire N2 active

**Date :** 2026-05-21  
**Statut :** Approuvé  
**Contexte :** ExoSync / SysEng — séparation N1 (Schema Designer, port 8001) / N2 (Registry Populator, port 8000)

---

## Problème

Le workflow actuel impose de quitter N2 pour aller dans N1 dès qu'une Classe doit être créée ou modifiée dans une Affaire active. Cela rompt le flux de travail et ne correspond pas à la réalité terrain : les Classes évoluent souvent en même temps que les Posts.

**Contraintes à respecter :**
- Ne pas fusionner N1 et N2 (séparation architecturale conservée)
- Permettre à la fois l'ajout de nouvelles Classes ET la modification de Classes existantes
- Quand une Classe est modifiée, les Posts existants de cette Classe doivent être signalés comme périmés
- L'onglet Schéma est intégré directement dans N2 (pas d'iframe N1)

---

## Approche retenue : Schema Router Factory

Refactoriser les routers N1 pour qu'ils exposent une factory `make_router(cfg)`, puis les monter dans N2 sous `/api/schema/`. Le comportement est identique ; seul le chemin des fichiers de config change via `cfg`.

---

## Section 1 — Backend : Schema Router Factory

### `web/schema_config.py` (nouveau)

```python
from dataclasses import dataclass
from typing import Callable

@dataclass
class SchemaConfigService:
    load_file_types: Callable
    save_file_types: Callable
    load_tables:     Callable
    save_tables:     Callable
    load_relations:  Callable
    save_relations:  Callable
    load_namespaces: Callable
    save_namespaces: Callable
    # ... autres load/save selon les routers existants
```

### Routers N1 refactorisés

Chaque router schema (`classes.py`, `relations.py`, `namespaces.py`, `templates.py`, `fonctions.py`) expose une factory :

```python
# Avant
router = APIRouter()
@router.get("/")
def list_classes():
    return load_file_types()

# Après
def make_router(cfg: SchemaConfigService) -> APIRouter:
    router = APIRouter()
    @router.get("/")
    def list_classes():
        return cfg.load_file_types()
    return router
```

### Montage dans N1 (`web/schema_app/main.py`) — inchangé fonctionnellement

```python
from web.schema_config import SchemaConfigService
from web.schema_app.services.config_service import load_file_types, save_file_types, ...

cfg = SchemaConfigService(
    load_file_types=load_file_types,
    save_file_types=save_file_types,
    ...
)
app.include_router(classes.make_router(cfg), prefix="/api/classes")
```

### Montage dans N2 (`web/registry_app/main.py`) — nouveau

```python
from web.schema_config import SchemaConfigService
from web.registry_app.services.config_service import load_file_types, save_file_types, ...

schema_cfg = SchemaConfigService(
    load_file_types=load_file_types,
    save_file_types=save_file_types,
    ...
)
app.include_router(classes.make_router(schema_cfg), prefix="/api/schema/classes")
app.include_router(relations.make_router(schema_cfg), prefix="/api/schema/relations")
# etc.
```

**Pattern cohérent avec :** `web/mxl_service.py`, `web/workspace_service.py`

---

## Section 2 — Versionnage du schéma

### `file_types.yaml` — champ `schema_version`

Chaque entrée de Classe reçoit `schema_version: int` (commence à 1, incrémenté automatiquement à chaque `PATCH /api/schema/classes/{class_id}` ou équivalent).

```yaml
uo_instance:
  schema_version: 3
  required_sheets: [...]
  ...
```

### `registre.json` — champ `schema_version` par Post

Chaque Post stocke la version du schéma au moment de sa dernière génération Excel :

```json
{ "id": "UO-001", "type_fichier": "uo_instance", "schema_version": 2 }
```

- Posé par `GET /api/xlsx/{file_id}` à la génération
- Mis à jour au même endpoint à chaque regénération
- Absent = considéré version 0 (périmé par défaut)

### API Registre — flag `schema_outdated`

`GET /api/registre` enrichit chaque Post :

```json
{
  "id": "UO-001",
  "type_fichier": "uo_instance",
  "schema_version": 2,
  "schema_outdated": true
}
```

`schema_outdated = (post.schema_version || 0) < class.schema_version`

---

## Section 3 — UI N2 : onglet Schéma

### Navigation

Nouvel onglet **"Schéma"** dans la barre principale N2, visible uniquement quand une Affaire est chargée.

### Structure

Navigation latérale gauche + panneau principal (même pattern que N1) :

```
[ Registre | Schéma | Tissage | ... ]
               │
    ┌──────────┼──────────────────────────────┐
    │ Classes  │  Liste des Classes           │
    │ Tables   │  + bouton "Nouvelle Classe"  │
    │ Relations│  Clic → éditeur inline       │
    │ Fonctions│  (formulaires identiques N1) │
    │ Espaces  │                              │
    │ Templates│                              │
    └──────────┴──────────────────────────────┘
```

### Implémentation UI

**Option retenue : composants Vue dans `app.js` N2** (pas d'iframe N1).

Les vues Schema en N2 appellent `/api/schema/...` en local. Volume modeste (liste + formulaire par entité), cohérence visuelle garantie, pas de dépendance au port N1.

### Comportement à l'enregistrement d'une Classe

1. `PATCH /api/schema/classes/{class_id}` → `schema_version` incrémenté dans `file_types.yaml`
2. Toast : *"Classe mise à jour — X Posts concernés sont maintenant signalés comme périmés"*
3. Badges ⚠ apparaissent dans l'onglet Registre (données réactives Vue)

---

## Section 4 — Bouton "Regénérer"

### Présentation

Le bouton **"⬇ Excel"** existant devient contextuel :

| État du Post | Bouton | Style |
|---|---|---|
| À jour | `⬇ Excel` | neutre (inchangé) |
| Périmé | `⬇ Excel ⚠` | orange, tooltip explicatif |

Tooltip : *"Schéma mis à jour — ce téléchargement applique la version courante"*

Pas de bouton séparé — même endpoint `GET /api/xlsx/{file_id}`.

### Flux à l'enregistrement

```
Modifier Classe → schema_version++ → badges ⚠ sur Posts concernés
→ Utilisateur clique ⬇ Excel ⚠ → nouveau gabarit Excel téléchargé
→ post.schema_version mis à jour dans registre.json → badge disparaît
```

### Hors scope (v1)

- Pas de migration automatique des données existantes dans l'Excel utilisateur
- Pas de "regénérer tous les Posts périmés en un clic"
- Pas de notification push / webhook
- Pas de diff de colonnes entre versions

---

## Résumé des fichiers impactés

| Fichier | Action |
|---|---|
| `web/schema_config.py` | Créer — dataclass `SchemaConfigService` |
| `web/schema_app/api/classes.py` | Refactoriser → `make_router(cfg)` |
| `web/schema_app/api/relations.py` | Refactoriser → `make_router(cfg)` |
| `web/schema_app/api/namespaces.py` | Refactoriser → `make_router(cfg)` |
| `web/schema_app/api/templates.py` | Refactoriser → `make_router(cfg)` |
| `web/schema_app/api/fonctions.py` | Refactoriser → `make_router(cfg)` |
| `web/schema_app/main.py` | Adapter — construire cfg, passer à make_router |
| `web/registry_app/main.py` | Étendre — monter routers schema sous `/api/schema/` |
| `web/registry_app/api/xlsx_generator.py` | Étendre — poser schema_version au téléchargement |
| `web/registry_app/api/registre.py` | Étendre — calculer schema_outdated |
| `web/registry_app/static/app.js` | Étendre — onglet Schéma + badges ⚠ |
| `file_types.yaml` (Affaires) | Données — ajouter schema_version par Classe |
| `registre.json` (Affaires) | Données — ajouter schema_version par Post |
