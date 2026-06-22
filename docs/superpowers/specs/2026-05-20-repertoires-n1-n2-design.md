# Design — Clarification des répertoires N1/N2

**Date :** 2026-05-20
**Statut :** approuvé
**Scope :** web/schema_app + web/registry_app

---

## Contexte

L'application ExoSync Studio est divisée en deux apps web :
- **N1 — Schema Designer** (port 8001) : crée et édite les gabarits (schémas d'écosystème)
- **N2 — Registry Populator** (port 8000) : gère la vie des Affaires (population, Posts, sync)

La gestion des répertoires était confuse : 4 chemins définis à deux endroits, avec
`affaire_dir` inutilisé, et `workspace_dir` / `gabarits_dir` éditables dans les deux apps
sans propriétaire clair.

---

## Modèle cible

### 3 chemins, 3 rôles distincts

| Champ | Fichier de stockage | Propriétaire | Scope |
|-------|---------------------|-------------|-------|
| `gabarits_dir` | `.exosync_workspace.json` | **N1** édite — N2 lit | Global |
| `workspace_dir` | `.exosync_workspace.json` | **N2** édite — N1 ignore | Global |
| `posts_dir` | `directories.json` dans l'Affaire | **N2** édite | Par Affaire |

`affaire_dir` est supprimé. La racine d'une Affaire est son dossier de config actif
(déjà connu via `.active_registry_ecosystem`), pas besoin de le stocker séparément.

### Principe : chemins racines configurés une fois, sous-dossiers créés automatiquement

L'utilisateur configure `gabarits_dir` et `workspace_dir` une seule fois.
Lors de la création d'un gabarit ou d'une affaire, il fournit uniquement un **nom court**
— le système crée le sous-dossier automatiquement :

```
gabarits_dir/
  ├── SysEng/          ← créé via nom "SysEng"
  └── SysEng-v2/

workspace_dir/
  ├── SNCF-2026/       ← créé via nom "SNCF-2026"
  └── RATP-2025/
```

Aucun champ "chemin complet" dans les modales de création.

---

## Flux utilisateur

### Flux 1 — Premier démarrage N1 (architecte)

1. Ouvrir N1 → vue "Workspace"
2. Définir `gabarits_dir` (chemin absolu existant)
3. Créer un gabarit : saisir un nom court → `gabarits_dir/<nom>/` créé automatiquement
4. N1 s'active sur ce dossier — l'architecte édite le schéma

### Flux 2 — Premier démarrage N2 (gestionnaire)

1. Ouvrir N2 → vue "Configuration globale"
2. Définir `workspace_dir` (chemin absolu existant)
3. Voir `gabarits_dir` en lecture seule (défini par N1)
4. Vue "Répertoires de l'Affaire" → définir `posts_dir` pour l'Affaire active

### Flux 3 — Créer une Affaire depuis un gabarit (N2)

1. Vue "Gabarits" → liste les sous-dossiers de `gabarits_dir` (scan auto)
2. Choisir un gabarit → saisir un nom court pour la nouvelle Affaire
3. Système crée `workspace_dir/<nom>/` avec les fichiers N1 copiés + N2 vides
4. N2 s'active automatiquement sur la nouvelle Affaire

---

## Changements API

### `web/registry_app/api/directories.py`

**Avant :**
```python
class DirsConfig(BaseModel):
    affaire_dir: str = ""
    posts_dir: str = ""
```

**Après :**
```python
class DirsConfig(BaseModel):
    posts_dir: str = ""
```

Validation : chemin absolu, peut ne pas encore exister (OneDrive, réseau).

### `web/registry_app/api/gabarits.py`

**`NewGabaritRequest` — avant :**
```python
class NewGabaritRequest(BaseModel):
    name: str
    path: str           # chemin explicite requis
    description: str = ""
```

**Après :**
```python
class NewGabaritRequest(BaseModel):
    name: str           # nom court → gabarits_dir/<name>/ créé automatiquement
    description: str = ""
```

Erreur 422 si `gabarits_dir` non configuré dans le workspace.

**`CloneRequest`** : inchangé — `dest_path` déjà optionnel, génère `workspace_dir/<name>/`
si absent. Erreur 422 si `workspace_dir` non configuré et `dest_path` vide.

### `web/schema_app/api/workspace.py`

Endpoint `PUT /api/workspace` : n'accepte plus `workspace_dir` (champ ignoré ou retiré
du modèle). N1 ne gère que `gabarits_dir`.

**Modèle N1 :**
```python
class WorkspaceConfig(BaseModel):
    gabarits_dir: str = ""
    # workspace_dir retiré
```

---

## Changements UI

### N1 — Vue Workspace

- Conserver uniquement le champ `gabarits_dir`
- Supprimer le champ `workspace_dir`
- Libellé : "Répertoire des Gabarits"

### N2 — Vue "Configuration globale" (ex "Workspace global")

- `workspace_dir` : éditable — "Répertoire des Affaires"
- `gabarits_dir` : affiché en lecture seule — "Répertoire des Gabarits (défini dans N1)"
- Valeur rechargée à chaque ouverture de la vue (appel API normal, pas de bouton dédié)

### N2 — Vue "Répertoires de l'Affaire" (ex "Répertoires")

- Supprimer le champ `affaire_dir`
- Conserver uniquement `posts_dir` — "Répertoire des Posts (Excel)"
- Label : chemin absolu, OneDrive accepté

### Modales de création (N1 + N2)

- **Créer gabarit** : champ `nom` + champ `description` (optionnel). Plus de champ `path`.
- **Créer Affaire / Cloner** : champ `nom` + champ `description` (optionnel). Plus de champ `dest_path`.

---

## Règles de validation

| Champ | Règle |
|-------|-------|
| `gabarits_dir` | Absolu, doit exister |
| `workspace_dir` | Absolu, doit exister |
| `posts_dir` | Absolu, peut ne pas exister (OneDrive) |
| Nom gabarit/affaire | Non vide, pas de `/\:*?"<>|`, sous-dossier cible ne doit pas exister |

---

## Documentation à mettre à jour

- **Obsidian vault** : `20-Projets/ExoSync/Conversations/CONV-08-Web-N2-Registry-Populator.md`
  — ajouter une section "Évolutions (2026-05-20c)" avec le nouveau modèle de répertoires
  (3 chemins, propriétaires, flux de création automatique)

---

## Fichiers modifiés

| Fichier | Modification |
|---------|-------------|
| `web/registry_app/api/directories.py` | Suppression `affaire_dir` |
| `web/registry_app/services/config_service.py` | `load_dirs` / `save_dirs` sans `affaire_dir` |
| `web/registry_app/api/gabarits.py` | `NewGabaritRequest` sans `path`, auto-génération |
| `web/schema_app/api/workspace.py` | Modèle sans `workspace_dir` |
| `web/registry_app/static/app.js` | UI N2 : champs supprimés, modales simplifiées |
| `web/schema_app/static/app.js` | UI N1 : champ `workspace_dir` supprimé |

---

## Migration des données existantes

Les `directories.json` existants peuvent contenir `affaire_dir`. Pydantic ignore les
champs inconnus à la lecture — la valeur est silencieusement abandonnée lors du prochain
`PUT /api/directories`. Aucune migration de données manuelle nécessaire.

---

## Ce qui ne change pas

- `.exosync_workspace.json` reste le fichier partagé N1+N2 (structure inchangée)
- `web/workspace_service.py` : code inchangé, les deux champs restent dans le fichier
- `CloneRequest.dest_path` : reste optionnel (rétrocompatibilité), UI n'en fait plus usage
- La logique de scan automatique de `gabarits_dir` (sous-dossiers) reste identique
