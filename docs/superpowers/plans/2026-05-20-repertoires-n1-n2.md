# Clarification répertoires N1/N2 — Plan d'implémentation

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Simplifier et clarifier la gestion des répertoires : 3 chemins avec propriétaires distincts, sous-dossiers créés automatiquement depuis un nom court, `affaire_dir` supprimé.

**Architecture:** N1 possède `gabarits_dir`, N2 possède `workspace_dir` + `posts_dir` par Affaire. Le fichier `.exosync_workspace.json` reste partagé — chaque app fait une mise à jour partielle (ne touche que son champ). Les sous-dossiers gabarits et affaires sont créés automatiquement : l'utilisateur donne un nom court, le système construit le chemin.

**Tech Stack:** Python 3, FastAPI, Pydantic v2, Vue 3 CDN — pas de nouveaux packages.

---

## Fichiers modifiés

| Fichier | Changement |
|---------|-----------|
| `web/registry_app/api/directories.py` | Supprimer `affaire_dir` de `DirsConfig` et validations |
| `web/registry_app/services/config_service.py` | Valeur par défaut `load_dirs()` sans `affaire_dir` |
| `web/registry_app/api/gabarits.py` | `NewGabaritRequest` et `SaveAsGabaritRequest` sans champ path ; auto-génération |
| `web/registry_app/api/workspace.py` | GET retourne les 2 champs ; PUT ne met à jour que `workspace_dir` |
| `web/schema_app/api/workspace.py` | `WorkspaceConfig` sans `workspace_dir` ; PUT partiel |
| `web/registry_app/static/app.js` | `ViewDirectories` sans `affaire_dir` ; `ViewWorkspaceGlobal` avec `gabarits_dir` RO ; `ViewGabarits` sans champs path |
| `web/schema_app/static/app.js` | `ViewWorkspace` sans `workspace_dir` |

---

## Task 1 — Backend N2 : supprimer `affaire_dir` de `directories.py`

**Fichiers :**
- Modifier : `web/registry_app/api/directories.py`

- [ ] **Remplacer le modèle et les routes**

```python
# web/registry_app/api/directories.py
"""Gestion du répertoire Posts — N2 Registry Populator.

Un seul répertoire par Affaire : posts_dir (chemin absolu des fichiers Excel Posts).
Stockage : directories.json dans le dossier config actif N2.
"""
from pathlib import Path
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from web.registry_app.services.config_service import load_dirs, save_dirs

router = APIRouter()


class DirsConfig(BaseModel):
    posts_dir: str = ""    # chemin absolu — peut pointer vers OneDrive ou tout autre emplacement


def get_posts_base_path() -> Path | None:
    """Retourne le chemin absolu du répertoire Posts.
    Retourne None si posts_dir n'est pas défini.
    """
    d = load_dirs()
    posts = d.get("posts_dir", "")
    if not posts:
        return None
    return Path(posts)


# ── Routes ─────────────────────────────────────────────────────────────────────

@router.get("", response_model=DirsConfig)
def get_dirs():
    return DirsConfig(**load_dirs())


@router.put("", response_model=DirsConfig)
def update_dirs(body: DirsConfig):
    if body.posts_dir:
        p = Path(body.posts_dir)
        if not p.is_absolute():
            raise HTTPException(400, "Répertoire Posts : le chemin doit être absolu")
        if p.exists() and not p.is_dir():
            raise HTTPException(400, "Posts : ce chemin n'est pas un répertoire")
        # chemin non existant accepté (OneDrive non encore synchronisé)

    save_dirs(body.model_dump())
    return body


@router.delete("", status_code=204)
def reset_dirs():
    """Remet les répertoires à zéro."""
    save_dirs({"posts_dir": ""})
```

- [ ] **Mettre à jour `load_dirs()` dans `config_service.py`**

Ligne à modifier dans `web/registry_app/services/config_service.py` :

```python
def load_dirs() -> dict:
    p = _p("directories.json")
    if not p.exists():
        return {"posts_dir": ""}          # était : {"affaire_dir": "", "posts_dir": "posts"}
    with open(p, encoding="utf-8") as f:
        return json.load(f)
```

- [ ] **Vérifier manuellement**

```bash
cd C:/Users/fabie/Documents/JLC/Python/SysEng
python run_web.py
# Dans un autre terminal :
curl http://localhost:8000/api/directories
# Attendu : {"posts_dir": ""}  — pas de affaire_dir
```

- [ ] **Commit**

```bash
git add web/registry_app/api/directories.py web/registry_app/services/config_service.py
git commit -m "feat(N2): supprimer affaire_dir — répertoires simplifiés à posts_dir seul"
```

---

## Task 2 — Backend N2 : `gabarits/new` avec chemin automatique

**Fichiers :**
- Modifier : `web/registry_app/api/gabarits.py`

- [ ] **Remplacer `NewGabaritRequest` et `new_gabarit()`**

Dans `web/registry_app/api/gabarits.py`, remplacer :

```python
class NewGabaritRequest(BaseModel):
    name: str
    path: str           # absolu ou relatif au projet
    description: str = ""
```

par :

```python
_INVALID_NAME_CHARS = set(r'\/:*?"<>|')

class NewGabaritRequest(BaseModel):
    name: str           # nom court → gabarits_dir/<name>/ créé automatiquement
    description: str = ""
```

Puis remplacer la fonction `new_gabarit()` entière :

```python
@router.post("/new", response_model=GabaritInfo, status_code=201)
def new_gabarit(body: NewGabaritRequest):
    """Crée un nouveau Gabarit vide sous gabarits_dir/<name>/."""
    gabarits_dir = get_gabarits_dir()
    if not gabarits_dir:
        raise HTTPException(422, "gabarits_dir non configuré — définissez-le dans N1 (Workspace)")

    if not body.name or any(c in _INVALID_NAME_CHARS for c in body.name):
        raise HTTPException(422, f"Nom invalide : '{body.name}' (caractères interdits: \\ / : * ? \" < > |)")

    path = gabarits_dir / body.name

    if path.exists() and any(path.iterdir()):
        raise HTTPException(409, f"Le dossier '{path}' existe déjà et n'est pas vide")

    path.mkdir(parents=True, exist_ok=True)

    with open(path / "file_types.yaml", "w", encoding="utf-8") as f:
        yaml.dump({"file_types": {}}, f, allow_unicode=True, default_flow_style=False)

    for fname, init in [
        ("schema_relations.json", {"version": "1", "relations": []}),
        ("functions.json",        {"version": "1", "functions": []}),
        ("templates.json",        {"version": "1", "templates": []}),
        ("namespaces.json",       {"version": "1", "namespaces": []}),
    ]:
        with open(path / fname, "w", encoding="utf-8") as f:
            json.dump(init, f, ensure_ascii=False, indent=2)

    return GabaritInfo(name=body.name, path=str(path), valid=True, class_count=0,
                       description=body.description)
```

Note : on ne touche plus `.gabarits.json` — le scan auto de `gabarits_dir` est la seule source.

- [ ] **Vérifier**

```bash
curl -X POST http://localhost:8000/api/gabarits/new \
  -H "Content-Type: application/json" \
  -d '{"name": "TestGab", "description": "test auto-path"}'
# Attendu 201 avec path = gabarits_dir/TestGab
# Vérifier que le dossier existe physiquement
```

- [ ] **Cas d'erreur : gabarits_dir non configuré**

```bash
# (si gabarits_dir vide dans workspace)
curl -X POST http://localhost:8000/api/gabarits/new \
  -H "Content-Type: application/json" \
  -d '{"name": "TestGab"}'
# Attendu : 422 "gabarits_dir non configuré"
```

- [ ] **Commit**

```bash
git add web/registry_app/api/gabarits.py
git commit -m "feat(N2): nouveau gabarit via nom seul — chemin auto sous gabarits_dir"
```

---

## Task 3 — Backend N2 : `gabarits/from-current` avec chemin automatique

**Fichiers :**
- Modifier : `web/registry_app/api/gabarits.py`

- [ ] **Remplacer `SaveAsGabaritRequest` et `save_current_as_gabarit()`**

Remplacer :

```python
class SaveAsGabaritRequest(BaseModel):
    name: str
    dest_path: str      # où copier les fichiers N1
    description: str = ""
```

par :

```python
class SaveAsGabaritRequest(BaseModel):
    name: str           # nom court → gabarits_dir/<name>/ créé automatiquement
    description: str = ""
```

Puis remplacer `save_current_as_gabarit()` :

```python
@router.post("/from-current", response_model=GabaritInfo, status_code=201)
def save_current_as_gabarit(body: SaveAsGabaritRequest):
    """Exporte le schéma N1 de l'Affaire active comme nouveau Gabarit."""
    source = get_active_config()
    if not _is_valid_gabarit(source):
        raise HTTPException(422, "L'affaire active n'a pas de file_types.yaml valide")

    gabarits_dir = get_gabarits_dir()
    if not gabarits_dir:
        raise HTTPException(422, "gabarits_dir non configuré — définissez-le dans N1 (Workspace)")

    if not body.name or any(c in _INVALID_NAME_CHARS for c in body.name):
        raise HTTPException(422, f"Nom invalide : '{body.name}'")

    dest = gabarits_dir / body.name

    if dest.exists() and any(dest.iterdir()):
        raise HTTPException(409, f"Le dossier '{dest}' existe déjà et n'est pas vide")

    dest.mkdir(parents=True, exist_ok=True)

    for fname in _N1_FILES:
        src_file = source / fname
        if src_file.exists():
            shutil.copy2(src_file, dest / fname)

    return GabaritInfo(name=body.name, path=str(dest), valid=True,
                       class_count=_count_classes(dest), description=body.description)
```

- [ ] **Vérifier le clone (déjà fonctionnel, confirmer)**

`clone_gabarit()` utilise déjà `workspace_dir/name` quand `dest_path=""`.
Vérifier que c'est bien le cas — pas de modification nécessaire sur la logique.

```bash
curl -X POST http://localhost:8000/api/gabarits/clone \
  -H "Content-Type: application/json" \
  -d '{"source_path": "<path_gabarit>", "name": "Affaire-Test", "activate": false}'
# Attendu : 201 avec path = workspace_dir/Affaire-Test
```

- [ ] **Commit**

```bash
git add web/registry_app/api/gabarits.py
git commit -m "feat(N2): export gabarit via nom seul — chemin auto sous gabarits_dir"
```

---

## Task 4 — Backend N1 : `workspace.py` sans `workspace_dir` (mise à jour partielle)

**Fichiers :**
- Modifier : `web/schema_app/api/workspace.py`

- [ ] **Remplacer le fichier entier**

```python
"""API Workspace — N1 Schema Designer.

N1 est propriétaire de gabarits_dir uniquement.
workspace_dir est géré par N2 et préservé lors de la mise à jour.
"""
from pathlib import Path
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from web.workspace_service import load_workspace, save_workspace

router = APIRouter()


class WorkspaceConfig(BaseModel):
    gabarits_dir: str = ""


@router.get("", response_model=WorkspaceConfig)
def get_workspace():
    data = load_workspace()
    return WorkspaceConfig(gabarits_dir=data.get("gabarits_dir", ""))


@router.put("", response_model=WorkspaceConfig)
def update_workspace(body: WorkspaceConfig):
    if body.gabarits_dir:
        p = Path(body.gabarits_dir)
        if not p.is_absolute():
            raise HTTPException(400, "Gabarits : le chemin doit être absolu")
        if p.exists() and not p.is_dir():
            raise HTTPException(400, "Gabarits : ce chemin n'est pas un répertoire")

    # Mise à jour partielle — préserver workspace_dir géré par N2
    current = load_workspace()
    current["gabarits_dir"] = body.gabarits_dir
    save_workspace(current)
    return body


@router.delete("", status_code=204)
def reset_workspace():
    # Préserver workspace_dir lors du reset N1
    current = load_workspace()
    current["gabarits_dir"] = ""
    save_workspace(current)
```

- [ ] **Vérifier**

```bash
# Dans l'app N1 (port 8001)
curl http://localhost:8001/api/workspace
# Attendu : {"gabarits_dir": "..."} — pas de workspace_dir

curl -X PUT http://localhost:8001/api/workspace \
  -H "Content-Type: application/json" \
  -d '{"gabarits_dir": "C:/Users/fabie/Documents/JLC/ProjEXOSync/Gabarits"}'
# Attendu : 200

# Vérifier que workspace_dir de N2 est préservé dans .exosync_workspace.json
cat "C:/Users/fabie/Documents/JLC/Python/SysEng/.exosync_workspace.json"
# Attendu : gabarits_dir mis à jour, workspace_dir inchangé
```

- [ ] **Commit**

```bash
git add web/schema_app/api/workspace.py
git commit -m "feat(N1): workspace réduit à gabarits_dir — mise à jour partielle du fichier partagé"
```

---

## Task 5 — Backend N2 : `workspace.py` mise à jour partielle de `workspace_dir`

**Fichiers :**
- Modifier : `web/registry_app/api/workspace.py`

- [ ] **Remplacer le fichier entier**

```python
"""API Workspace — N2 Registry Populator.

N2 est propriétaire de workspace_dir uniquement.
gabarits_dir est géré par N1 : affiché en lecture seule, préservé lors de la mise à jour.
"""
from pathlib import Path
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

from web.workspace_service import load_workspace, save_workspace

router = APIRouter()


class WorkspaceResponse(BaseModel):
    """Réponse GET : les deux champs pour affichage."""
    gabarits_dir: str = ""
    workspace_dir: str = ""


class WorkspaceDirUpdate(BaseModel):
    """Corps PUT : uniquement workspace_dir (propriété N2)."""
    workspace_dir: str = ""


@router.get("", response_model=WorkspaceResponse)
def get_workspace():
    data = load_workspace()
    return WorkspaceResponse(
        gabarits_dir=data.get("gabarits_dir", ""),
        workspace_dir=data.get("workspace_dir", ""),
    )


@router.put("", response_model=WorkspaceResponse)
def update_workspace(body: WorkspaceDirUpdate):
    if body.workspace_dir:
        p = Path(body.workspace_dir)
        if not p.is_absolute():
            raise HTTPException(400, "Workspace : le chemin doit être absolu")
        if p.exists() and not p.is_dir():
            raise HTTPException(400, "Workspace : ce chemin n'est pas un répertoire")

    # Mise à jour partielle — préserver gabarits_dir géré par N1
    current = load_workspace()
    current["workspace_dir"] = body.workspace_dir
    save_workspace(current)
    return WorkspaceResponse(
        gabarits_dir=current.get("gabarits_dir", ""),
        workspace_dir=body.workspace_dir,
    )


@router.delete("", status_code=204)
def reset_workspace():
    # Préserver gabarits_dir lors du reset N2
    current = load_workspace()
    current["workspace_dir"] = ""
    save_workspace(current)
```

- [ ] **Vérifier**

```bash
curl http://localhost:8000/api/workspace
# Attendu : {"gabarits_dir": "...", "workspace_dir": "..."}

curl -X PUT http://localhost:8000/api/workspace \
  -H "Content-Type: application/json" \
  -d '{"workspace_dir": "C:/Users/fabie/Documents/JLC/ProjEXOSync"}'
# Attendu : 200 avec les deux champs, gabarits_dir préservé
```

- [ ] **Commit**

```bash
git add web/registry_app/api/workspace.py
git commit -m "feat(N2): workspace réduit à workspace_dir — gabarits_dir préservé en lecture"
```

---

## Task 6 — UI N2 : `ViewDirectories` (supprimer `affaire_dir`)

**Fichiers :**
- Modifier : `web/registry_app/static/app.js` — remplacer `ViewDirectories`

- [ ] **Remplacer `ViewDirectories` (lignes 2221–2308)**

```javascript
// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Répertoires de l'Affaire — Posts
// ═══════════════════════════════════════════════════════════════════════════════
const ViewDirectories = {
  setup() {
    const form    = ref({ posts_dir: '' });
    const saved   = ref(false);
    const loading = ref(true);
    const err     = ref('');

    onMounted(async () => {
      try {
        const d = await GET('/api/directories');
        form.value = { posts_dir: d.posts_dir || '' };
      } catch(e) { toastErr(e); }
      loading.value = false;
    });

    const save = async () => {
      err.value = '';
      try {
        await PUT('/api/directories', form.value);
        saved.value = true;
        setTimeout(() => saved.value = false, 2000);
        toast('Répertoire Posts enregistré');
      } catch(e) { err.value = e.message; toastErr(e); }
    };

    const reset = async () => {
      if (!confirm('Remettre le répertoire Posts à zéro ?')) return;
      await DEL('/api/directories');
      form.value = { posts_dir: '' };
      toast('Répertoire Posts réinitialisé');
    };

    return { form, saved, loading, err, save, reset };
  },
  template: `
    <div v-if="loading" class="loading">Chargement…</div>
    <div v-else style="max-width:640px;">

      <div style="margin-bottom:20px;padding:14px 16px;background:var(--surface2);border-radius:8px;border-left:3px solid var(--accent);font-size:0.82rem;color:var(--text-dim);line-height:1.6;">
        Définissez le dossier <strong style="color:var(--text)">Posts</strong> qui contient les fichiers Excel
        de cette Affaire. Ce chemin est utilisé par le moteur de synchronisation pour localiser les Posts.
      </div>

      <div class="form-group">
        <label class="form-label">Répertoire Posts (Excel)</label>
        <input v-model="form.posts_dir" class="form-control"
          placeholder="ex: C:\\Users\\user\\OneDrive\\ExoSync\\Posts" />
        <div style="font-size:0.75rem;color:var(--text-dim);margin-top:4px;">
          Chemin <strong>absolu</strong>. Peut pointer vers OneDrive, un réseau ou tout autre emplacement.
          Un chemin non encore existant est accepté (synchronisation OneDrive différée).
        </div>
      </div>

      <div v-if="err" style="margin-top:12px;padding:10px 14px;background:#7f1d1d;border-radius:6px;font-size:0.82rem;color:#fca5a5;">
        ⚠ {{ err }}
      </div>

      <div style="display:flex;gap:10px;margin-top:20px;">
        <button class="btn btn-primary" @click="save">
          {{ saved ? '✓ Enregistré' : '💾 Enregistrer' }}
        </button>
        <button class="btn btn-ghost" @click="reset" style="color:#ef4444;border-color:#ef4444;">
          ✕ Réinitialiser
        </button>
      </div>

      <div style="margin-top:24px;padding:12px 16px;background:var(--surface2);border-radius:8px;font-size:0.78rem;color:var(--text-dim);">
        <div style="font-weight:600;color:var(--text);margin-bottom:6px;">Impact sur le moteur Sync</div>
        <div>Les chemins relatifs des Posts dans le Registre sont résolus comme :</div>
        <div style="font-family:monospace;margin:6px 0;color:var(--accent);">
          {{ form.posts_dir || '[posts_dir]' }} \\ [chemin du Post]
        </div>
        <div>Les chemins absolus dans le Registre ne sont pas affectés.</div>
      </div>
    </div>
  `
};
```

- [ ] **Vérifier dans le navigateur**

Ouvrir `http://localhost:8000` → vue "Répertoires de l'Affaire" → un seul champ `posts_dir`, pas de `affaire_dir`. Sauvegarder et vérifier dans `directories.json` de l'Affaire active.

- [ ] **Commit**

```bash
git add web/registry_app/static/app.js
git commit -m "feat(N2 UI): ViewDirectories — affaire_dir supprimé, posts_dir seul"
```

---

## Task 7 — UI N2 : `ViewWorkspaceGlobal` (`gabarits_dir` lecture seule)

**Fichiers :**
- Modifier : `web/registry_app/static/app.js` — remplacer `ViewWorkspaceGlobal`

- [ ] **Remplacer `ViewWorkspaceGlobal` (lignes 2151–2215)**

```javascript
// ═══════════════════════════════════════════════════════════════════════════════
// VIEW: Configuration globale (workspace partagé N1 + N2)
// ═══════════════════════════════════════════════════════════════════════════════
const ViewWorkspaceGlobal = {
  setup() {
    const workspace_dir  = ref('');
    const gabarits_dir   = ref('');   // lecture seule — géré par N1
    const loading = ref(true);
    const saved   = ref(false);
    const err     = ref('');

    onMounted(async () => {
      try {
        const d = await GET('/api/workspace');
        workspace_dir.value = d.workspace_dir || '';
        gabarits_dir.value  = d.gabarits_dir  || '';
      } catch(e) { toastErr(e); }
      loading.value = false;
    });

    const save = async () => {
      err.value = '';
      try {
        await PUT('/api/workspace', { workspace_dir: workspace_dir.value });
        saved.value = true;
        setTimeout(() => saved.value = false, 2000);
        toast('Répertoire des Affaires enregistré');
      } catch(e) { err.value = e.message; toastErr(e); }
    };

    const reset = async () => {
      if (!confirm('Réinitialiser le répertoire des Affaires ?')) return;
      await DEL('/api/workspace');
      workspace_dir.value = '';
      toast('Répertoire des Affaires réinitialisé');
    };

    return { workspace_dir, gabarits_dir, loading, saved, err, save, reset };
  },
  template: `
    <div v-if="loading" class="loading">Chargement…</div>
    <div v-else style="max-width:640px;">
      <div style="margin-bottom:20px;padding:14px 16px;background:var(--surface2);border-radius:8px;border-left:3px solid #a5b4fc;font-size:0.82rem;color:var(--text-dim);line-height:1.6;">
        Paramètres globaux partagés entre N1 et N2. Stockés dans
        <code style="color:#a5b4fc;">.exosync_workspace.json</code> à la racine du projet.
      </div>

      <!-- workspace_dir — éditable N2 -->
      <div class="form-group">
        <label class="form-label">Répertoire des Affaires</label>
        <input v-model="workspace_dir" class="form-control"
          placeholder="ex: C:\\ExoSync\\Affaires" />
        <div style="font-size:0.75rem;color:var(--text-dim);margin-top:4px;">
          Dossier parent où les nouvelles Affaires sont créées automatiquement lors du clonage d'un gabarit.
        </div>
      </div>

      <!-- gabarits_dir — lecture seule, géré par N1 -->
      <div class="form-group" style="margin-top:16px;">
        <label class="form-label" style="color:var(--text-dim);">
          Répertoire des Gabarits
          <span style="font-size:0.72rem;background:var(--surface2);padding:2px 6px;border-radius:4px;margin-left:6px;color:#a5b4fc;">défini dans N1</span>
        </label>
        <input :value="gabarits_dir || '(non configuré — ouvrir N1 › Workspace)'" class="form-control"
          style="color:var(--text-dim);cursor:not-allowed;" disabled />
        <div style="font-size:0.75rem;color:var(--text-dim);margin-top:4px;">
          Lecture seule depuis N2. Pour modifier, ouvrir <strong>N1 Schema Designer</strong> › onglet Workspace.
        </div>
      </div>

      <div v-if="err" style="margin-top:12px;padding:10px 14px;background:#7f1d1d;border-radius:6px;font-size:0.82rem;color:#fca5a5;">⚠ {{ err }}</div>
      <div style="display:flex;gap:10px;margin-top:20px;">
        <button class="btn btn-primary" @click="save">{{ saved ? '✓ Enregistré' : '💾 Enregistrer' }}</button>
        <button class="btn btn-ghost" @click="reset" style="color:#ef4444;border-color:#ef4444;">✕ Réinitialiser</button>
      </div>
    </div>
  `
};
```

- [ ] **Vérifier dans le navigateur**

`http://localhost:8000` → vue "Configuration globale" → `workspace_dir` éditable, `gabarits_dir` grisé/disabled avec mention "défini dans N1". Modifier `workspace_dir`, vérifier que `gabarits_dir` n'est pas écrasé dans `.exosync_workspace.json`.

- [ ] **Commit**

```bash
git add web/registry_app/static/app.js
git commit -m "feat(N2 UI): ViewWorkspaceGlobal — workspace_dir éditable, gabarits_dir lecture seule"
```

---

## Task 8 — UI N2 : `ViewGabarits` (supprimer tous les champs path)

**Fichiers :**
- Modifier : `web/registry_app/static/app.js` — 4 endroits dans `ViewGabarits`

- [ ] **`newForm` : supprimer `path`**

```javascript
// Avant
const newForm = reactive({ name: '', path: '', description: '' });

// Après
const newForm = reactive({ name: '', description: '' });
```

- [ ] **`createNew()` : supprimer la garde sur `path`**

```javascript
// Avant
async function createNew() {
  if (!newForm.name || !newForm.path) return;

// Après
async function createNew() {
  if (!newForm.name) return;
```

- [ ] **`cloneForm` : supprimer `dest_path`**

```javascript
// Avant
const cloneForm = reactive({ dest_path: '', name: '', activate: true });

// Après
const cloneForm = reactive({ name: '', activate: true });
```

- [ ] **`cloneGabarit()` : supprimer la garde sur `dest_path` et le champ dans le POST**

```javascript
// Avant
async function cloneGabarit() {
  if (!cloneForm.dest_path || !cloneForm.name) return;
  loading.value = true;
  try {
    await POST('/api/gabarits/clone', {
      source_path: showClone.value.path,
      dest_path: cloneForm.dest_path,
      name: cloneForm.name,
      activate: cloneForm.activate,
    });
    Object.assign(cloneForm, { dest_path: '', name: '', activate: true });

// Après
async function cloneGabarit() {
  if (!cloneForm.name) return;
  loading.value = true;
  try {
    await POST('/api/gabarits/clone', {
      source_path: showClone.value.path,
      name: cloneForm.name,
      activate: cloneForm.activate,
    });
    Object.assign(cloneForm, { name: '', activate: true });
```

- [ ] **`exportForm` : supprimer `dest_path`**

```javascript
// Avant
const exportForm = reactive({ name: '', dest_path: '', description: '' });

// Après
const exportForm = reactive({ name: '', description: '' });
```

- [ ] **`exportCurrent()` : supprimer la garde sur `dest_path`**

```javascript
// Avant
async function exportCurrent() {
  if (!exportForm.name || !exportForm.dest_path) return;

// Après
async function exportCurrent() {
  if (!exportForm.name) return;
```

- [ ] **Template — formulaire "Nouveau Gabarit" : supprimer le champ Chemin**

Remplacer le bloc "Formulaire : Nouveau Gabarit" :

```html
<!-- Formulaire : Nouveau Gabarit -->
<div v-if="showNew" style="background:var(--surface2);border-radius:8px;padding:16px;margin-bottom:16px;">
  <div style="font-weight:600;margin-bottom:12px;">Nouveau Gabarit vide</div>
  <div style="font-size:0.78rem;color:var(--text-dim);margin-bottom:10px;">
    Un sous-dossier sera créé automatiquement sous <code>gabarits_dir/&lt;nom&gt;/</code>.
  </div>
  <div class="form-row" style="margin-bottom:10px;">
    <div class="form-group" style="margin:0;">
      <label>Nom *</label>
      <input v-model="newForm.name" placeholder="SysEng-v2" />
    </div>
    <div class="form-group" style="margin:0;">
      <label>Description</label>
      <input v-model="newForm.description" placeholder="Description optionnelle" />
    </div>
  </div>
  <div style="display:flex;gap:8px;justify-content:flex-end;">
    <button class="btn btn-ghost btn-sm" @click="showNew=false">Annuler</button>
    <button class="btn btn-primary btn-sm" :disabled="loading" @click="createNew">Créer</button>
  </div>
</div>
```

- [ ] **Template — modal "Cloner vers Affaire" : supprimer le champ Chemin de destination**

```html
<!-- Modal : Cloner vers une Affaire -->
<div v-if="showClone" class="modal-overlay" @click.self="showClone=null">
  <div class="modal">
    <div class="modal-title">Cloner "{{ showClone.name }}" → nouvelle Affaire</div>
    <div style="font-size:0.8rem;color:var(--text-dim);margin-bottom:16px;">
      Le schéma N1 sera copié dans <code>workspace_dir/&lt;nom&gt;/</code>. L'Affaire sera indépendante du Gabarit.
    </div>
    <div class="form-group">
      <label>Nom de l'Affaire *</label>
      <input v-model="cloneForm.name" placeholder="SNCF-2026" />
    </div>
    <div class="form-group" style="display:flex;align-items:center;gap:8px;">
      <input type="checkbox" id="cb-activate" v-model="cloneForm.activate" />
      <label for="cb-activate" style="margin:0;cursor:pointer;">Activer immédiatement cette Affaire</label>
    </div>
    <div style="display:flex;gap:8px;justify-content:flex-end;margin-top:16px;">
      <button class="btn btn-ghost" @click="showClone=null">Annuler</button>
      <button class="btn btn-primary" :disabled="loading" @click="cloneGabarit">⎘ Créer l'Affaire</button>
    </div>
  </div>
</div>
```

- [ ] **Template — formulaire "Exporter l'Affaire active" : supprimer le champ Destination**

```html
<!-- Formulaire : Exporter l'Affaire active comme Gabarit -->
<div v-if="showExport" style="background:var(--surface2);border-radius:8px;padding:16px;margin-bottom:16px;">
  <div style="font-weight:600;margin-bottom:12px;">Exporter le schéma N1 de l'Affaire active</div>
  <div style="font-size:0.78rem;color:var(--text-dim);margin-bottom:10px;">
    Un sous-dossier sera créé automatiquement sous <code>gabarits_dir/&lt;nom&gt;/</code>.
  </div>
  <div class="form-row" style="margin-bottom:10px;">
    <div class="form-group" style="margin:0;">
      <label>Nom du Gabarit *</label>
      <input v-model="exportForm.name" placeholder="SysEng-export" />
    </div>
    <div class="form-group" style="margin:0;">
      <label>Description</label>
      <input v-model="exportForm.description" placeholder="Description optionnelle" />
    </div>
  </div>
  <div style="display:flex;gap:8px;justify-content:flex-end;">
    <button class="btn btn-ghost btn-sm" @click="showExport=false">Annuler</button>
    <button class="btn btn-primary btn-sm" :disabled="loading" @click="exportCurrent">Exporter</button>
  </div>
</div>
```

- [ ] **Vérifier dans le navigateur**

`http://localhost:8000` → vue "Gabarits & Affaires" :
- "Nouveau" → modal avec seulement Nom + Description (plus de champ Chemin)
- "Cloner → Affaire" → modal avec seulement Nom de l'Affaire (plus de chemin destination)
- "Exporter l'Affaire active" → modal avec Nom + Description (plus de champ Destination)

Créer un gabarit test et vérifier qu'il apparaît dans la liste (scan auto `gabarits_dir`).

- [ ] **Commit**

```bash
git add web/registry_app/static/app.js
git commit -m "feat(N2 UI): ViewGabarits — champs path supprimés, création par nom seul"
```

---

## Task 9 — UI N1 : `ViewWorkspace` (`gabarits_dir` seul)

**Fichiers :**
- Modifier : `web/schema_app/static/app.js` — remplacer `ViewWorkspace`

- [ ] **Remplacer `ViewWorkspace` (lignes 1630–1727)**

```javascript
const ViewWorkspace = {
  setup() {
    const form    = ref({ gabarits_dir: '' });
    const loading = ref(true);
    const saved   = ref(false);
    const err     = ref('');

    onMounted(async () => {
      try {
        const d = await GET('/api/workspace');
        form.value = { gabarits_dir: d.gabarits_dir || '' };
      } catch(e) { toastErr(e); }
      loading.value = false;
    });

    const save = async () => {
      err.value = '';
      try {
        await PUT('/api/workspace', form.value);
        saved.value = true;
        setTimeout(() => saved.value = false, 2000);
        toast('Répertoire des Gabarits enregistré');
      } catch(e) { err.value = e.message; toastErr(e); }
    };

    const reset = async () => {
      if (!confirm('Réinitialiser le répertoire des Gabarits ?')) return;
      await DEL('/api/workspace');
      form.value = { gabarits_dir: '' };
      toast('Répertoire des Gabarits réinitialisé');
    };

    return { form, loading, saved, err, save, reset };
  },
  template: `
    <div v-if="loading" class="loading">Chargement…</div>
    <div v-else style="max-width:640px;">

      <div style="margin-bottom:20px;padding:14px 16px;background:var(--surface2);border-radius:8px;border-left:3px solid #a5b4fc;font-size:0.82rem;color:var(--text-dim);line-height:1.6;">
        N1 gère le <strong style="color:var(--text)">Répertoire des Gabarits</strong> —
        les schémas réutilisables copiés par N2 pour créer des Affaires.
        Stocké dans <code style="color:#a5b4fc;">.exosync_workspace.json</code> (partagé avec N2).
      </div>

      <div class="form-group">
        <label class="form-label">Répertoire des Gabarits</label>
        <input v-model="form.gabarits_dir" class="form-control"
          placeholder="ex: C:\\ExoSync\\Gabarits  ou  /home/user/gabarits" />
        <div style="font-size:0.75rem;color:var(--text-dim);margin-top:4px;">
          Chemin absolu. Chaque sous-dossier contenant <code>file_types.yaml</code> est détecté comme un gabarit.
          Les nouveaux gabarits créés ici seront placés automatiquement sous ce dossier.
        </div>
      </div>

      <div v-if="err" style="margin-top:12px;padding:10px 14px;background:#7f1d1d;border-radius:6px;font-size:0.82rem;color:#fca5a5;">
        ⚠ {{ err }}
      </div>

      <div style="display:flex;gap:10px;margin-top:20px;">
        <button class="btn btn-primary" @click="save">
          {{ saved ? '✓ Enregistré' : '💾 Enregistrer' }}
        </button>
        <button class="btn btn-ghost" @click="reset" style="color:#ef4444;border-color:#ef4444;">
          ✕ Réinitialiser
        </button>
      </div>

      <div style="margin-top:24px;padding:12px 16px;background:var(--surface2);border-radius:8px;font-size:0.78rem;color:var(--text-dim);">
        <div style="font-weight:600;color:var(--text);margin-bottom:8px;">Structure attendue</div>
        <pre style="font-family:monospace;font-size:0.75rem;line-height:1.7;margin:0;color:var(--text-dim);">{{ form.gabarits_dir || 'gabarits_dir' }}/
  SysEng/                 ← gabarit créé depuis N1
    file_types.yaml       ← requis
    schema_relations.json
    functions.json
  SysEng-v2/
    file_types.yaml</pre>
      </div>
    </div>
  `
};
```

- [ ] **Vérifier dans le navigateur**

`http://localhost:8001` → onglet Workspace → un seul champ `gabarits_dir`. Modifier et sauvegarder. Vérifier dans `.exosync_workspace.json` que `workspace_dir` (N2) est préservé.

- [ ] **Commit**

```bash
git add web/schema_app/static/app.js
git commit -m "feat(N1 UI): ViewWorkspace — workspace_dir supprimé, gabarits_dir seul"
```

---

## Task 10 — Vérification end-to-end et push

- [ ] **Test workflow complet**

```
1. N1 (http://localhost:8001) → Workspace → définir gabarits_dir
2. N2 (http://localhost:8000) → Configuration globale → vérifier gabarits_dir (RO) → définir workspace_dir
3. N2 → Gabarits → Nouveau → saisir un nom → vérifier que le dossier est créé dans gabarits_dir
4. N2 → Gabarits → Cloner → saisir un nom → vérifier que l'Affaire est créée dans workspace_dir
5. N2 → Répertoires de l'Affaire → définir posts_dir
6. Vérifier .exosync_workspace.json : les deux champs présents, cohérents
```

- [ ] **Mettre à jour la mémoire projet**

Dans `C:\Users\fabie\.claude\projects\C--Users-fabie-Documents-JLC-ObsidainRech\memory\project_exosync_state.md`, mettre à jour la section "Architecture Répertoires / Workspace" pour refléter le modèle 3 chemins / 3 propriétaires.

- [ ] **Push**

```bash
git push origin master
```
