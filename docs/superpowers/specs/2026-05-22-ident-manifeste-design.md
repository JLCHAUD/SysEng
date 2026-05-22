# IDENT — Déclaration des champs identitaires dans le manifeste MXL

> **Pour les agents d'implémentation :** Utiliser `superpowers:subagent-driven-development` ou `superpowers:executing-plans` pour implémenter ce plan tâche par tâche.

**Objectif :** Permettre à chaque manifeste MXL de s'autodéclarer ses champs identitaires via un nouveau mot-clé `IDENT`, sans dépendance à `file_types.yaml` au runtime.

**Architecture :** Nouveau `IdentNode` dans le parser MXL, génération `IDENT` dans `mxl_service.py`, lecture dans `execute_collects` depuis le `_Manifeste` du child.

**Périmètre :** `src/parser.py`, `web/mxl_service.py`, `src/executor.py`, nouveaux tests.

---

## Contexte

Actuellement, quand un fichier parent exécute un `COLLECT`, le moteur doit connaître
les colonnes identitaires du child pour les injecter en tête de chaque
ligne collectée. Cette information vient de `file_types.yaml` → `min_fields[is_key=True]`
— une source de vérité **externe** au fichier Excel.

Le problème : un fichier Excel n'est pas autonome. Si `file_types.yaml` est absent ou
désynchronisé, le moteur ne sait plus quels champs identifient ce Post.

**Solution :** Chaque manifeste déclare ses champs identitaires directement, via le
mot-clé `IDENT`. Le fichier Excel devient self-describing.

---

## Syntaxe MXL

### Structure du `_Manifeste` avec IDENT

```
A1 : MANIFESTE_V=1
A2 : # UO-001 — uo_instance

A3 : FILE_TYPE: uo_instance
A4 : FILE_ID:   UO-001
A5 : VERSION:   1
A6 : DOC:

     # ─── Champs identitaires ───────────────────────────────────────
A7 : IDENT nom         : LABEL="Nom de l'UO"       │ B7 : (valeur saisie par l'utilisateur)
A8 : IDENT responsable : LABEL="Responsable"        │ B8 : (valeur saisie par l'utilisateur)
A9 : IDENT site        : LABEL="Site de déploiement"│ B9 : (valeur saisie par l'utilisateur)

     # ─── Lecture locale ────────────────────────────────────────────
A10: DEF $activites = GET_TABLE(Activites, TabActivites)
...
```

### Règles

- `IDENT` se place **après l'en-tête fixe** (`FILE_TYPE` / `FILE_ID` / `VERSION` / `DOC`) et **avant les `DEF`**
- **Colonne A** = déclaration figée (générée, ne pas modifier)
- **Colonne B** = valeur saisie par l'utilisateur dans Excel
- Plusieurs `IDENT` autorisés sur un même Post
- `IDENT` = sémantique `is_key=True` : ces champs sont les colonnes identitaires injectées en tête de chaque ligne lors des `COLLECT`, pour identifier la source (quel child a produit cette ligne)
- Les champs metadata non-clés (`is_key=False`) qui servent aux filtres `LIST DYNAMIC WHERE` restent en lignes d'en-tête classiques (`nom: valeur`)
- `LABEL=` est optionnel ; si absent, `name` est utilisé comme label

---

## Parser (`src/parser.py`)

### Nouveau dataclass

```python
@dataclass
class IdentNode:
    name: str       # "nom"
    label: str      # "Nom de l'UO"
    value: str = "" # valeur col B, saisie par l'utilisateur
```

### `ManifestAST`

Nouveau champ :

```python
idents: List[IdentNode] = field(default_factory=list)
```

### `_parse_ident(line, anchor)`

```python
def _parse_ident(line: str, anchor: str) -> Optional[IdentNode]:
    m = re.match(r'^IDENT\s+([\w_]+)\s*:\s*(.*)$', line.strip())
    if not m:
        return None
    name  = m.group(1)
    attrs = _parse_kv_attrs(m.group(2))   # réutilise l'existant
    label = attrs.get("LABEL", name)
    return IdentNode(name=name, label=label, value=anchor.strip())
```

### `_parse_header_line`

Ajouter `"IDENT"` à la liste des mots-clés réservés pour éviter qu'il soit capturé comme `manifest_metadata`.

### `parse_lines`

Ajouter le cas `IDENT` dans le switch keyword :

```python
elif keyword == "IDENT":
    node = _parse_ident(instr, anchor or "")
    if node:
        ast.idents.append(node)
    else:
        ast.errors.append(ParseError(line_num, instr, "Syntaxe IDENT invalide"))
```

### `ast_summary`

Ajouter la ligne : `f"IDENT     : {len(ast.idents)} champ(s) identitaire(s)"`

---

## Générateur (`web/mxl_service.py`)

Dans `build_class_mxl_lines`, remplacer la génération des champs identitaires en en-tête par une section `IDENT` dédiée.

### Avant

```python
identitaires = [
    f for f in ft.get("min_fields", [])
    if f.get("source") in ("user_input", "reference", "parametre")
]
for f in identitaires:
    label = f.get("label", f["name"])
    lines.append(f"{f['name']}:   # {label}")
```

### Après

```python
# Champs identitaires (is_key=True) → section IDENT
ident_fields = [f for f in ft.get("min_fields", []) if f.get("is_key")]
if ident_fields:
    lines.append("# -- IDENTIFICATION -------------------------------------------")
    for f in ident_fields:
        label = f.get("label", f["name"])
        lines.append(f'IDENT {f["name"]} : LABEL="{label}"')
    lines.append("")

# Champs metadata non-clés (is_key=False, user_input) → en-tête classique
meta_fields = [
    f for f in ft.get("min_fields", [])
    if not f.get("is_key") and f.get("source") in ("user_input", "reference", "parametre")
]
for f in meta_fields:
    label = f.get("label", f["name"])
    lines.append(f"{f['name']}:   # {label}")
```

**Note :** Le générateur Excel (`xlsx_generator.py`) n'est pas modifié — il appelle
`build_class_mxl_lines` et la colonne B des lignes IDENT est vide à la génération
(l'utilisateur la remplit dans Excel).

---

## Moteur (`src/executor.py`)

### `execute_collects`

Après `wb_child = load_workbook(...)`, avant la lecture de la table source :

```python
from src.parser import parse_sheet, MANIFESTE_SHEET

# Lire les idents depuis le _Manifeste du child
ident_prefix: Dict[str, Any] = {}
if MANIFESTE_SHEET in wb_child.sheetnames:
    child_ast = parse_sheet(wb_child[MANIFESTE_SHEET])
    if child_ast.idents:
        # IDENT déclaré → manifeste = seule source de vérité
        ident_prefix = {i.name: i.value for i in child_ast.idents}
    else:
        # Ancien manifeste sans IDENT → fallback sur entry.context
        ident_prefix = dict(entry.context)
else:
    ident_prefix = dict(entry.context)

# Filtre WITH si spécifié dans le COLLECT
with_fields = getattr(collect, "with_fields", []) or []
if with_fields:
    ident_prefix = {k: v for k, v in ident_prefix.items() if k in with_fields}
```

Remplacer le bloc d'enrichissement existant (lignes 553-564) par :

```python
for row in rows:
    enriched: Dict[str, Any] = {"_source_file_id": entry.file_id}
    enriched.update(ident_prefix)
    enriched.update(row)
    all_rows.append(enriched)
```

**Points clés :**
- Le workbook child est déjà ouvert → pas de surcoût I/O
- `with_fields` continue de fonctionner pour filtrer les colonnes identitaires
- Fallback `entry.context` pour les anciens manifestes (rétro-compat)
- Aucun appel à `file_types.yaml` au runtime pour un manifeste IDENT

---

## Tests

### `tests/test_parser_ident.py` (nouveau)

```python
from src.parser import parse_lines, IdentNode

def test_parse_ident_basic():
    lines = [
        ("FILE_TYPE: uo_instance", ""),
        ("FILE_ID:   UO-001", ""),
        ("IDENT nom : LABEL=\"Nom de l'UO\"", "Jean Dupont"),
        ("IDENT responsable : LABEL=\"Responsable\"", "Marie Martin"),
        ("DEF $activites = GET_TABLE(Activites, TabActivites)", ""),
    ]
    ast = parse_lines(lines)
    assert len(ast.idents) == 2
    assert ast.idents[0].name == "nom"
    assert ast.idents[0].label == "Nom de l'UO"
    assert ast.idents[0].value == "Jean Dupont"
    assert ast.idents[1].name == "responsable"
    assert ast.idents[1].value == "Marie Martin"
    assert ast.errors == []

def test_ident_not_in_metadata():
    lines = [("IDENT nom : LABEL=\"Nom\"", "Jean")]
    ast = parse_lines(lines)
    assert "nom" not in ast.header.manifest_metadata
    assert len(ast.idents) == 1

def test_ident_label_fallback():
    lines = [("IDENT site :", "Paris")]
    ast = parse_lines(lines)
    assert ast.idents[0].label == "site"
    assert ast.idents[0].value == "Paris"

def test_no_ident_is_valid():
    lines = [("FILE_TYPE: uo_instance", ""), ("DEF $t = GET_TABLE(S, T)", "")]
    ast = parse_lines(lines)
    assert ast.idents == []
    assert ast.errors == []

def test_multiple_idents():
    lines = [
        ("IDENT nom : LABEL=\"Nom\"", "Alice"),
        ("IDENT site : LABEL=\"Site\"", "Paris"),
        ("IDENT region : LABEL=\"Région\"", "Île-de-France"),
    ]
    ast = parse_lines(lines)
    assert len(ast.idents) == 3
    assert [i.name for i in ast.idents] == ["nom", "site", "region"]
```

### `tests/test_mxl_service_ident.py` (nouveau)

```python
from web.mxl_service import build_class_mxl_lines

def test_generates_ident_for_key_fields():
    ft = {"min_fields": [
        {"name": "nom", "label": "Nom de l'UO", "is_key": True},
        {"name": "statut", "label": "Statut", "is_key": False, "source": "user_input"},
    ]}
    lines = build_class_mxl_lines("uo", "UO-001", ft, [])
    ident_lines = [l for l in lines if l.startswith("IDENT")]
    assert len(ident_lines) == 1
    assert 'IDENT nom : LABEL="Nom de l\'UO"' in ident_lines[0]

def test_non_key_fields_remain_header():
    ft = {"min_fields": [
        {"name": "nom", "label": "Nom", "is_key": True},
        {"name": "statut", "label": "Statut", "is_key": False, "source": "user_input"},
    ]}
    lines = build_class_mxl_lines("uo", "UO-001", ft, [])
    header_lines = [l for l in lines if l.startswith("statut:")]
    assert len(header_lines) == 1
```

### `tests/test_executor_integration.py` (extension)

Ajouter un test COLLECT avec un child dont le `_Manifeste` contient des `IDENT` et vérifier que les colonnes `nom` et `responsable` apparaissent en tête des lignes collectées.

---

## Fichiers touchés

| Fichier | Nature |
|---|---|
| `src/parser.py` | Modifier — `IdentNode`, `ast.idents`, `_parse_ident()`, switch IDENT |
| `web/mxl_service.py` | Modifier — section IDENT depuis `min_fields[is_key=True]` |
| `src/executor.py` | Modifier — `execute_collects` lit `child_ast.idents` |
| `tests/test_parser_ident.py` | Créer — 5 tests parser |
| `tests/test_mxl_service_ident.py` | Créer — 2 tests générateur |
| `tests/test_executor_integration.py` | Modifier — 1 test COLLECT avec IDENT |
