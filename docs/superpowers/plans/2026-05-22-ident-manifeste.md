# IDENT — Déclaration des champs identitaires dans le manifeste MXL

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ajouter le mot-clé `IDENT` au langage MXL pour que chaque manifeste s'autodéclare ses champs identitaires, utilisés en tête des lignes lors des COLLECT.

**Architecture:** Trois changements indépendants en cascade — (1) parser : nouveau `IdentNode` + keyword IDENT ; (2) générateur : `build_class_mxl_lines` émet des lignes `IDENT` depuis `min_fields[is_key=True]` ; (3) executor : `execute_collects` lit `child_ast.idents` pour construire le préfixe identitaire de chaque ligne collectée, avec fallback sur `entry.context` pour les anciens manifestes.

**Tech Stack:** Python 3.13, dataclasses, re, openpyxl, pytest.

---

## Contexte codebase

### `src/parser.py` — structure actuelle pertinente

- **Ligne 154** : fin de `CollectNode` → c'est ici qu'on insère `IdentNode`
- **Ligne 161** : `ManifestAST` → on y ajoute le champ `idents`
- **Ligne 240** : `_KNOWN_HEADER_KEYS = {"FILE_TYPE", "FILE_ID", "VERSION", "DOC"}`
- **Lignes 265-266** : set des mots-clés MXL réservés dans `_parse_header_line` — on y ajoute `"IDENT"`
- **Ligne 595** : fin de `_parse_collect` → on insère `_parse_ident` juste après
- **Lignes 688-697** : bloc `elif keyword == "COLLECT"` suivi de `else:` → on insère le cas IDENT avant le `else:`
- **Lignes 953-961** : `ast_summary` → on ajoute la ligne IDENT

### `web/mxl_service.py` — section à remplacer

- **Lignes 77-86** : bloc `identitaires = [...]` + boucle `lines.append(...)` → remplacé par deux blocs distincts (IDENT pour is_key=True, en-tête classique pour is_key=False)

### `src/executor.py` — section à modifier

- **Ligne 31** : `from src.parser import ManifestAST` → ajouter `parse_sheet, MANIFESTE_SHEET`
- **Lignes 553-565** : bloc "Colonnes contextuelles" + "Enrichissement" → remplacé par la lecture des idents du child

---

## Fichiers

| Fichier | Action |
|---|---|
| `tests/test_parser_ident.py` | Créer |
| `src/parser.py` | Modifier |
| `tests/test_mxl_service_ident.py` | Créer |
| `web/mxl_service.py` | Modifier |
| `tests/test_collect_ident.py` | Créer |
| `src/executor.py` | Modifier |

---

## Task 1 — Parser : IdentNode + keyword IDENT

**Files:**
- Create: `tests/test_parser_ident.py`
- Modify: `src/parser.py`

- [ ] **Step 1 : Écrire les tests (ils vont échouer)**

Créer `tests/test_parser_ident.py` :

```python
"""Tests du mot-clé IDENT dans le parser MXL."""
import pytest
from src.parser import parse_lines, ManifestAST


def test_parse_ident_basic():
    """IDENT avec LABEL et valeur col B."""
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


def test_ident_not_captured_as_metadata():
    """IDENT ne doit PAS être stocké dans manifest_metadata."""
    lines = [("IDENT nom : LABEL=\"Nom\"", "Jean")]
    ast = parse_lines(lines)
    assert "nom" not in ast.header.manifest_metadata
    assert "ident" not in ast.header.manifest_metadata
    assert len(ast.idents) == 1


def test_ident_label_fallback_to_name():
    """Sans LABEL=, le label prend la valeur de name."""
    lines = [("IDENT site :", "Paris")]
    ast = parse_lines(lines)
    assert ast.idents[0].name == "site"
    assert ast.idents[0].label == "site"
    assert ast.idents[0].value == "Paris"


def test_no_ident_is_valid():
    """Un manifeste sans IDENT est valide — ast.idents est vide."""
    lines = [("FILE_TYPE: uo_instance", ""), ("DEF $t = GET_TABLE(S, T)", "")]
    ast = parse_lines(lines)
    assert ast.idents == []
    assert ast.errors == []


def test_multiple_idents():
    """Plusieurs IDENT sur un même Post."""
    lines = [
        ("IDENT nom : LABEL=\"Nom\"", "Alice"),
        ("IDENT site : LABEL=\"Site\"", "Paris"),
        ("IDENT region : LABEL=\"Région\"", "Île-de-France"),
    ]
    ast = parse_lines(lines)
    assert len(ast.idents) == 3
    assert [i.name for i in ast.idents] == ["nom", "site", "region"]
    assert [i.value for i in ast.idents] == ["Alice", "Paris", "Île-de-France"]
```

- [ ] **Step 2 : Vérifier que les tests échouent**

```
pytest tests/test_parser_ident.py -v
```

Attendu : `ImportError` ou `AttributeError: 'ManifestAST' object has no attribute 'idents'`

- [ ] **Step 3 : Ajouter `IdentNode` après `CollectNode` (ligne 154)**

Dans `src/parser.py`, ajouter après la fin de la classe `CollectNode` et avant `ParseError` :

```python
@dataclass
class IdentNode:
    name: str       # "nom"
    label: str      # "Nom de l'UO"
    value: str = "" # valeur col B, saisie par l'utilisateur dans Excel
```

- [ ] **Step 4 : Ajouter `idents` à `ManifestAST` (ligne 173)**

Dans `ManifestAST`, ajouter après le champ `collects` :

```python
    collects: List[CollectNode] = field(default_factory=list)
    idents:   List[IdentNode]   = field(default_factory=list)   # ← nouveau
```

- [ ] **Step 5 : Exclure `IDENT` de `manifest_metadata` (ligne 265)**

Dans `_parse_header_line`, modifier le set des mots-clés réservés :

```python
    elif key not in {"DEF", "COL", "BIND", "PUSH", "PULL",
                     "VALIDATE", "EXTENDS", "NOTIFY", "LIST", "COLLECT", "IDENT"}:
```

- [ ] **Step 6 : Ajouter `_parse_ident` après `_parse_collect`**

Insérer la fonction suivante juste après la fin de `_parse_collect` (après la ligne `return None` de _parse_collect) :

```python
def _parse_ident(line: str, anchor: str) -> Optional[IdentNode]:
    """
    IDENT nom : LABEL="Nom de l'UO"
    La valeur du champ vient de la colonne B (anchor).
    """
    m = re.match(r'^IDENT\s+([\w_]+)\s*:\s*(.*)$', line.strip())
    if not m:
        return None
    name  = m.group(1)
    attrs = _parse_kv_attrs(m.group(2))   # réutilise le parser d'attributs existant
    label = attrs.get("LABEL", name)
    return IdentNode(name=name, label=label, value=anchor.strip())
```

- [ ] **Step 7 : Ajouter le cas `IDENT` dans `parse_lines` (avant le `else:` final)**

Dans `parse_lines`, ajouter avant la ligne `else:` (le dernier `else:` qui génère "Mot-clé inconnu") :

```python
        elif keyword == "IDENT":
            node = _parse_ident(instr, anchor or "")
            if node:
                ast.idents.append(node)
            else:
                ast.errors.append(ParseError(line_num, instr, "Syntaxe IDENT invalide"))
```

- [ ] **Step 8 : Mettre à jour `ast_summary`**

Dans `ast_summary`, ajouter la ligne IDENT après la ligne COLLECT :

```python
        f"COLLECT   : {len(ast.collects)} agrégation(s)",
        f"IDENT     : {len(ast.idents)} champ(s) identitaire(s)",   # ← nouveau
        f"Metadata  : {ast.header.manifest_metadata or '—'}",
```

- [ ] **Step 9 : Lancer les tests**

```
pytest tests/test_parser_ident.py -v
```

Attendu : `5 passed`

- [ ] **Step 10 : Vérifier que les tests existants passent encore**

```
pytest tests/ -v --ignore=tests/test_parser_ident.py
```

Attendu : tous verts (aucune régression)

- [ ] **Step 11 : Commit**

```bash
git add tests/test_parser_ident.py src/parser.py
git commit -m "feat: parser MXL — nouveau mot-clé IDENT (IdentNode + ast.idents)

Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>"
```

---

## Task 2 — Générateur : `build_class_mxl_lines` émet des lignes IDENT

**Files:**
- Create: `tests/test_mxl_service_ident.py`
- Modify: `web/mxl_service.py:77-86`

- [ ] **Step 1 : Écrire les tests (ils vont échouer)**

Créer `tests/test_mxl_service_ident.py` :

```python
"""Tests de génération IDENT dans build_class_mxl_lines."""
from web.mxl_service import build_class_mxl_lines


def test_is_key_true_generates_ident_line():
    """min_fields avec is_key=True → ligne IDENT dans la sortie."""
    ft = {
        "min_fields": [
            {"name": "nom", "label": "Nom de l'UO", "is_key": True},
            {"name": "statut", "label": "Statut", "is_key": False, "source": "user_input"},
        ]
    }
    lines = build_class_mxl_lines("uo", "UO-001", ft, [])
    ident_lines = [l for l in lines if l.startswith("IDENT")]
    assert len(ident_lines) == 1
    assert ident_lines[0] == 'IDENT nom : LABEL="Nom de l\'UO"'


def test_is_key_false_remains_header_metadata():
    """min_fields avec is_key=False → ligne en-tête classique (nom: # label)."""
    ft = {
        "min_fields": [
            {"name": "nom",    "label": "Nom", "is_key": True},
            {"name": "statut", "label": "Statut", "is_key": False, "source": "user_input"},
        ]
    }
    lines = build_class_mxl_lines("uo", "UO-001", ft, [])
    header_lines = [l for l in lines if l.startswith("statut:")]
    assert len(header_lines) == 1
    assert "# Statut" in header_lines[0]


def test_no_min_fields_no_ident():
    """Classe sans min_fields → aucune ligne IDENT."""
    ft = {}
    lines = build_class_mxl_lines("uo", "UO-001", ft, [])
    assert not any(l.startswith("IDENT") for l in lines)


def test_multiple_is_key_fields():
    """Plusieurs is_key=True → plusieurs lignes IDENT dans l'ordre."""
    ft = {
        "min_fields": [
            {"name": "nom",  "label": "Nom",  "is_key": True},
            {"name": "site", "label": "Site", "is_key": True},
        ]
    }
    lines = build_class_mxl_lines("uo", "UO-001", ft, [])
    ident_lines = [l for l in lines if l.startswith("IDENT")]
    assert len(ident_lines) == 2
    assert ident_lines[0] == 'IDENT nom : LABEL="Nom"'
    assert ident_lines[1] == 'IDENT site : LABEL="Site"'
```

- [ ] **Step 2 : Vérifier que les tests échouent**

```
pytest tests/test_mxl_service_ident.py -v
```

Attendu : `FAILED` — les tests vérifient `startswith("IDENT")` mais la version actuelle génère `nom:   # Nom de l'UO` à la place.

- [ ] **Step 3 : Modifier `web/mxl_service.py` lignes 77-86**

Remplacer le bloc existant :

```python
    # min_fields identitaires / paramètres → métadonnées libres dans l'en-tête
    # Le parser les stocke dans manifest_metadata pour les filtres LIST DYNAMIC.
    identitaires = [
        f for f in ft.get("min_fields", [])
        if f.get("source") in ("user_input", "reference", "parametre")
    ]
    for f in identitaires:
        label = f.get("label", f["name"])
        lines.append(f"{f['name']}:   # {label}")
    lines.append("")
```

Par :

```python
    # Champs identitaires (is_key=True) → section IDENT (autodéclaration)
    ident_fields = [f for f in ft.get("min_fields", []) if f.get("is_key")]
    if ident_fields:
        lines.append("# -- IDENTIFICATION -------------------------------------------")
        for f in ident_fields:
            label = f.get("label", f["name"])
            lines.append(f'IDENT {f["name"]} : LABEL="{label}"')
        lines.append("")

    # Champs metadata non-clés (is_key=False, user_input) → en-tête classique
    # Stockés dans manifest_metadata, utilisables dans LIST DYNAMIC WHERE.
    meta_fields = [
        f for f in ft.get("min_fields", [])
        if not f.get("is_key") and f.get("source") in ("user_input", "reference", "parametre")
    ]
    for f in meta_fields:
        label = f.get("label", f["name"])
        lines.append(f"{f['name']}:   # {label}")
    if meta_fields:
        lines.append("")
```

- [ ] **Step 4 : Lancer les tests**

```
pytest tests/test_mxl_service_ident.py -v
```

Attendu : `4 passed`

- [ ] **Step 5 : Vérifier la suite complète**

```
pytest tests/ -v --ignore=tests/test_collect_ident.py
```

Attendu : tous verts

- [ ] **Step 6 : Commit**

```bash
git add tests/test_mxl_service_ident.py web/mxl_service.py
git commit -m "feat: mxl_service — génère IDENT pour min_fields[is_key=True]

Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>"
```

---

## Task 3 — Executor : COLLECT lit les idents du child

**Files:**
- Create: `tests/test_collect_ident.py`
- Modify: `src/executor.py:31` et `src/executor.py:553-565`

- [ ] **Step 1 : Écrire le test d'intégration (il va échouer)**

Créer `tests/test_collect_ident.py` :

```python
"""
Test d'intégration : COLLECT injecte les colonnes IDENT du child en tête des lignes.

Crée deux vrais fichiers Excel (child + parent) dans tmp_path,
exécute le parent via execute_ast, et vérifie que les colonnes identitaires
déclarées par IDENT dans le child apparaissent en tête de la table collectée.
"""
import pytest
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table
from openpyxl.utils import get_column_letter

from src.parser import parse_file
from src.executor import execute_ast
from src import store as Store


def _make_named_table(ws, table_name: str, rows: list) -> None:
    """
    Crée une table Excel nommée dans la feuille ws.
    rows[0] = headers, rows[1:] = données.
    """
    for row in rows:
        ws.append(row)
    n_cols = len(rows[0])
    n_rows = len(rows)
    ref = f"A1:{get_column_letter(n_cols)}{n_rows}"
    ws.add_table(Table(displayName=table_name, ref=ref))


def test_collect_uses_ident_columns(tmp_path):
    """
    COLLECT injecte les colonnes identitaires (IDENT) du child en tête.

    Setup :
    - child  ENFANT-001.xlsx : IDENT nom="UO-Alpha", IDENT site="Paris"
                               table TabTaches : [id, libelle]
    - parent parent.xlsx     : LIST MesUOs FROM TABLE ListeUOs
                               COLLECT TabTaches FROM MesUOs INTO VueTaches

    Attendu : VueTaches contient _source_file_id + nom + site + id + libelle,
              avec "UO-Alpha" et "Paris" correctement injectés.
    """
    # ── Child : ENFANT-001.xlsx ───────────────────────────────────────────────
    child_path = tmp_path / "ENFANT-001.xlsx"
    wbc = Workbook()
    wbc.remove(wbc.active)

    # _Manifeste du child avec deux IDENT
    ws_mc = wbc.create_sheet("_Manifeste")
    ws_mc["A1"] = "MANIFESTE_V=1"
    # row 2 = titre décoratif, ignoré par parse_sheet (min_row=3)
    ws_mc["A3"] = "FILE_TYPE: uo_test"
    ws_mc["A4"] = "FILE_ID:   ENFANT-001"
    ws_mc["A5"] = "VERSION:   1"
    ws_mc["A6"] = 'IDENT nom : LABEL="Nom"'
    ws_mc["B6"] = "UO-Alpha"           # col B = valeur saisie par l'utilisateur
    ws_mc["A7"] = 'IDENT site : LABEL="Site"'
    ws_mc["B7"] = "Paris"

    # Feuille Taches avec table nommée TabTaches
    ws_tc = wbc.create_sheet("Taches")
    _make_named_table(ws_tc, "TabTaches", [
        ["id",  "libelle"],
        ["T1",  "Analyse"],
        ["T2",  "Conception"],
    ])
    wbc.save(str(child_path))

    # ── Parent : parent.xlsx ──────────────────────────────────────────────────
    parent_path = tmp_path / "parent.xlsx"
    wbp = Workbook()
    wbp.remove(wbp.active)

    # _Manifeste du parent : LIST + COLLECT
    ws_mp = wbp.create_sheet("_Manifeste")
    ws_mp["A1"] = "MANIFESTE_V=1"
    ws_mp["A3"] = "FILE_TYPE: projet"
    ws_mp["A4"] = "FILE_ID:   PROJ-001"
    ws_mp["A5"] = "VERSION:   1"
    ws_mp["A6"] = "LIST MesUOs FROM TABLE ListeUOs"
    ws_mp["A7"] = "COLLECT TabTaches FROM MesUOs INTO VueTaches"

    # Table de liste des UOs (colonne FILE_ID obligatoire)
    ws_l = wbp.create_sheet("Liste")
    _make_named_table(ws_l, "ListeUOs", [
        ["FILE_ID"],
        ["ENFANT-001"],
    ])

    # Table cible (pré-existante, sera écrasée par COLLECT)
    ws_v = wbp.create_sheet("Vue")
    _make_named_table(ws_v, "VueTaches", [
        ["nom", "site", "id", "libelle"],
        ["",    "",     "",   ""],        # ligne vide initiale
    ])
    wbp.save(str(parent_path))

    # ── Exécution ─────────────────────────────────────────────────────────────
    ast = parse_file(parent_path)
    result = execute_ast(ast, parent_path, Store)

    assert not result.errors, f"Erreurs executor : {result.errors}"
    assert len(result.collected) == 1, f"COLLECT attendu, got : {result.collected}"

    # ── Vérification de VueTaches ─────────────────────────────────────────────
    wb_r = load_workbook(str(parent_path), data_only=True)
    # Trouver la feuille contenant VueTaches
    ws_result = None
    for sn in wb_r.sheetnames:
        if "VueTaches" in wb_r[sn].tables:
            ws_result = wb_r[sn]
            break
    assert ws_result is not None, "Table VueTaches introuvable dans le parent après COLLECT"

    tbl_ref = ws_result.tables["VueTaches"].ref
    cells   = list(ws_result[tbl_ref])
    headers = [c.value for c in cells[0]]
    data_rows = [
        [c.value for c in row]
        for row in cells[1:]
        if any(c.value is not None for c in row)
    ]
    wb_r.close()

    # Les colonnes identitaires doivent précéder les colonnes de données du child
    assert "nom"  in headers, f"'nom' absent des headers : {headers}"
    assert "site" in headers, f"'site' absent des headers : {headers}"
    assert "id"   in headers, f"'id' absent des headers : {headers}"
    assert headers.index("nom")  < headers.index("id"),  "nom doit précéder id"
    assert headers.index("site") < headers.index("id"),  "site doit précéder id"

    # Les valeurs IDENT sont correctement injectées sur chaque ligne
    nom_idx  = headers.index("nom")
    site_idx = headers.index("site")
    id_idx   = headers.index("id")

    assert len(data_rows) == 2, f"Attendu 2 lignes collectées, got {len(data_rows)}"
    assert data_rows[0][nom_idx]  == "UO-Alpha"
    assert data_rows[0][site_idx] == "Paris"
    assert data_rows[0][id_idx]   == "T1"
    assert data_rows[1][nom_idx]  == "UO-Alpha"
    assert data_rows[1][site_idx] == "Paris"
    assert data_rows[1][id_idx]   == "T2"
```

- [ ] **Step 2 : Vérifier que le test échoue**

```
pytest tests/test_collect_ident.py -v
```

Attendu : `FAILED` — les colonnes `nom` et `site` seront absentes ou mal placées car le code lit encore `entry.context` et non `child_ast.idents`.

- [ ] **Step 3 : Mettre à jour l'import en tête de `src/executor.py` (ligne 31)**

Remplacer :

```python
from src.parser import ManifestAST
```

Par :

```python
from src.parser import ManifestAST, parse_sheet, MANIFESTE_SHEET
```

- [ ] **Step 4 : Modifier `execute_collects` dans `src/executor.py` (lignes 553-565)**

Remplacer le bloc :

```python
                # Colonnes contextuelles : WITH (liste DYNAMIC) ou tout (liste TABLE)
                with_fields = getattr(collect, "with_fields", []) or []
                if with_fields:
                    context = {k: entry.context.get(k) for k in with_fields}
                else:
                    context = dict(entry.context)

                # Enrichissement : _source_file_id en premier, puis contexte, puis données fils
                for row in rows:
                    enriched: Dict[str, Any] = {"_source_file_id": entry.file_id}
                    enriched.update(context)
                    enriched.update(row)
                    all_rows.append(enriched)
```

Par :

```python
                # Colonnes identitaires : lire les IDENT depuis le _Manifeste du child
                # (wb_child est déjà ouvert — pas de surcoût I/O)
                ident_prefix: Dict[str, Any] = {}
                if MANIFESTE_SHEET in wb_child.sheetnames:
                    child_ast = parse_sheet(wb_child[MANIFESTE_SHEET])
                    if child_ast.idents:
                        # Manifeste IDENT → source de vérité autonome
                        ident_prefix = {i.name: i.value for i in child_ast.idents}
                    else:
                        # Ancien manifeste sans IDENT → fallback sur entry.context
                        ident_prefix = dict(entry.context)
                else:
                    ident_prefix = dict(entry.context)

                # Filtre WITH si spécifié dans le COLLECT
                with_fields = getattr(collect, "with_fields", []) or []
                if with_fields:
                    ident_prefix = {k: v for k, v in ident_prefix.items()
                                    if k in with_fields}

                # Enrichissement : _source_file_id en premier, puis idents, puis données fils
                for row in rows:
                    enriched: Dict[str, Any] = {"_source_file_id": entry.file_id}
                    enriched.update(ident_prefix)
                    enriched.update(row)
                    all_rows.append(enriched)
```

- [ ] **Step 5 : Lancer le test d'intégration**

```
pytest tests/test_collect_ident.py -v
```

Attendu : `1 passed`

- [ ] **Step 6 : Lancer la suite complète**

```
pytest tests/ -v
```

Attendu : tous verts — vérifier en particulier `test_schema_factory.py` (10 tests) et `test_parser_ident.py` (5 tests) et `test_mxl_service_ident.py` (4 tests).

- [ ] **Step 7 : Commit**

```bash
git add tests/test_collect_ident.py src/executor.py
git commit -m "feat: executor COLLECT — lit les IDENT du child au lieu de file_types.yaml

Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>"
```
