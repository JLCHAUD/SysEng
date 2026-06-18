# Refonte `_Manifeste` MXL + Pilotes n:m — Plan d'implémentation

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Réécrire les feuilles `_Manifeste` des générateurs cockpit et dashboard en format MXL mono-colonne (col A=instruction, col B=ancre, col C=commentaire), et ajouter un champ `pilotes: Dict[str, str]` sur `UOInstance` pour permettre au dashboard de découvrir ses UOs automatiquement via `LIST DYNAMIC`.

**Architecture:** Les générateurs produisaient le format tabulaire multi-colonnes (parsé par `passerelle.py` legacy) alors que `sync.py` essaie d'abord le parser MXL mono-colonne (`parser.py`). Ce plan corrige ce désalignement. Le champ `pilotes` sur l'UO est injecté à la génération et stocké dans `manifest_metadata` lors du sync, permettant à `LIST DYNAMIC WHERE pilote_metier_ts=USR004` de fonctionner via l'Exomap déjà implémentée dans `executor.py`.

**Tech Stack:** Python 3.10+, openpyxl, pytest. Fichiers clés : `src/models.py`, `src/config_loader.py`, `src/generators/cockpit_ingenieur_generator.py`, `src/generators/dashboard_metier_generator.py`, `tests/test_cockpit_ingenieur.py`, `tests/test_dashboard_metier.py`.

---

## Fichiers modifiés

| Fichier | Action |
|---------|--------|
| `src/models.py:162` | Ajouter `pilotes: Dict[str, str]` sur `UOInstance` |
| `src/config_loader.py:120` | Lire `pilotes` depuis le JSON |
| `src/generators/cockpit_ingenieur_generator.py:290` | Réécrire `_sheet_manifeste()` en MXL + ajouter `tbl_mes_uos` dans `_sheet_mes_uos()` |
| `src/generators/dashboard_metier_generator.py:332` | Réécrire `_sheet_manifeste_dashboard()` en MXL |
| `tests/test_cockpit_ingenieur.py:120` | Mettre à jour `TestCockpitManifeste` pour le format MXL |
| `tests/test_dashboard_metier.py:156` | Mettre à jour `TestDashboardManifeste` pour le format MXL |

---

## Contexte codebase (lire avant de commencer)

### Format MXL attendu
```
col A : instruction MXL         ex: FILE_TYPE: cockpit_ingenieur
col B : ancre (cellule cible)   ex: Synthèse.G5  (vide si pas d'ancre)
col C : commentaire français    ex: Type de fichier ExoSync
```
`parser.py:parse_sheet` lit col A depuis la ligne 3 (ligne 1 = version, ligne 2 = skippée).

### Constantes de style disponibles dans les générateurs
```python
from src.generators.cockpit_ingenieur_generator import (
    BLUE_DARK, body_font, solid_fill, left
)
```
Ces fonctions/constantes existent déjà — ne pas les recréer.

### Table openpyxl nommée
```python
from openpyxl.worksheet.table import Table, TableStyleInfo
tbl = Table(displayName="tbl_mes_uos", ref="A5:I8")
ws.add_table(tbl)
```
`executor.py:_read_table_from_ws(ws, "tbl_mes_uos")` cherche le nom exact.

### Fixture de test réutilisable dans les deux fichiers de test
```python
from src.models import UOInstance, UOType, Activity, System, Project, StatutUO
from datetime import date

def _make_uo(uid, engineer, hours, end=date(2026,7,1), pilotes=None):
    return UOInstance(
        id=uid, uo_type_id="TS", system_id="SYS1", project_id="PRJ1",
        engineer_name=engineer, total_hours=hours,
        start_date=date(2026,1,1), end_date=end, statut=StatutUO.EN_COURS,
        pilotes=pilotes or {},
        uo_type=UOType(id="TS", name="Type TS"),
        system=System(id="SYS1", name="Système 1"),
        project=Project(id="PRJ1", name="Projet 1"),
    )
```

---

## Task 1 — Champ `pilotes` sur `UOInstance`

**Files:**
- Modify: `src/models.py:1-5,162-175`
- Modify: `src/config_loader.py:1-14,120-137`
- Test: `tests/test_models_pilotes.py` (nouveau fichier)

### Objectif
Ajouter `pilotes: Dict[str, str]` sur `UOInstance`. Clés = rôles libres (`metier_ts`, `metier_projet`…), valeurs = IDs acteurs (`USR004`). Valeur par défaut = `{}`.

- [ ] **Step 1 : Écrire le test qui échoue**

```python
# tests/test_models_pilotes.py
import json
from pathlib import Path
from unittest.mock import patch

def test_uo_instance_has_pilotes_field():
    """UOInstance doit avoir un champ pilotes vide par défaut."""
    from src.models import UOInstance, StatutUO
    from datetime import date
    uo = UOInstance(
        id="X", uo_type_id="T", system_id="S", project_id="P",
        engineer_name="Alice", total_hours=10,
        start_date=date(2026,1,1), end_date=date(2026,6,1),
    )
    assert hasattr(uo, "pilotes")
    assert uo.pilotes == {}


def test_uo_instance_pilotes_populated():
    """Le champ pilotes accepte un dict rôle → id."""
    from src.models import UOInstance, StatutUO
    from datetime import date
    uo = UOInstance(
        id="X", uo_type_id="T", system_id="S", project_id="P",
        engineer_name="Alice", total_hours=10,
        start_date=date(2026,1,1), end_date=date(2026,6,1),
        pilotes={"metier_ts": "USR004", "metier_projet": "USR007"},
    )
    assert uo.pilotes["metier_ts"] == "USR004"
    assert uo.pilotes["metier_projet"] == "USR007"


def test_config_loader_reads_pilotes(tmp_path):
    """load_uo_instances() doit lire le champ pilotes depuis le JSON."""
    import json
    from unittest.mock import patch
    from src.config_loader import load_uo_instances

    uo_data = [{
        "id": "UO-TEST",
        "uo_type_id": "TS",
        "system_id": "SYS1",
        "project_id": "PRJ1",
        "engineer_name": "Alice Dubois",
        "total_hours": 32,
        "start_date": "2026-01-01",
        "end_date": "2026-06-30",
        "statut": "EN_COURS",
        "pilotes": {"metier_ts": "USR004"},
    }]

    with patch("src.config_loader._load_json") as mock_load:
        def side_effect(filename):
            if filename == "uo_instances.json":
                return uo_data
            return {} if filename.endswith(".json") else []
        mock_load.side_effect = side_effect

        # load_uo_instances charge aussi les référentiels — on fournit des stubs
        with patch("src.config_loader.load_uo_types", return_value={}), \
             patch("src.config_loader.load_systems", return_value={}), \
             patch("src.config_loader.load_projects", return_value={}):
            instances = load_uo_instances()

    assert len(instances) == 1
    assert instances[0].pilotes == {"metier_ts": "USR004"}
```

- [ ] **Step 2 : Lancer le test — vérifier qu'il échoue**

```
pytest tests/test_models_pilotes.py -v
```
Expected: FAIL — `UOInstance.__init__() got an unexpected keyword argument 'pilotes'`

- [ ] **Step 3 : Modifier `src/models.py`**

Ajouter `Dict` à l'import typing (ligne 5) et le champ `pilotes` sur `UOInstance` :

```python
# src/models.py — ligne 4 : remplacer
from typing import List, Optional
# par :
from typing import Dict, List, Optional
```

```python
# src/models.py — après owner_id (ligne 174), avant le commentaire "# Références résolues"
    pilotes: Dict[str, str] = field(default_factory=dict)
    # ex: {"metier_ts": "USR004", "metier_projet": "USR007"}
```

- [ ] **Step 4 : Modifier `src/config_loader.py`**

Dans `load_uo_instances()`, ajouter `pilotes` à l'appel `UOInstance(...)` (ligne ~133) :

```python
        instances.append(UOInstance(
            id=item["id"],
            uo_type_id=item["uo_type_id"],
            system_id=item["system_id"],
            project_id=item["project_id"],
            engineer_name=item["engineer_name"],
            total_hours=item["total_hours"],
            start_date=date.fromisoformat(item["start_date"]),
            end_date=date.fromisoformat(item["end_date"]),
            statut=StatutUO(item.get("statut", "BROUILLON")),
            degrade=item.get("degrade", False),
            degrade_note=item.get("degrade_note", ""),
            owner_id=item.get("owner_id"),
            pilotes=item.get("pilotes", {}),   # ← ligne ajoutée
            uo_type=resolved_type,
            system=system,
            project=project,
        ))
```

- [ ] **Step 5 : Lancer les tests — vérifier qu'ils passent**

```
pytest tests/test_models_pilotes.py -v
```
Expected: 3 PASSED

- [ ] **Step 6 : Vérifier non-régression**

```
python -m pytest -q
```
Expected: tous les tests existants passent toujours.

- [ ] **Step 7 : Commit**

```
git add src/models.py src/config_loader.py tests/test_models_pilotes.py
git commit -m "feat: champ pilotes Dict[role,id] sur UOInstance + lecture config_loader"
```

---

## Task 2 — Table nommée `tbl_mes_uos` dans l'onglet Mes UOs

**Files:**
- Modify: `src/generators/cockpit_ingenieur_generator.py:75-156`
- Test: `tests/test_cockpit_ingenieur.py` (ajouter une assertion dans la classe existante)

### Objectif
Ajouter un tableau openpyxl nommé `tbl_mes_uos` sur la plage `A5:I{last_row}` après écriture des données. `executor.py:_read_table_from_ws(ws, "tbl_mes_uos")` cherche ce nom exact.

- [ ] **Step 1 : Ajouter l'import Table dans le générateur**

Dans `src/generators/cockpit_ingenieur_generator.py`, ajouter en tête des imports :

```python
from openpyxl.worksheet.table import Table, TableStyleInfo
```

- [ ] **Step 2 : Modifier `_sheet_mes_uos()` — ajouter la table après la boucle**

Localiser la ligne `last_row = 5 + len(uo_list)` (ligne ~133).
Après les `conditional_formatting.add(...)` et avant `set_column_widths(...)`, ajouter :

```python
    # Table nommée pour GET_TABLE(Mes UOs, tbl_mes_uos)
    if uo_list:
        tbl_ref = f"A5:I{last_row}"
        tbl = Table(displayName="tbl_mes_uos", ref=tbl_ref)
        tbl.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2", showFirstColumn=False,
            showLastColumn=False, showRowStripes=True, showColumnStripes=False,
        )
        ws.add_table(tbl)
```

- [ ] **Step 3 : Écrire le test (ajouter dans `tests/test_cockpit_ingenieur.py`)**

Dans la classe `TestCockpitMesUOs`, ajouter :

```python
    def test_table_nommee_tbl_mes_uos_presente(self, tmp_path):
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["Mes UOs"]
        table_names = list(ws.tables.keys())
        assert "tbl_mes_uos" in table_names
```

- [ ] **Step 4 : Lancer les tests**

```
pytest tests/test_cockpit_ingenieur.py::TestCockpitMesUOs::test_table_nommee_tbl_mes_uos_presente -v
```
Expected: PASS

- [ ] **Step 5 : Non-régression**

```
pytest tests/test_cockpit_ingenieur.py -v
```
Expected: tous les tests du fichier passent.

- [ ] **Step 6 : Commit**

```
git add src/generators/cockpit_ingenieur_generator.py tests/test_cockpit_ingenieur.py
git commit -m "feat: table nommee tbl_mes_uos dans onglet Mes UOs (prerequis GET_TABLE MXL)"
```

---

## Task 3 — Réécriture `_sheet_manifeste()` cockpit en MXL mono-colonne (TDD)

**Files:**
- Modify: `tests/test_cockpit_ingenieur.py:120-156` — mettre à jour `TestCockpitManifeste`
- Modify: `src/generators/cockpit_ingenieur_generator.py:290-332` — réécrire `_sheet_manifeste()`

### Objectif
Remplacer le format tabulaire 12-colonnes par MXL mono-colonne. Supprimer la constante `MANIFESTE_HEADERS`. Commentaires en col C.

Format attendu après refonte :
```
A1: MANIFESTE_V=1
A3: FILE_TYPE: cockpit_ingenieur          C3: Type de fichier ExoSync
A4: FILE_ID: Cockpit_Alice_Dubois         C4: Identifiant unique du cockpit
A5: ingenieur: Alice Dubois               C5: Nom de l'ingénieur propriétaire
A7: DEF $mes_uos = GET_TABLE(Mes UOs, tbl_mes_uos)   C7: Référence à la table des UOs
A8: COL $mes_uos.avancement : WRITE=engineer          C8: % avancement saisi par l'ingénieur
A9: COL $mes_uos.heures_realisees : WRITE=engineer    C9: Heures réalisées saisies
A11: PUSH $mes_uos -> cockpit.Cockpit_Alice_Dubois.mes_uos   C11: Export vers store central
```

- [ ] **Step 1 : Réécrire `TestCockpitManifeste` pour le nouveau format**

Remplacer entièrement la classe `TestCockpitManifeste` dans `tests/test_cockpit_ingenieur.py` :

```python
class TestCockpitManifeste:
    def test_version_manifeste_a1(self, tmp_path):
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        assert str(ws["A1"].value).startswith("MANIFESTE_V=")

    def test_ligne2_vide(self, tmp_path):
        """Ligne 2 doit être vide — le parser MXL la skippe."""
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        assert ws["A2"].value is None

    def test_file_type_en_a3(self, tmp_path):
        """A3 doit contenir FILE_TYPE: cockpit_ingenieur."""
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        assert ws["A3"].value == "FILE_TYPE: cockpit_ingenieur"

    def test_commentaires_en_colonne_c(self, tmp_path):
        """Chaque instruction MXL doit avoir un commentaire non vide en colonne C."""
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        # Lignes avec une instruction en col A doivent avoir un commentaire en col C
        for r in range(3, 15):
            instr = ws.cell(row=r, column=1).value
            if instr and str(instr).strip():
                comment = ws.cell(row=r, column=3).value
                assert comment and len(str(comment)) > 5, \
                    f"Commentaire manquant ou trop court en ligne {r}: '{comment}'"

    def test_push_instruction_presente(self, tmp_path):
        """Une instruction PUSH $mes_uos -> ... doit être présente."""
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        instrs = [ws.cell(row=r, column=1).value for r in range(1, 20)]
        assert any(
            str(v).startswith("PUSH $mes_uos") for v in instrs if v
        )

    def test_def_get_table_presente(self, tmp_path):
        """Une instruction DEF $mes_uos = GET_TABLE(...) doit être présente."""
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        instrs = [ws.cell(row=r, column=1).value for r in range(1, 20)]
        assert any(
            str(v).startswith("DEF $mes_uos = GET_TABLE") for v in instrs if v
        )

    def test_colonne_b_non_polluee(self, tmp_path):
        """Col B = ancres uniquement. Les commentaires ne doivent PAS être en col B."""
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        for r in range(3, 15):
            b_val = ws.cell(row=r, column=2).value
            # Col B est vide ou contient une ancre (pas un commentaire long)
            if b_val:
                assert len(str(b_val)) < 60, \
                    f"Col B ligne {r} semble contenir un commentaire : '{b_val}'"

    def test_mxl_parseable_zero_erreurs(self, tmp_path):
        """Le _Manifeste généré doit être parseable par parser.py sans erreur."""
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        from src.parser import parse_file
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        ast = parse_file(path)
        assert ast is not None, "parse_file() a retourné None — pas de feuille _Manifeste"
        errors = [f"L{e.line_num}: {e.message}" for e in ast.errors]
        assert not ast.errors, f"Erreurs de parse MXL : {errors}"
```

- [ ] **Step 2 : Lancer les tests — vérifier les échecs**

```
pytest tests/test_cockpit_ingenieur.py::TestCockpitManifeste -v
```
Expected: plusieurs FAIL (format tabulaire actuel vs MXL attendu).
Le test `test_version_manifeste_a1` passe déjà — c'est normal.

- [ ] **Step 3 : Réécrire `_sheet_manifeste()` dans le générateur**

Dans `src/generators/cockpit_ingenieur_generator.py`, remplacer entièrement la fonction `_sheet_manifeste()` et supprimer la constante `MANIFESTE_HEADERS` :

```python
# Supprimer cette constante (lignes ~28-31) :
# MANIFESTE_HEADERS = ["TYPE","SCOPE",...]

def _sheet_manifeste(wb: Workbook, engineer_name: str, uo_list: List[UOInstance]):
    """Génère l'onglet _Manifeste au format MXL mono-colonne.
    Col A = instruction MXL, col B = ancre, col C = commentaire français.
    """
    ws = wb.create_sheet("_Manifeste")
    ws.sheet_view.showGridLines = False

    safe_name = engineer_name.replace(" ", "_")

    def _mxl_row(row: int, instr: str, ancre: str = "", comment: str = ""):
        cell_a = ws.cell(row=row, column=1, value=instr)
        cell_a.font = body_font(size=9.5, bold=instr.startswith(("DEF ", "PUSH ", "PULL ")))
        cell_a.alignment = left()
        if ancre:
            ws.cell(row=row, column=2, value=ancre).font = body_font(size=9, color="888888")
        if comment:
            c = ws.cell(row=row, column=3, value=comment)
            c.font = body_font(size=9, color="666666")
            c.fill = solid_fill("F9F9F9")

    # ── Ligne 1 : version ────────────────────────────────────────────────────────
    ws["A1"] = "MANIFESTE_V=1"
    ws["A1"].font = body_font(bold=True, color="1F3864")
    # Ligne 2 intentionnellement vide (skippée par parser.py)

    # ── Métadonnées ──────────────────────────────────────────────────────────────
    _mxl_row(3, "FILE_TYPE: cockpit_ingenieur",     comment="Type de fichier ExoSync")
    _mxl_row(4, f"FILE_ID: Cockpit_{safe_name}",    comment="Identifiant unique du cockpit")
    _mxl_row(5, f"ingenieur: {engineer_name}",      comment="Nom de l'ingénieur propriétaire")

    # ── Définition de la table ────────────────────────────────────────────────────
    _mxl_row(7, "DEF $mes_uos = GET_TABLE(Mes UOs, tbl_mes_uos)",
             comment="Référence à la table des UOs de l'ingénieur")
    _mxl_row(8, "COL $mes_uos.avancement : WRITE=engineer",
             comment="% avancement saisi par l'ingénieur (zone jaune)")
    _mxl_row(9, "COL $mes_uos.heures_realisees : WRITE=engineer",
             comment="Heures réalisées saisies par l'ingénieur (zone jaune)")

    # ── Export vers store ────────────────────────────────────────────────────────
    _mxl_row(11, f"PUSH $mes_uos -> cockpit.{safe_name}.mes_uos",
             comment="Remonte les saisies ingénieur vers le store central ExoSync")

    set_column_widths(ws, {"A": 60, "B": 18, "C": 55})
```

**Important :** mettre à jour l'appel à `_sheet_manifeste()` dans `generate_cockpit_ingenieur()`.
Chercher la ligne `_sheet_manifeste(wb, uo_list)` et la remplacer par :
```python
    _sheet_manifeste(wb, engineer_name, uo_list)
```

- [ ] **Step 4 : Lancer les tests**

```
pytest tests/test_cockpit_ingenieur.py::TestCockpitManifeste -v
```
Expected: 8 PASSED

- [ ] **Step 5 : Non-régression complète**

```
python -m pytest -q
```
Expected: 361+ PASS, 0 FAIL.

- [ ] **Step 6 : Commit**

```
git add src/generators/cockpit_ingenieur_generator.py tests/test_cockpit_ingenieur.py
git commit -m "refactor: _Manifeste cockpit en MXL mono-colonne (col A instruction, col C commentaire)"
```

---

## Task 4 — Réécriture `_sheet_manifeste_dashboard()` en MXL mono-colonne (TDD)

**Files:**
- Modify: `tests/test_dashboard_metier.py:156-187` — mettre à jour `TestDashboardManifeste`
- Modify: `src/generators/dashboard_metier_generator.py:332+` — réécrire `_sheet_manifeste_dashboard()`

### Objectif
Remplacer le format tabulaire par MXL mono-colonne avec `LIST DYNAMIC` pour l'auto-découverte des UOs. Le `pilote_id` du dashboard est injecté comme métadonnée pour que le `LIST DYNAMIC WHERE pilote_id=USR004` fonctionne.

Format attendu :
```
A1: MANIFESTE_V=1
A3: FILE_TYPE: dashboard_pilote         C3: Type de fichier ExoSync
A4: FILE_ID: Dashboard_USR004           C4: Identifiant unique du dashboard
A5: pilote_id: USR004                   C5: Identifiant du pilote propriétaire
A7: LIST mes_uos TYPE=uo_instance WHERE pilote_id=USR004   C7: Découverte auto des UOs
A8: COLLECT Activites FROM mes_uos INTO vue_synthese        C8: Agrégation données équipe
```

- [ ] **Step 1 : Réécrire `TestDashboardManifeste` dans `tests/test_dashboard_metier.py`**

Remplacer entièrement la classe `TestDashboardManifeste` :

```python
class TestDashboardManifeste:
    def test_version_manifeste(self, tmp_path):
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        assert str(ws["A1"].value).startswith("MANIFESTE_V=")

    def test_ligne2_vide(self, tmp_path):
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        assert ws["A2"].value is None

    def test_file_type_en_a3(self, tmp_path):
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        assert ws["A3"].value == "FILE_TYPE: dashboard_pilote"

    def test_pilote_id_en_a5(self, tmp_path):
        """A5 doit contenir pilote_id: <id_acteur>."""
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        assert ws["A5"].value == f"pilote_id: {acteur.id}"

    def test_list_dynamic_presente(self, tmp_path):
        """Une instruction LIST mes_uos TYPE=uo_instance WHERE pilote_id=... doit exister."""
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        instrs = [ws.cell(row=r, column=1).value for r in range(1, 20)]
        assert any(
            str(v).startswith("LIST mes_uos TYPE=uo_instance") for v in instrs if v
        )

    def test_collect_presente(self, tmp_path):
        """Une instruction COLLECT ... FROM mes_uos doit exister."""
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        instrs = [ws.cell(row=r, column=1).value for r in range(1, 20)]
        assert any(
            str(v).startswith("COLLECT") and "FROM mes_uos" in str(v)
            for v in instrs if v
        )

    def test_commentaires_en_colonne_c(self, tmp_path):
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        for r in range(3, 12):
            instr = ws.cell(row=r, column=1).value
            if instr and str(instr).strip():
                comment = ws.cell(row=r, column=3).value
                assert comment and len(str(comment)) > 5, \
                    f"Commentaire manquant ligne {r}: '{comment}'"

    def test_mxl_parseable_zero_erreurs(self, tmp_path):
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        from src.parser import parse_file
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        ast = parse_file(path)
        assert ast is not None
        errors = [f"L{e.line_num}: {e.message}" for e in ast.errors]
        assert not ast.errors, f"Erreurs de parse MXL : {errors}"
```

- [ ] **Step 2 : Lancer les tests — vérifier les échecs**

```
pytest tests/test_dashboard_metier.py::TestDashboardManifeste -v
```
Expected: plusieurs FAIL.

- [ ] **Step 3 : Réécrire `_sheet_manifeste_dashboard()` dans le générateur**

Dans `src/generators/dashboard_metier_generator.py`, remplacer entièrement la fonction `_sheet_manifeste_dashboard()` et supprimer la constante `_MANIFESTE_HEADERS` :

```python
# Supprimer la constante _MANIFESTE_HEADERS (chercher et supprimer le bloc)

def _sheet_manifeste_dashboard(wb: Workbook, acteur, uo_list: List[UOInstance]):
    """Génère l'onglet _Manifeste dashboard au format MXL mono-colonne.
    Col A = instruction MXL, col B = ancre, col C = commentaire français.
    Le pilote_id en métadonnée permet à LIST DYNAMIC de découvrir ce dashboard.
    """
    ws = wb.create_sheet("_Manifeste")
    ws.sheet_view.showGridLines = False

    def _mxl_row(row: int, instr: str, ancre: str = "", comment: str = ""):
        cell_a = ws.cell(row=row, column=1, value=instr)
        cell_a.font = body_font(size=9.5, bold=instr.startswith(("DEF ", "PUSH ", "PULL ", "LIST ", "COLLECT ")))
        cell_a.alignment = left()
        if ancre:
            ws.cell(row=row, column=2, value=ancre).font = body_font(size=9, color="888888")
        if comment:
            c = ws.cell(row=row, column=3, value=comment)
            c.font = body_font(size=9, color="666666")
            c.fill = solid_fill("F9F9F9")

    # ── Ligne 1 : version ────────────────────────────────────────────────────────
    ws["A1"] = "MANIFESTE_V=1"
    ws["A1"].font = body_font(bold=True, color="1F3864")
    # Ligne 2 intentionnellement vide

    # ── Métadonnées ──────────────────────────────────────────────────────────────
    _mxl_row(3, "FILE_TYPE: dashboard_pilote",       comment="Type de fichier ExoSync")
    _mxl_row(4, f"FILE_ID: Dashboard_{acteur.id}",   comment="Identifiant unique du dashboard")
    _mxl_row(5, f"pilote_id: {acteur.id}",            comment="Identifiant du pilote propriétaire")

    # ── Découverte automatique des UOs ───────────────────────────────────────────
    _mxl_row(7, f"LIST mes_uos TYPE=uo_instance WHERE pilote_id={acteur.id}",
             comment="Découverte automatique des UOs dont ce pilote est responsable")
    _mxl_row(8, "COLLECT Activites FROM mes_uos INTO vue_synthese",
             comment="Agrégation de toutes les activités des UOs de l'équipe")

    set_column_widths(ws, {"A": 65, "B": 18, "C": 55})
```

**Important :** vérifier et mettre à jour l'appel `_sheet_manifeste_dashboard(wb, acteur, uo_list)` dans `generate_dashboard_metier()` — la signature doit correspondre.

- [ ] **Step 4 : Lancer les tests**

```
pytest tests/test_dashboard_metier.py::TestDashboardManifeste -v
```
Expected: 8 PASSED

- [ ] **Step 5 : Non-régression complète**

```
python -m pytest -q
```
Expected: tous les tests passent.

- [ ] **Step 6 : Commit**

```
git add src/generators/dashboard_metier_generator.py tests/test_dashboard_metier.py
git commit -m "refactor: _Manifeste dashboard en MXL mono-colonne avec LIST DYNAMIC pilote_id"
```

---

## Task 5 — Mise à jour `scripts/demo_cockpits.py` + push final

**Files:**
- Modify: `scripts/demo_cockpits.py` — retirer les vérifications obsolètes basées sur l'ancien format

### Objectif
Vérifier que le script de démo fonctionne encore, adapter les vérifications si nécessaire, et pousser la branche.

- [ ] **Step 1 : Lancer le script de démo**

```
python scripts/demo_cockpits.py
```
Expected: 4 vérifications ✅ (les vérifications portent sur Synthèse/Alertes, pas sur _Manifeste).
Si une vérification échoue, diagnostiquer et corriger dans `scripts/demo_cockpits.py`.

- [ ] **Step 2 : Vérification manuelle dans Excel**

Ouvrir `output/cockpits/Cockpit_Alice_Dubois.xlsx` → onglet `_Manifeste`.
Vérifier :
- A1 = `MANIFESTE_V=1`
- A2 vide
- A3 = `FILE_TYPE: cockpit_ingenieur`
- Col C remplie pour chaque instruction
- Col B vide (pas de commentaires parasites)

Ouvrir `output/cockpits/Dashboard_Jean-Luc_Bernard.xlsx` → onglet `_Manifeste`.
Vérifier :
- A7 = `LIST mes_uos TYPE=uo_instance WHERE pilote_id=USR004`

- [ ] **Step 3 : Suite complète finale**

```
python -m pytest -q
```
Expected: tous les tests passent, zéro échec.

- [ ] **Step 4 : Commit final + push**

```
git add -A
git commit -m "chore: demo_cockpits compatible refonte MXL manifeste"
git push origin master
```

---

## Self-Review du plan

**Couverture spec :**
- ✅ Format MXL mono-colonne → Tasks 3, 4
- ✅ Col C pour commentaires → Tasks 3, 4
- ✅ GET_TABLE pas GET_CELL → Task 3 (tbl_mes_uos + DEF $mes_uos = GET_TABLE)
- ✅ pilotes Dict sur UOInstance → Task 1
- ✅ LIST DYNAMIC WHERE pilote_id → Task 4
- ✅ parser.py parse sans erreur → test `test_mxl_parseable_zero_erreurs` (Tasks 3, 4)

**Risque executor.py :** `resolve_lists` avec `LIST DYNAMIC` appelle `ecosystem.get_files_by_type()`. En test unitaire (sans Exomap peuplée), la liste sera vide — c'est attendu. Le peuplement de l'Exomap se fait lors du sync réel. Aucune tâche supplémentaire requise ici.

**Type consistency :**
- `_sheet_manifeste(wb, engineer_name, uo_list)` → appel mis à jour dans Task 3
- `_sheet_manifeste_dashboard(wb, acteur, uo_list)` → vérification de signature dans Task 4
- `pilotes: Dict[str, str]` défini en Task 1, utilisé nulle part dans les générateurs (c'est correct — il sera injecté dans le MXL quand on génèrera les UO Excel eux-mêmes, hors scope de ce plan)
