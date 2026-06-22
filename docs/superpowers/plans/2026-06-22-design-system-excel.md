# Design System Excel (`xl_design`) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Centraliser toute la charte graphique Excel d'ExoSync dans `src/xl_design.py` (classe `XD`) et migrer les 4 générateurs pour qu'aucun n'ait de style inline.

**Architecture:** Un module unique `XD` portant palette (11 familles × 3 tons), primitives openpyxl, et composants (bannière, en-têtes, tables nommées, spine santé, badges). Les générateurs importent `XD`, déclarent la famille de chaque onglet (`key`), et les invariants s'appliquent. Migration en 3 vagues : (1) construire le module isolément, (2) basculer le style des générateurs sans changer la mise en page (les 382 tests restent verts), (3) ajouter la colonne spine santé feuille par feuille (avec mise à jour des tests de position).

**Tech Stack:** Python 3.14, openpyxl, pytest. Spec : `docs/superpowers/specs/2026-06-22-design-system-excel-design.md`.

**Branche :** `feature/design-system` (déjà créée, spec commitée).

**État de référence :** `python -m pytest tests/ -q` doit afficher le même nombre de `passed` (≈382) à la fin qu'au début, + les nouveaux tests de cette feature.

---

## Structure des fichiers

- **Créer** : `src/xl_design.py` — le module charte (classe `XD`, `SheetStyle`).
- **Créer** : `tests/test_xl_design.py` — tests unitaires du module.
- **Modifier** : `src/generators/cockpit_ingenieur_generator.py` — `src.styles` → `XD` + spine Mes UOs.
- **Modifier** : `src/generators/dashboard_metier_generator.py` — `src.styles` → `XD` + Alertes famille `oil` + spine Synthèse.
- **Modifier** : `projet_TrainSystem/creer_uo.py` — `design_b` → `XD` + bannière à glyphes + spine Activités.
- **Modifier** : `projet_TrainSystem/creer_cockpit_se.py` — helpers inline → `XD`.
- **Modifier** : `tests/test_cockpit_ingenieur.py`, `tests/test_dashboard_metier.py` — mise à jour positions après ajout spine.
- **Conserver tel quel** : `projet_TrainSystem/design_b.py`, `src/styles.py` (retrait différé).

---

## VAGUE 1 — Construire `src/xl_design.py`

### Task 1: Constantes de palette + primitives

**Files:**
- Create: `src/xl_design.py`
- Test: `tests/test_xl_design.py`

- [ ] **Step 1: Écrire le test qui échoue**

```python
# tests/test_xl_design.py
"""Tests TDD pour le module de design system xl_design."""
from openpyxl.styles import Font, PatternFill, Border

from src.xl_design import XD


class TestPrimitives:
    def test_font_family_segoe(self):
        f = XD.fnt(12, bold=True, color="FFFFFF")
        assert isinstance(f, Font)
        assert f.name == "Segoe UI"
        assert f.size == 12
        assert f.bold is True
        assert f.color.rgb.endswith("FFFFFF")

    def test_font_defaut(self):
        f = XD.fnt()
        assert f.size == 10
        assert f.bold is False
        assert f.color.rgb.endswith("2C2C2A")

    def test_fill_solide(self):
        fill = XD.fill("0C447C")
        assert isinstance(fill, PatternFill)
        assert fill.fgColor.rgb.endswith("0C447C")

    def test_input_jaune_constant(self):
        assert XD.INPUT == "FFF2CC"

    def test_hair_border(self):
        assert isinstance(XD.HAIR, Border)
        assert XD.HAIR.left.style == "thin"

    def test_alignements(self):
        assert XD.center().horizontal == "center"
        assert XD.center().wrap_text is True
        assert XD.left().horizontal == "left"
```

- [ ] **Step 2: Lancer le test, vérifier l'échec**

Run: `python -m pytest tests/test_xl_design.py::TestPrimitives -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'src.xl_design'`

- [ ] **Step 3: Créer le module avec constantes + primitives**

```python
# src/xl_design.py
"""xl_design — charte graphique Excel centralisée d'ExoSync (classe XD).

Importé par tous les générateurs : aucun style n'est défini inline ailleurs.
Voir docs/superpowers/specs/2026-06-22-design-system-excel-design.md.
"""
from dataclasses import dataclass

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


@dataclass(frozen=True)
class SheetStyle:
    banner: str   # ton foncé (bannière, texte blanc)
    header: str   # ton moyen (= tabColor + en-tête de tableau)
    accent: str   # ton clair (lignes alternées, cartes)
    glyph: str    # glyphe monochrome de la bannière


class XD:
    FONT_FAMILY = "Segoe UI"
    DEFAULT_TEXT = "2C2C2A"

    # ── Palette transversale (statuts / neutres) ───────────────
    WHITE = "FFFFFF"
    INPUT = "FFF2CC"
    GREEN_L = "EAF3DE"; GREEN_D = "27500A"
    BLUE_L = "E6F1FB";  NAVY_D = "0C447C"
    AMBER_L = "FAEEDA"; AMBER_D = "854F0B"
    RED_L = "FCEBEB";   RED_D = "791F1F"
    GREY_L = "F1EFE8";  GREY_D = "5F5E5A"; GREY_B = "D3D1C7"

    # ── Bordure fine (4 côtés) ─────────────────────────────────
    _SIDE = Side(style="thin", color="D3D1C7")
    HAIR = Border(left=_SIDE, right=_SIDE, top=_SIDE, bottom=_SIDE)

    # ── Primitives ─────────────────────────────────────────────
    @staticmethod
    def fnt(size=10, bold=False, color="2C2C2A", italic=False):
        return Font(name=XD.FONT_FAMILY, size=size, bold=bold, color=color,
                    italic=italic)

    @staticmethod
    def fill(hex_color):
        return PatternFill("solid", fgColor=hex_color)

    @staticmethod
    def center():
        return Alignment(horizontal="center", vertical="center", wrap_text=True)

    @staticmethod
    def left():
        return Alignment(horizontal="left", vertical="center", wrap_text=True)
```

- [ ] **Step 4: Lancer le test, vérifier le succès**

Run: `python -m pytest tests/test_xl_design.py::TestPrimitives -v`
Expected: PASS (6 tests)

- [ ] **Step 5: Commit**

```bash
git add src/xl_design.py tests/test_xl_design.py
git commit -m "feat(xl_design): constantes de palette + primitives"
```

---

### Task 2: Registre `SHEETS` + accès aux familles

**Files:**
- Modify: `src/xl_design.py`
- Test: `tests/test_xl_design.py`

- [ ] **Step 1: Écrire le test qui échoue**

```python
# tests/test_xl_design.py — ajouter
class TestRegistreOnglets:
    def test_onze_familles(self):
        assert len(XD.SHEETS) == 11

    def test_cles_attendues(self):
        attendues = {"general", "dashboard", "description", "planning",
                     "donnees_entree", "activites", "livrables", "oil",
                     "kpi", "orga", "manifeste"}
        assert set(XD.SHEETS) == attendues

    def test_triple_activites(self):
        s = XD.sheet("activites")
        assert s.banner == "084434"
        assert s.header == "0C5E49"
        assert s.accent == "E1F5EE"
        assert s.glyph == "✔"

    def test_triple_oil_rouge(self):
        assert XD.sheet("oil").header == "A32D2D"

    def test_cle_inconnue_leve(self):
        import pytest
        with pytest.raises(KeyError):
            XD.sheet("inexistant")

    def test_tab_colors_mappe_le_ton_moyen(self):
        tc = XD.tab_colors()
        assert tc["general"] == "0C447C"
        assert tc["kpi"] == "534AB7"
```

- [ ] **Step 2: Lancer le test, vérifier l'échec**

Run: `python -m pytest tests/test_xl_design.py::TestRegistreOnglets -v`
Expected: FAIL — `AttributeError: type object 'XD' has no attribute 'SHEETS'`

- [ ] **Step 3: Ajouter le registre et les accesseurs**

Ajouter dans la classe `XD` (après `HAIR`) :

```python
    # ── Registre des familles d'onglets (palette verrouillée) ──
    SHEETS = {
        "general":        SheetStyle("08335E", "0C447C", "E6F1FB", "⬢"),
        "dashboard":      SheetStyle("0E4474", "1763A8", "E3EFFA", "◉"),
        "description":    SheetStyle("1C5E92", "2E86C8", "E7F2FB", "✎"),
        "planning":       SheetStyle("074E60", "0A6E88", "DEEFF3", "◷"),
        "donnees_entree": SheetStyle("0A6149", "0F8A66", "E1F5EE", "⤓"),
        "activites":      SheetStyle("084434", "0C5E49", "E1F5EE", "✔"),
        "livrables":      SheetStyle("386114", "4F8A1E", "EBF3DE", "▣"),
        "oil":            SheetStyle("791F1F", "A32D2D", "FCEBEB", "⚑"),
        "kpi":            SheetStyle("3C3489", "534AB7", "EEEDFE", "▲"),
        "orga":           SheetStyle("4D4C47", "6B6A64", "F1EFE8", "❖"),
        "manifeste":      SheetStyle("1C1C1A", "2C2C2A", "F1EFE8", "⚙"),
    }

    @classmethod
    def sheet(cls, key):
        return cls.SHEETS[key]

    @classmethod
    def tab_colors(cls):
        return {k: v.header for k, v in cls.SHEETS.items()}
```

- [ ] **Step 4: Lancer le test, vérifier le succès**

Run: `python -m pytest tests/test_xl_design.py::TestRegistreOnglets -v`
Expected: PASS (6 tests)

- [ ] **Step 5: Commit**

```bash
git add src/xl_design.py tests/test_xl_design.py
git commit -m "feat(xl_design): registre SHEETS des 11 familles + accesseurs"
```

---

### Task 3: Composant `banner()`

**Files:**
- Modify: `src/xl_design.py`
- Test: `tests/test_xl_design.py`

- [ ] **Step 1: Écrire le test qui échoue**

```python
# tests/test_xl_design.py — ajouter
from openpyxl import Workbook


class TestBanner:
    def _ws(self):
        wb = Workbook()
        return wb.active

    def test_tab_color_pose_ton_moyen(self):
        ws = self._ws()
        XD.banner(ws, "activites", "UO L09U1 — Activités", n_cols=10)
        assert ws.sheet_properties.tabColor.rgb.endswith("0C5E49")

    def test_titre_avec_glyphe_en_a1(self):
        ws = self._ws()
        XD.banner(ws, "activites", "UO L09U1 — Activités", n_cols=10)
        assert "✔" in str(ws["A1"].value)
        assert "Activités" in str(ws["A1"].value)

    def test_fond_banniere_fonce(self):
        ws = self._ws()
        XD.banner(ws, "activites", "T", n_cols=10)
        assert ws["A1"].fill.fgColor.rgb.endswith("084434")

    def test_sous_titre_et_se_a_droite(self):
        ws = self._ws()
        XD.banner(ws, "activites", "T", subtitle="Clim", se="J. Dujardin", n_cols=10)
        valeurs = [ws.cell(row=1, column=c).value for c in range(1, 11)]
        joined = " ".join(str(v) for v in valeurs if v)
        assert "Clim" in joined
        assert "J. Dujardin" in joined

    def test_hauteur_ligne1(self):
        ws = self._ws()
        XD.banner(ws, "activites", "T", n_cols=10, height=30)
        assert ws.row_dimensions[1].height == 30
```

- [ ] **Step 2: Lancer le test, vérifier l'échec**

Run: `python -m pytest tests/test_xl_design.py::TestBanner -v`
Expected: FAIL — `AttributeError: type object 'XD' has no attribute 'banner'`

- [ ] **Step 3: Implémenter `banner()`**

Ajouter dans `XD` :

```python
    @classmethod
    def banner(cls, ws, key, title, subtitle="", se="", n_cols=10, height=30):
        """Bannière 1 ligne : glyphe + titre à gauche, sous-titre · SE à droite.
        Pose aussi tabColor = ton moyen de la famille."""
        s = cls.sheet(key)
        ws.sheet_properties.tabColor = s.header
        for c in range(1, n_cols + 1):
            ws.cell(row=1, column=c).fill = cls.fill(s.banner)

        t = ws.cell(row=1, column=1, value=f"{s.glyph}  {title}")
        t.font = cls.fnt(14, bold=True, color=cls.WHITE)
        t.alignment = Alignment(vertical="center", indent=1)
        left_end = max(n_cols - 3, 1)
        if left_end > 1:
            ws.merge_cells(start_row=1, start_column=1, end_row=1,
                           end_column=left_end)

        right_parts = [p for p in (subtitle, se) if p]
        if right_parts and n_cols > left_end + 1:
            r = ws.cell(row=1, column=left_end + 1,
                        value="   ·   ".join(right_parts))
            r.font = cls.fnt(10, color=cls.WHITE)
            r.alignment = Alignment(horizontal="right", vertical="center",
                                    indent=1)
            ws.merge_cells(start_row=1, start_column=left_end + 1,
                           end_row=1, end_column=n_cols)
        ws.row_dimensions[1].height = height
```

- [ ] **Step 4: Lancer le test, vérifier le succès**

Run: `python -m pytest tests/test_xl_design.py::TestBanner -v`
Expected: PASS (5 tests)

- [ ] **Step 5: Commit**

```bash
git add src/xl_design.py tests/test_xl_design.py
git commit -m "feat(xl_design): composant banner sans navigation"
```

---

### Task 4: `table_header()` + `data_row()`

**Files:**
- Modify: `src/xl_design.py`
- Test: `tests/test_xl_design.py`

- [ ] **Step 1: Écrire le test qui échoue**

```python
# tests/test_xl_design.py — ajouter
class TestTableHeaderEtDataRow:
    def _ws(self):
        return Workbook().active

    def test_header_au_ton_onglet(self):
        ws = self._ws()
        XD.table_header(ws, 5, ["id", "désignation", "statut"], "activites")
        assert ws.cell(row=5, column=1).fill.fgColor.rgb.endswith("0C5E49")
        assert ws.cell(row=5, column=1).font.color.rgb.endswith("FFFFFF")
        assert ws.cell(row=5, column=2).value == "désignation"

    def test_data_row_paire_blanche(self):
        ws = self._ws()
        XD.data_row(ws, 6, 0, 1, 3, "activites")
        assert ws.cell(row=6, column=1).fill.fgColor.rgb.endswith("FFFFFF")

    def test_data_row_impaire_accent(self):
        ws = self._ws()
        XD.data_row(ws, 7, 1, 1, 3, "activites")
        assert ws.cell(row=7, column=1).fill.fgColor.rgb.endswith("E1F5EE")
```

- [ ] **Step 2: Lancer le test, vérifier l'échec**

Run: `python -m pytest tests/test_xl_design.py::TestTableHeaderEtDataRow -v`
Expected: FAIL — `AttributeError: ... has no attribute 'table_header'`

- [ ] **Step 3: Implémenter les deux composants**

Ajouter dans `XD` :

```python
    @classmethod
    def table_header(cls, ws, row, headers, key, col_start=1):
        """En-tête de tableau coloré au ton moyen de l'onglet, texte blanc."""
        s = cls.sheet(key)
        for i, h in enumerate(headers):
            c = ws.cell(row=row, column=col_start + i, value=h)
            c.fill = cls.fill(s.header)
            c.font = cls.fnt(10, bold=True, color=cls.WHITE)
            c.alignment = cls.center()
            c.border = cls.HAIR
        ws.row_dimensions[row].height = 20

    @classmethod
    def data_row(cls, ws, row, i, col_start, col_end, key):
        """Ligne de données : alternance blanc (i pair) / accent (i impair)."""
        s = cls.sheet(key)
        bg = s.accent if i % 2 else cls.WHITE
        for c in range(col_start, col_end + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = cls.fill(bg)
            cell.font = cls.fnt(10)
            cell.border = cls.HAIR
```

- [ ] **Step 4: Lancer le test, vérifier le succès**

Run: `python -m pytest tests/test_xl_design.py::TestTableHeaderEtDataRow -v`
Expected: PASS (3 tests)

- [ ] **Step 5: Commit**

```bash
git add src/xl_design.py tests/test_xl_design.py
git commit -m "feat(xl_design): table_header + data_row"
```

---

### Task 5: `named_table()` — style clair + en-tête manuel

**Files:**
- Modify: `src/xl_design.py`
- Test: `tests/test_xl_design.py`

- [ ] **Step 1: Écrire le test qui échoue**

```python
# tests/test_xl_design.py — ajouter
class TestNamedTable:
    def _ws_avec_donnees(self):
        ws = Workbook().active
        ws["A5"] = "id"; ws["B5"] = "désignation"; ws["C5"] = "statut"
        ws["A6"] = "ACT-1"; ws["B6"] = "x"; ws["C6"] = "A_FAIRE"
        return ws

    def test_table_nommee_creee(self):
        ws = self._ws_avec_donnees()
        XD.named_table(ws, "tbl_test", "A5:C6", "activites")
        assert "tbl_test" in ws.tables

    def test_style_clair(self):
        ws = self._ws_avec_donnees()
        XD.named_table(ws, "tbl_test", "A5:C6", "activites")
        assert ws.tables["tbl_test"].tableStyleInfo.name == "TableStyleLight15"

    def test_entete_colore_manuellement(self):
        ws = self._ws_avec_donnees()
        XD.named_table(ws, "tbl_test", "A5:C6", "activites")
        # ligne 5 = en-tête de la table → ton moyen de l'onglet
        assert ws.cell(row=5, column=1).fill.fgColor.rgb.endswith("0C5E49")
        assert ws.cell(row=5, column=1).font.color.rgb.endswith("FFFFFF")
```

- [ ] **Step 2: Lancer le test, vérifier l'échec**

Run: `python -m pytest tests/test_xl_design.py::TestNamedTable -v`
Expected: FAIL — `AttributeError: ... has no attribute 'named_table'`

- [ ] **Step 3: Implémenter `named_table()`**

Ajouter en tête du fichier l'import :

```python
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.table import Table, TableStyleInfo
```

Ajouter dans `XD` :

```python
    @classmethod
    def named_table(cls, ws, display_name, ref, key):
        """Table Excel nommée (pour GET_TABLE/COLLECT) avec STYLE CLAIR sans
        en-tête imposé + coloration manuelle de l'en-tête au ton de l'onglet.
        L'AutoFilter natif d'Excel est actif automatiquement."""
        s = cls.sheet(key)
        tbl = Table(displayName=display_name, ref=ref)
        tbl.tableStyleInfo = TableStyleInfo(
            name="TableStyleLight15", showRowStripes=True,
            showFirstColumn=False, showLastColumn=False,
            showColumnStripes=False,
        )
        ws.add_table(tbl)
        min_col, min_row, max_col, _ = range_boundaries(ref)
        for c in range(min_col, max_col + 1):
            cell = ws.cell(row=min_row, column=c)
            cell.fill = cls.fill(s.header)
            cell.font = cls.fnt(10, bold=True, color=cls.WHITE)
            cell.alignment = cls.center()
```

- [ ] **Step 4: Lancer le test, vérifier le succès**

Run: `python -m pytest tests/test_xl_design.py::TestNamedTable -v`
Expected: PASS (3 tests)

- [ ] **Step 5: Commit**

```bash
git add src/xl_design.py tests/test_xl_design.py
git commit -m "feat(xl_design): named_table style clair + en-tete manuel"
```

---

### Task 6: `statut_cf()` + `criticite_cf()`

**Files:**
- Modify: `src/xl_design.py`
- Test: `tests/test_xl_design.py`

- [ ] **Step 1: Écrire le test qui échoue**

```python
# tests/test_xl_design.py — ajouter
class TestBadgesConditionnels:
    def test_statut_cf_ajoute_des_regles(self):
        ws = Workbook().active
        XD.statut_cf(ws, "F6:F20")
        assert len(ws.conditional_formatting) >= 1

    def test_criticite_cf_ajoute_des_regles(self):
        ws = Workbook().active
        XD.criticite_cf(ws, "G6:G20")
        assert len(ws.conditional_formatting) >= 1
```

- [ ] **Step 2: Lancer le test, vérifier l'échec**

Run: `python -m pytest tests/test_xl_design.py::TestBadgesConditionnels -v`
Expected: FAIL — `AttributeError: ... has no attribute 'statut_cf'`

- [ ] **Step 3: Implémenter les deux helpers**

Ajouter l'import :

```python
from openpyxl.formatting.rule import CellIsRule
```

Ajouter dans `XD` :

```python
    @classmethod
    def statut_cf(cls, ws, rng):
        """Badges colorés par statut d'activité/livrable."""
        rules = [
            ("TERMINEE", cls.GREEN_L, cls.GREEN_D),
            ("VALIDE",   cls.GREEN_L, cls.GREEN_D),
            ("LIVRE",    cls.BLUE_L,  cls.NAVY_D),
            ("EN_COURS", cls.BLUE_L,  cls.NAVY_D),
            ("A_FAIRE",  cls.GREY_L,  cls.GREY_D),
            ("STAND_BY", cls.AMBER_L, cls.AMBER_D),
        ]
        for val, bg, fg in rules:
            ws.conditional_formatting.add(rng, CellIsRule(
                operator="equal", formula=[f'"{val}"'],
                fill=cls.fill(bg),
                font=cls.fnt(10, bold=True, color=fg)))

    @classmethod
    def criticite_cf(cls, ws, rng):
        """Badges colorés par criticité OIL."""
        rules = [
            ("HAUTE",   cls.RED_L,   cls.RED_D),
            ("MOYENNE", cls.AMBER_L, cls.AMBER_D),
            ("BASSE",   cls.GREEN_L, cls.GREEN_D),
        ]
        for val, bg, fg in rules:
            ws.conditional_formatting.add(rng, CellIsRule(
                operator="equal", formula=[f'"{val}"'],
                fill=cls.fill(bg),
                font=cls.fnt(10, bold=True, color=fg)))
```

- [ ] **Step 4: Lancer le test, vérifier le succès**

Run: `python -m pytest tests/test_xl_design.py::TestBadgesConditionnels -v`
Expected: PASS (2 tests)

- [ ] **Step 5: Commit**

```bash
git add src/xl_design.py tests/test_xl_design.py
git commit -m "feat(xl_design): badges conditionnels statut + criticite"
```

---

### Task 7: `traffic_light()` + `card_border()` + `section_box()`

**Files:**
- Modify: `src/xl_design.py`
- Test: `tests/test_xl_design.py`

- [ ] **Step 1: Écrire le test qui échoue**

```python
# tests/test_xl_design.py — ajouter
class TestTrafficEtCadres:
    def test_traffic_rouge_sous_50(self):
        ws = Workbook().active
        XD.traffic_light(ws, 6, 3, 0.3)
        assert ws.cell(row=6, column=3).fill.fgColor.rgb.endswith("FCEBEB")

    def test_traffic_ambre_entre_50_80(self):
        ws = Workbook().active
        XD.traffic_light(ws, 6, 3, 0.65)
        assert ws.cell(row=6, column=3).fill.fgColor.rgb.endswith("FAEEDA")

    def test_traffic_vert_au_dessus_80(self):
        ws = Workbook().active
        XD.traffic_light(ws, 6, 3, 0.9)
        assert ws.cell(row=6, column=3).fill.fgColor.rgb.endswith("EAF3DE")

    def test_card_border_pose_un_cadre(self):
        ws = Workbook().active
        XD.card_border(ws, 2, 2, 4, 4)
        assert ws.cell(row=2, column=2).border.top.style == "thin"

    def test_section_box_titre_et_fond(self):
        ws = Workbook().active
        XD.section_box(ws, "Titre section", 2, 2, 5, 4, "kpi")
        assert ws.cell(row=2, column=2).value == "Titre section"
```

- [ ] **Step 2: Lancer le test, vérifier l'échec**

Run: `python -m pytest tests/test_xl_design.py::TestTrafficEtCadres -v`
Expected: FAIL — `AttributeError: ... has no attribute 'traffic_light'`

- [ ] **Step 3: Implémenter les trois helpers**

Ajouter dans `XD` :

```python
    @classmethod
    def traffic_light(cls, ws, row, col, value, thresholds=(0.5, 0.8)):
        """Cellule au fond rouge/ambre/vert selon value et les seuils."""
        lo, hi = thresholds
        color = cls.RED_L if value < lo else (cls.AMBER_L if value < hi
                                              else cls.GREEN_L)
        cell = ws.cell(row=row, column=col)
        cell.fill = cls.fill(color)
        cell.border = cls.HAIR
        cell.alignment = cls.center()
        return cell

    @classmethod
    def card_border(cls, ws, r1, c1, r2, c2, color=None):
        """Encadre une zone rectangulaire d'une bordure fine."""
        thin = Side(style="thin", color=color or cls.GREY_B)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                cell = ws.cell(row=r, column=c)
                old = cell.border
                cell.border = Border(
                    left=thin if c == c1 else old.left,
                    right=thin if c == c2 else old.right,
                    top=thin if r == r1 else old.top,
                    bottom=thin if r == r2 else old.bottom,
                )

    @classmethod
    def section_box(cls, ws, title, r1, c1, r2, c2, key):
        """Bande de titre (accent de l'onglet) + cadre fin."""
        s = cls.sheet(key)
        for c in range(c1, c2 + 1):
            ws.cell(row=r1, column=c).fill = cls.fill(s.accent)
        tc = ws.cell(row=r1, column=c1, value=title)
        tc.font = cls.fnt(11, bold=True, color=s.banner)
        tc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[r1].height = 20
        cls.card_border(ws, r1, c1, r2, c2)
```

- [ ] **Step 4: Lancer le test, vérifier le succès**

Run: `python -m pytest tests/test_xl_design.py::TestTrafficEtCadres -v`
Expected: PASS (5 tests)

- [ ] **Step 5: Commit**

```bash
git add src/xl_design.py tests/test_xl_design.py
git commit -m "feat(xl_design): traffic_light + card_border + section_box"
```

---

### Task 8: `health_spine()` — colonne santé par mise en forme conditionnelle

**Files:**
- Modify: `src/xl_design.py`
- Test: `tests/test_xl_design.py`

Note : les règles sont pilotées par la **colonne statut** (sans ambiguïté de
barème, contrairement à l'avancement qui est tantôt 0–1, tantôt 0–100). Le
paramètre `pct_col` est accepté mais réservé pour un raffinement ultérieur.

- [ ] **Step 1: Écrire le test qui échoue**

```python
# tests/test_xl_design.py — ajouter
from openpyxl.utils import get_column_letter


class TestHealthSpine:
    def test_largeur_colonne_fine(self):
        ws = Workbook().active
        XD.health_spine(ws, "activites", header_row=5, row_start=6,
                        row_end=10, status_col=6, spine_col=1)
        assert ws.column_dimensions[get_column_letter(1)].width <= 3

    def test_entete_spine_au_ton_banniere(self):
        ws = Workbook().active
        XD.health_spine(ws, "activites", header_row=5, row_start=6,
                        row_end=10, status_col=6, spine_col=1)
        assert ws.cell(row=5, column=1).fill.fgColor.rgb.endswith("084434")

    def test_regles_conditionnelles_posees(self):
        ws = Workbook().active
        XD.health_spine(ws, "activites", header_row=5, row_start=6,
                        row_end=10, status_col=6, spine_col=1)
        assert len(ws.conditional_formatting) >= 1
```

- [ ] **Step 2: Lancer le test, vérifier l'échec**

Run: `python -m pytest tests/test_xl_design.py::TestHealthSpine -v`
Expected: FAIL — `AttributeError: ... has no attribute 'health_spine'`

- [ ] **Step 3: Implémenter `health_spine()`**

Ajouter les imports :

```python
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
```

Ajouter dans `XD` :

```python
    # Couleurs santé de la spine (indépendantes des familles)
    SPINE_DONE = "3B6D11"; SPINE_OK = "0F8A66"; SPINE_WATCH = "EF9F27"
    SPINE_ALERT = "A32D2D"; SPINE_TODO = "6B6A64"

    @classmethod
    def health_spine(cls, ws, key, header_row, row_start, row_end,
                     status_col, spine_col=1, pct_col=None):
        """Colonne A fine + en-tête au ton bannière. Pose les règles de mise en
        forme conditionnelle santé, lues sur la colonne statut (recolore live)."""
        s = cls.sheet(key)
        spine_L = get_column_letter(spine_col)
        stat_L = get_column_letter(status_col)
        ws.column_dimensions[spine_L].width = 2.5
        ws.cell(row=header_row, column=spine_col).fill = cls.fill(s.banner)

        rng = f"{spine_L}{row_start}:{spine_L}{row_end}"

        def rule(formula, color):
            ws.conditional_formatting.add(rng, FormulaRule(
                formula=[formula], stopIfTrue=True, fill=cls.fill(color)))

        # ordre = priorité (premier vrai gagne)
        rule(f'OR(${stat_L}{row_start}="TERMINEE",${stat_L}{row_start}="VALIDE")',
             cls.SPINE_DONE)
        rule(f'OR(${stat_L}{row_start}="OUVERT",${stat_L}{row_start}="HAUTE")',
             cls.SPINE_ALERT)
        rule(f'${stat_L}{row_start}="STAND_BY"', cls.SPINE_WATCH)
        rule(f'${stat_L}{row_start}="EN_COURS"', cls.SPINE_OK)
        rule(f'OR(${stat_L}{row_start}="A_FAIRE",${stat_L}{row_start}="EN_ATTENTE")',
             cls.SPINE_TODO)
```

- [ ] **Step 4: Lancer le test, vérifier le succès**

Run: `python -m pytest tests/test_xl_design.py::TestHealthSpine -v`
Expected: PASS (3 tests)

- [ ] **Step 5: Vérifier l'ensemble du module + commit**

Run: `python -m pytest tests/test_xl_design.py -q`
Expected: PASS (tous)

```bash
git add src/xl_design.py tests/test_xl_design.py
git commit -m "feat(xl_design): health_spine (mise en forme conditionnelle santé)"
```

---

## VAGUE 2 — Migration du style (mise en page inchangée, 382 tests verts)

> Principe : on remplace les helpers de style inline par `XD`, **sans déplacer
> aucune colonne**. Les tests existants (qui vérifient valeurs/positions) doivent
> rester verts. La colonne spine est ajoutée séparément en Vague 3.

### Task 9: Migrer `creer_cockpit_se.py` → `XD`

**Files:**
- Modify: `projet_TrainSystem/creer_cockpit_se.py`
- Test: génération manuelle (script sans test pytest dédié)

Mapping de couleurs (cet onglet « Mes UOs » = famille `general`) :
`"0C447C"`→`XD.sheet("general").banner` (`08335E`) pour la bannière ;
en-tête `"1F3864"`→`XD.sheet("general").header` ; jaune `"FFF2CC"` conservé via
`XD.INPUT` ; police `Segoe UI` partout (supprime `Calibri`).

- [ ] **Step 1: Remplacer les imports et helpers**

Remplacer les lignes 19-21 et 70-86 (imports openpyxl.styles + helpers
`_fill/_fnt/_center/_left/_thin_border`) par :

```python
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from src.xl_design import XD
```

(supprimer entièrement le bloc `# ─── Helpers style ───` lignes 70-86).

- [ ] **Step 2: Migrer `_sheet_mes_uos`**

Remplacer le corps stylé de `_sheet_mes_uos` (bannière + en-têtes + lignes) en
utilisant `XD` :

```python
def _sheet_mes_uos(wb, se_name, uo_list):
    ws = wb.create_sheet("Mes UOs")
    ws.sheet_view.showGridLines = False

    XD.banner(ws, "general",
              f"Cockpit Ingénieur — {se_name}",
              subtitle=date.today().strftime('%d/%m/%Y'), n_cols=8)

    headers = ["UO ID", "Système", "Projet", "Charge (h)",
               "% Avancement", "H réalisées", "Date fin", "Alerte"]
    row_h = 3
    XD.table_header(ws, row_h, headers, "general")

    for i, uo in enumerate(uo_list):
        row = row_h + 1 + i
        XD.data_row(ws, row, i, 1, 8, "general")
        ws.cell(row=row, column=1, value=uo["file_id"]).font = XD.fnt(color="0563C1")
        ws.cell(row=row, column=2, value=uo["systeme"])
        ws.cell(row=row, column=3, value=uo["projet"])
        ws.cell(row=row, column=4, value=uo["heures"])
        for col in (5, 6):
            c = ws.cell(row=row, column=col, value=0)
            c.fill = XD.fill(XD.INPUT)
            c.border = XD.HAIR
            c.alignment = XD.center()
        ws.cell(row=row, column=5).number_format = "0%"

    last_row = row_h + len(uo_list)
    if uo_list:
        XD.named_table(ws, "tbl_mes_uos", f"A{row_h}:H{last_row}", "general")

    for col, w in zip("ABCDEFGH", [20, 18, 22, 12, 16, 14, 14, 22]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = f"A{row_h + 1}"
```

- [ ] **Step 3: Migrer le `_Manifeste` (supprimer Calibri)**

Dans `_sheet_manifeste`, remplacer la fonction interne `w(...)` :

```python
    def w(row, instr, comment=""):
        c = ws.cell(row=row, column=1, value=instr)
        bold = any(instr.startswith(k) for k in ("DEF ", "PUSH ", "PULL ", "LIST "))
        c.font = XD.fnt(9.5, bold=bold, color="0C447C" if bold else "2C2C2A")
        c.alignment = XD.left()
        if comment:
            ws.cell(row=row, column=3, value=comment).font = XD.fnt(9, color="5F5E5A", italic=True)
```

et `ws["A1"].font = Font(name="Calibri", ...)` →
`ws["A1"].font = XD.fnt(10, bold=True, color="0C447C")`.

Ajouter `ws.sheet_properties.tabColor = XD.sheet("manifeste").header`.

- [ ] **Step 4: Vérifier qu'il n'y a plus de style inline + génération**

Run: `grep -nE "PatternFill|Font\(|Calibri" projet_TrainSystem/creer_cockpit_se.py`
Expected: aucune occurrence (hors imports éventuels inutilisés à retirer).

Run: `python -m pytest tests/ -q`
Expected: même nombre de `passed` qu'avant (≈382), 0 failed.

- [ ] **Step 5: Commit**

```bash
git add projet_TrainSystem/creer_cockpit_se.py
git commit -m "refactor(creer_cockpit_se): style inline -> xl_design"
```

---

### Task 10: Migrer `cockpit_ingenieur_generator.py` → `XD` (style)

**Files:**
- Modify: `src/generators/cockpit_ingenieur_generator.py`
- Test: `tests/test_cockpit_ingenieur.py` (doit rester vert)

Mapping (onglets Mes UOs = `general`, Agenda = `planning`, _Manifeste =
`manifeste`) : `BLUE_DARK`→`general.banner`, `BLUE_MID`→`general.header`,
`BLUE_LIGHT`→`general.accent`, `YELLOW_LIGHT`→`XD.INPUT`, alertes RED/ORANGE/GREEN
conservées via `XD.RED_L/AMBER_L/GREEN_L`.

- [ ] **Step 1: Remplacer le bloc d'import (lignes 12-17)**

```python
from src.xl_design import XD
```

- [ ] **Step 2: Remplacer les appels de style — bannière & en-têtes**

Pour chaque feuille, remplacer le titre fusionné + `solid_fill/header_font` par
`XD.banner(ws, KEY, titre, subtitle=..., n_cols=N)` et chaque ligne d'en-tête par
`XD.table_header(ws, row, headers, KEY)`. Exemple pour `_sheet_mes_uos` (titre
ligne 51-57) :

```python
    XD.banner(ws, "general",
              f"Mes UOs — {engineer_name}",
              subtitle=date.today().strftime('%d/%m/%Y'), n_cols=9)
```

et remplacer `style_header_row(ws, 5, 1, 9, color=BLUE_MID)` par
`XD.table_header(ws, 5, headers, "general")`.

- [ ] **Step 3: Remplacer `style_data_row`, `solid_fill`, fonts par XD**

Substitutions mécaniques dans tout le fichier :
- `style_data_row(ws, r, a, b, alternate=(i % 2 == 1))` → `XD.data_row(ws, r, i, a, b, "general")` (ou `"planning"` dans l'Agenda).
- `solid_fill(YELLOW_LIGHT)` → `XD.fill(XD.INPUT)`.
- `solid_fill(RED_LIGHT)` → `XD.fill(XD.RED_L)`, `ORANGE_LIGHT`→`XD.AMBER_L`, `GREEN_LIGHT`→`XD.GREEN_L`.
- `body_font(...)` → `XD.fnt(...)`, `header_font(...)` → `XD.fnt(bold=True, color=XD.WHITE, ...)`.
- `center()`/`left()` → `XD.center()`/`XD.left()`, `THIN_BORDER` → `XD.HAIR`.
- Table `tbl_mes_uos` : remplacer le bloc `Table(...) + TableStyleInfo(Medium2)` (lignes 152-159) par `XD.named_table(ws, "tbl_mes_uos", tbl_ref, "general")`.

Poser les tabColors : `ws.sheet_properties.tabColor` est déjà mis par `XD.banner`.

- [ ] **Step 4: Vérifier les tests existants + absence de style inline**

Run: `python -m pytest tests/test_cockpit_ingenieur.py -q`
Expected: PASS (tous les tests existants restent verts — positions inchangées).

Run: `grep -nE "PatternFill|from src.styles" src/generators/cockpit_ingenieur_generator.py`
Expected: aucune occurrence.

- [ ] **Step 5: Commit**

```bash
git add src/generators/cockpit_ingenieur_generator.py
git commit -m "refactor(cockpit_ingenieur): src.styles -> xl_design"
```

---

### Task 11: Migrer `dashboard_metier_generator.py` → `XD` (style)

**Files:**
- Modify: `src/generators/dashboard_metier_generator.py`
- Test: `tests/test_dashboard_metier.py` (doit rester vert)

Mapping (Synthèse/Vue Synthèse = `dashboard`, Par Ingénieur = `activites`,
Alertes = `oil`, _Manifeste = `manifeste`).

- [ ] **Step 1: Remplacer le bloc d'import (lignes 12-17) par `from src.xl_design import XD`**

- [ ] **Step 2: Migrer chaque feuille**

- `_sheet_synthese` : titre → `XD.banner(ws, "dashboard", f"Dashboard Métier — {acteur.nom}", subtitle=date.today().strftime('%d/%m/%Y'), n_cols=10)` ; `style_header_row(... BLUE_MID)` → `XD.table_header(ws, 5, headers, "dashboard")` ; `style_data_row` → `XD.data_row(..., "dashboard")`.
- `_sheet_vue_synthese` : `Table + TableStyleInfo(Medium2)` (lignes 183-187) → `XD.named_table(ws, "tbl_vue_synthese", f"A1:{chr(64+len(headers))}2", "dashboard")`.
- `_sheet_par_ingenieur` : bannière `dashboard`, en-têtes par ingénieur → `XD.table_header(..., "activites")`, lignes → `XD.data_row(..., "activites")`.
- `_sheet_alertes` : titre `t.fill = solid_fill("C00000")` → `XD.banner(ws, "oil", f"Alertes & Risques — {date.today():%d/%m/%Y}", n_cols=5)` ; en-tête → `XD.table_header(ws, 2, headers, "oil")` ; fonds d'alerte `RED_LIGHT/ORANGE_LIGHT` → `XD.RED_L/XD.AMBER_L`.
- `_sheet_manifeste_dashboard` : fonts `body_font` → `XD.fnt`, ajouter `ws.sheet_properties.tabColor = XD.sheet("manifeste").header`.

- [ ] **Step 3: Substitutions mécaniques restantes**

`solid_fill`→`XD.fill`, `body_font`→`XD.fnt`, `header_font`→`XD.fnt(bold=True, color=XD.WHITE)`, `center/left`→`XD.center/XD.left`, `THIN_BORDER`→`XD.HAIR`, `GREY_LIGHT`→`XD.GREY_L`, `GREEN_LIGHT`→`XD.GREEN_L`.

- [ ] **Step 4: Vérifier les tests + absence de style inline**

Run: `python -m pytest tests/test_dashboard_metier.py -q`
Expected: PASS (positions inchangées).

Run: `grep -nE "PatternFill|from src.styles" src/generators/dashboard_metier_generator.py`
Expected: aucune occurrence.

- [ ] **Step 5: Commit**

```bash
git add src/generators/dashboard_metier_generator.py
git commit -m "refactor(dashboard_metier): src.styles -> xl_design + Alertes famille oil"
```

---

### Task 12: Migrer `creer_uo.py` → `XD` (style + bannière à glyphes)

**Files:**
- Modify: `projet_TrainSystem/creer_uo.py`
- Test: `python -m pytest tests/ -q` + génération réelle

Mapping des onglets vers les familles `XD` (remplace le dict `TAB` lignes 55-67 et
les `banner_B/banner_teal/banner_amber` de `design_b`) :
General→`general`, Description_Besoin→`description`, Donnees_Entree→`donnees_entree`,
Activites→`activites`, Livrables→`livrables`, OIL→`oil`, KPI→`kpi`,
Dashboard→`dashboard`, Planning→`planning`, Orga→`orga`, _Manifeste→`manifeste`.

- [ ] **Step 1: Remplacer l'import `design_b` (lignes 38-46) par `XD`**

```python
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))
from src.xl_design import XD
```

- [ ] **Step 2: Remplacer le dict `TAB` et `_tab()` par `XD`**

Supprimer le dict `TAB` (lignes 55-67) et `_tab` (lignes 100-103). Remplacer
chaque `banner_X(ws, sous_titre, ncols, **bkw)` + `_tab(ws)` par :

```python
    XD.banner(ws, "<key>", uo_title, subtitle=proj_line, se=se, n_cols=<ncols>)
```

où `<key>` est la famille de l'onglet (table de mapping ci-dessus). La bannière
pose déjà le `tabColor`.

- [ ] **Step 3: Migrer `_write_table` vers `XD`**

Remplacer le corps de `_write_table` (lignes 106-129) par :

```python
def _write_table(ws, name, headers, rows, key, widths=None, start_row=T):
    hr = start_row
    for ci, h in enumerate(headers):
        ws.cell(row=hr, column=ci + 1, value=h)
    for ri, rd in enumerate(rows, hr + 1):
        for ci, h in enumerate(headers, 1):
            c = ws.cell(row=ri, column=ci, value=rd.get(h))
            c.border = XD.HAIR
            c.font = XD.fnt(10)
        ws.row_dimensions[ri].height = 20
    last = hr + max(len(rows), 1)
    XD.named_table(ws, name, f"A{hr}:{get_column_letter(len(headers))}{last}", key)
    if widths:
        for col, w in widths.items():
            ws.column_dimensions[col].width = w
    return hr + 1, last
```

Mettre à jour chaque appel `_write_table(..., header_color=XXX)` →
`_write_table(..., key="<famille>")`.

- [ ] **Step 4: Remplacer les `fnt/fill/statut_cf/criticite_cf/section_box/kpi_card_B` de design_b**

`fnt(...)`→`XD.fnt(...)`, `fill(...)`→`XD.fill(...)`, `statut_cf`→`XD.statut_cf`,
`criticite_cf`→`XD.criticite_cf`, `section_box(ws, t, r1,c1,r2,c2)`→
`XD.section_box(ws, t, r1, c1, r2, c2, "kpi")`. Les constantes `NAVY_D` etc. →
`XD.NAVY_D`/`XD.sheet("general").header`. Pour les cartes KPI du Dashboard,
remplacer `kpi_card_B` par un appel à `XD.card_border` + cellules (garder la même
disposition de colonnes).

- [ ] **Step 5: Générer une UO et vérifier (MXL + ouverture) + tests + commit**

Run: `python projet_TrainSystem/creer_uo.py L09U1-TEST01-CLIM --se "Jean Dujardin" --heures 200 --output projet_TrainSystem`
Expected: `[OK] ...L09U1-TEST01-CLIM.xlsx` sans exception.

Run: `python scripts/valider_un.py projet_TrainSystem/L09U1-TEST01-CLIM.xlsx`
Expected: synchronisation 0 erreur (Manifeste parseable).

Run: `python -m pytest tests/ -q`
Expected: ≈382 passed, 0 failed.

```bash
git add projet_TrainSystem/creer_uo.py
git commit -m "refactor(creer_uo): design_b -> xl_design + banniere a glyphes"
```

---

## VAGUE 3 — Colonne spine santé (déplace les colonnes, met à jour les tests)

> Chaque tâche ajoute la spine en colonne A, décale la table d'une colonne, et
> met à jour le test de position correspondant. À faire après la Vague 2.

### Task 13: Spine sur le cockpit « Mes UOs »

**Files:**
- Modify: `src/generators/cockpit_ingenieur_generator.py` (`_sheet_mes_uos`)
- Test: `tests/test_cockpit_ingenieur.py`

- [ ] **Step 1: Mettre à jour les tests de position (échec attendu)**

Dans `tests/test_cockpit_ingenieur.py`, décaler d'une colonne (les données
commencent désormais en colonne B, en-têtes en B5). Remplacer dans
`test_seules_les_uo_de_alice` : `column=1` → `column=2`. Dans
`test_en_tetes_onglet_mes_uo` : `range(1, 10)` → `range(2, 11)`. Dans
`test_zone_saisie_avancement_col_f` : `column=6/7` → `column=7/8`. Dans
`test_formule_alerte_presente` : `column=9` → `column=10`. Ajouter :

```python
    def test_spine_presente_colonne_a(self, tmp_path):
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["Mes UOs"]
        # colonne A fine
        assert ws.column_dimensions["A"].width <= 3
        # mise en forme conditionnelle posée
        assert len(ws.conditional_formatting) >= 1
```

- [ ] **Step 2: Lancer, vérifier l'échec**

Run: `python -m pytest tests/test_cockpit_ingenieur.py -q`
Expected: FAIL (positions décalées pas encore implémentées).

- [ ] **Step 3: Décaler le tableau de `_sheet_mes_uos` en colonne B + spine**

Réécrire `_sheet_mes_uos` pour : table en `B..J` (au lieu de `A..I`), spine en
colonne A. Points clés :
- En-têtes : `XD.table_header(ws, 5, headers, "general", col_start=2)`.
- Données : écrire en colonnes 2..10 ; statut/alerte décalés (`column=2` = UO ID … `column=10` = Alerte).
- Formule alerte : ajuster les références de colonnes (G→H, E→F, H→I) car tout est décalé de +1.
- Table nommée : `XD.named_table(ws, "tbl_mes_uos", f"B5:J{last_row}", "general")`.
- Spine : `XD.health_spine(ws, "general", header_row=5, row_start=6, row_end=last_row, status_col=10, spine_col=1)` (la colonne « Alerte » sert de statut ; sinon ajouter une colonne statut dédiée masquée).
- Largeurs : `A`=2.5 puis `B..J` = anciennes largeurs.
- `ws.freeze_panes = "B6"`.

- [ ] **Step 4: Lancer les tests, vérifier le succès**

Run: `python -m pytest tests/test_cockpit_ingenieur.py -q`
Expected: PASS (tous, y compris `test_spine_presente_colonne_a`).

- [ ] **Step 5: Commit**

```bash
git add src/generators/cockpit_ingenieur_generator.py tests/test_cockpit_ingenieur.py
git commit -m "feat(cockpit_ingenieur): colonne spine sante (Mes UOs)"
```

---

### Task 14: Spine sur le dashboard « Synthèse »

**Files:**
- Modify: `src/generators/dashboard_metier_generator.py` (`_sheet_synthese`)
- Test: `tests/test_dashboard_metier.py`

- [ ] **Step 1: Mettre à jour / ajouter les tests (échec attendu)**

Adapter dans `tests/test_dashboard_metier.py` les assertions de position de
`_sheet_synthese` (+1 colonne pour les données, table en B). Ajouter :

```python
    def test_spine_synthese(self, tmp_path):
        # ... générer le dashboard comme les autres tests du fichier ...
        ws = wb["Synthèse"]
        assert ws.column_dimensions["A"].width <= 3
        assert len(ws.conditional_formatting) >= 1
```

(Reproduire le motif de génération déjà utilisé dans ce fichier de test.)

- [ ] **Step 2: Lancer, vérifier l'échec**

Run: `python -m pytest tests/test_dashboard_metier.py -q`
Expected: FAIL.

- [ ] **Step 3: Décaler la table de `_sheet_synthese` en colonne B + spine**

- En-têtes : `XD.table_header(ws, 5, headers, "dashboard", col_start=2)`.
- Données UO : colonnes 2..11 (UO ID=2 … Alerte=11).
- Table nommée (si présente) en `B5:...`.
- Spine : `XD.health_spine(ws, "dashboard", header_row=5, row_start=6, row_end=last, status_col=11, spine_col=1)` (colonne Alerte comme source statut).
- `A` largeur 2.5 ; `freeze_panes = "B6"`.

- [ ] **Step 4: Lancer les tests, vérifier le succès**

Run: `python -m pytest tests/test_dashboard_metier.py -q`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/generators/dashboard_metier_generator.py tests/test_dashboard_metier.py
git commit -m "feat(dashboard_metier): colonne spine sante (Synthese)"
```

---

### Task 15: Spine sur l'UO « Activités » (`creer_uo.py`)

**Files:**
- Modify: `projet_TrainSystem/creer_uo.py` (feuille Activites)
- Test: génération réelle + MXL

- [ ] **Step 1: Décaler la table Activités en colonne B + spine**

Dans le bloc `# ── Activites ──` : appeler `_write_table` avec un paramètre de
décalage (`start_col=2`) ou écrire la table à partir de la colonne B. Adapter les
formules `heures_allouees` et `reste_a_faire` (références de colonnes décalées de
+1 : C→D, D→E, E→F, G→H, I→J). Régler `ws.column_dimensions["A"].width = 2.5`.
Puis :

```python
    XD.health_spine(ws, "activites", header_row=T, row_start=T + 1,
                    row_end=T + n, status_col=<col_statut_decalee>, spine_col=1)
```

(`<col_statut_decalee>` = index de la colonne `statut` après décalage.)

- [ ] **Step 2: Générer et vérifier MXL**

Run: `python projet_TrainSystem/creer_uo.py L09U1-TEST01-CLIM --se "Jean Dujardin" --heures 200 --output projet_TrainSystem`
Expected: `[OK]` sans exception.

Run: `python scripts/valider_un.py projet_TrainSystem/L09U1-TEST01-CLIM.xlsx`
Expected: 0 erreur — la table `tbl_activites` garde ses colonnes (la spine est hors table), donc `GET_TABLE` reste correct.

- [ ] **Step 3: Suite complète + commit**

Run: `python -m pytest tests/ -q`
Expected: ≈382 passed, 0 failed.

```bash
git add projet_TrainSystem/creer_uo.py
git commit -m "feat(creer_uo): colonne spine sante (Activites)"
```

---

### Task 16: Vérification finale (critères de succès)

**Files:** aucun (vérification)

- [ ] **Step 1: Suite complète verte**

Run: `python -m pytest tests/ -q`
Expected: ≥382 passed (référence + nouveaux tests `test_xl_design`), 0 failed.

- [ ] **Step 2: Zéro style inline dans les générateurs**

Run: `grep -rnE "PatternFill\(|from src.styles import|Calibri" src/generators/ projet_TrainSystem/creer_uo.py projet_TrainSystem/creer_cockpit_se.py`
Expected: aucune occurrence (hors `design_b.py` conservé).

- [ ] **Step 3: Cohérence visuelle — générer un de chaque**

Run:
```bash
python projet_TrainSystem/creer_uo.py L09U1-TEST01-CLIM --se "Jean Dujardin" --heures 200 --output projet_TrainSystem
python projet_TrainSystem/creer_cockpit_se.py --output projet_TrainSystem
```
Ouvrir une UO + un cockpit dans Excel : mêmes couleurs marine/teal, police Segoe UI
partout, bannières à glyphes, jaune `FFF2CC` identique, colonnes spine colorées.

- [ ] **Step 4: Cocher les critères de succès de la spec**

Relire `docs/superpowers/specs/2026-06-22-design-system-excel-design.md` §11 et
cocher chaque item.

- [ ] **Step 5: Commit final de clôture**

```bash
git add -A
git commit -m "chore(design-system): verification finale CONV-13"
```

---

## Notes d'exécution

- **Risque principal** : la Vague 3 décale des colonnes → toujours mettre à jour
  le test de position **dans la même tâche** que le changement de générateur.
- **Spine hors table nommée** : ne jamais inclure la colonne A dans la plage
  d'une table Excel nommée — sinon `GET_TABLE`/`COLLECT` lit une colonne en trop.
- **Barème d'avancement** : la spine est pilotée par la colonne **statut** (valeurs
  textuelles stables), pas par le pourcentage (0–1 dans les cockpits, 0–100 dans
  les UO).
- **`src/styles.py`** : ne pas supprimer tant que `grep -rn "from src.styles"`
  retourne des résultats hors générateurs migrés. Retrait dans un PR ultérieur.
