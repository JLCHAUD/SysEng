# design_b.py — Module de style Design B (Phase 1) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Extraire les primitives de style de `design_demo.py` dans un module autonome `projet_TrainSystem/design_b.py`, puis refactorer `design_demo.py` pour importer depuis ce module — sans changer le rendu visuel.

**Architecture:** Un seul nouveau fichier `design_b.py` contient palette, helpers, banners et composants. `design_demo.py` supprime ses définitions locales et importe depuis `design_b`. `creer_uo.py` n'est pas touché.

**Tech Stack:** Python 3.x, openpyxl

---

## Fichiers

- **Créer** : `projet_TrainSystem/design_b.py`
- **Modifier** : `projet_TrainSystem/design_demo.py`
- Tests : pas de nouveaux fichiers de test — validation par exécution de `design_demo.py` avant et après.

---

### Task 1 : Capturer la baseline

**Files:**
- Aucun fichier modifié

- [ ] **Step 1 : Générer les fichiers de référence**

```bash
cd C:\Users\fabie\Documents\JLC\Python\SysEng
python projet_TrainSystem/design_demo.py
```

Attendu :
```
[OK] Design_A_Studio.xlsx
[OK] Design_B_Cockpit.xlsx
```

- [ ] **Step 2 : Noter les propriétés clés de Design_B pour validation post-refactor**

```bash
python - << 'EOF'
from openpyxl import load_workbook
wb = load_workbook("projet_TrainSystem/Design_B_Cockpit.xlsx")
print("Feuilles :", wb.sheetnames)
for name in wb.sheetnames:
    ws = wb[name]
    print(f"  {name} tabColor={ws.sheet_properties.tabColor} gridLines={ws.sheet_view.showGridLines}")
print("Tableaux Activites :", list(wb["Activites"].tables.keys()))
print("Tableaux OIL :", list(wb["OIL"].tables.keys()))
print("B1 Dashboard fill :", wb["Dashboard"]["B1"].fill.fgColor.rgb)
EOF
```

Garde les valeurs affichées en tête — elles serviront à la Task 5.

---

### Task 2 : Créer `design_b.py` — palette & imports

**Files:**
- Créer : `projet_TrainSystem/design_b.py`

- [ ] **Step 1 : Créer le fichier avec palette complète**

Contenu exact de `projet_TrainSystem/design_b.py` :

```python
"""
design_b.py — Bibliothèque de style Design B « Cockpit bandeau ».
=================================================================
Palette, primitives et composants réutilisables par creer_uo.py et design_demo.py.
Ne contient aucune donnée de démo.
"""
from openpyxl.chart import DoughnutChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import CellIsRule, DataBarRule, IconSetRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ── Palette ───────────────────────────────────────────────────────────────────
NAVY_D = "0C447C"   # marine foncé — titres, bandeaux
NAVY   = "185FA5"   # marine — accents
BLUE   = "378ADD"   # bleu — data bars
BLUE_L = "E6F1FB"   # bleu pâle — fonds doux
GREY_D = "5F5E5A"   # gris foncé — texte secondaire
GREY_L = "F5F4F0"   # gris chaud pâle — fonds de carte
GREY_B = "D3D1C7"   # gris — bordures
GREEN  = "639922";  GREEN_L  = "EAF3DE";  GREEN_D  = "27500A"
AMBER  = "EF9F27";  AMBER_L  = "FAEEDA";  AMBER_D  = "854F0B"
RED    = "E24B4A";  RED_L    = "FCEBEB";  RED_D    = "791F1F"
WHITE  = "FFFFFF"

# Bandeaux d'onglets
TEAL      = "0F6E56"; TEAL_CHIP  = "085041"; TEAL_TINT  = "9FE1CB"
AMB_B     = "BA7517"; AMB_CHIP   = "854F0B"; AMB_TINT   = "FAC775"

# Typographie & bordures globales
F      = "Segoe UI"
THIN_G = Side(style="thin", color=GREY_B)
HAIR   = Border(left=THIN_G, right=THIN_G, top=THIN_G, bottom=THIN_G)
```

- [ ] **Step 2 : Vérifier que le fichier s'importe sans erreur**

```bash
python -c "import sys; sys.path.insert(0,'projet_TrainSystem'); import design_b; print('OK palette:', design_b.NAVY_D, design_b.TEAL)"
```

Attendu : `OK palette: 0C447C 0F6E56`

---

### Task 3 : Ajouter les primitives de style

**Files:**
- Modifier : `projet_TrainSystem/design_b.py`

- [ ] **Step 1 : Ajouter les 6 fonctions primitives à la suite du fichier**

```python
# ── Primitives ────────────────────────────────────────────────────────────────

def fnt(size=11, bold=False, color="2C2C2A", italic=False):
    return Font(name=F, size=size, bold=bold, color=color, italic=italic)


def fill(color):
    return PatternFill("solid", fgColor=color)


def card_border(ws, r1, c1, r2, c2, side=None, color=GREY_B):
    """Encadre une zone rectangulaire d'une bordure fine, option accent gauche."""
    thin = Side(style="thin", color=color)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            b = {}
            if r == r1: b["top"] = thin
            if r == r2: b["bottom"] = thin
            if c == c1: b["left"] = side or thin
            if c == c2: b["right"] = thin
            old = cell.border
            cell.border = Border(
                left=b.get("left", old.left),
                right=b.get("right", old.right),
                top=b.get("top", old.top),
                bottom=b.get("bottom", old.bottom),
            )


def add_table(ws, name, ref):
    """Crée un tableau Excel nommé avec style Light15 et lignes alternées."""
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight15", showRowStripes=True)
    ws.add_table(t)


def statut_cf(ws, rng):
    """Badges colorés par valeur de statut (colonne statut des activités)."""
    rules = [
        ("TERMINEE",  GREEN_L,  GREEN_D),
        ("EN_COURS",  BLUE_L,   NAVY_D),
        ("A_FAIRE",   GREY_L,   GREY_D),
        ("STAND_BY",  AMBER_L,  AMBER_D),
    ]
    for val, bg, fg in rules:
        ws.conditional_formatting.add(rng, CellIsRule(
            operator="equal",
            formula=[f'"{val}"'],
            fill=fill(bg),
            font=Font(name=F, size=10, bold=True, color=fg),
        ))


def criticite_cf(ws, rng):
    """Badges colorés par criticité OIL (HAUTE/MOYENNE/BASSE)."""
    rules = [
        ("HAUTE",   RED_L,   RED_D),
        ("MOYENNE", AMBER_L, AMBER_D),
        ("BASSE",   GREEN_L, GREEN_D),
    ]
    for val, bg, fg in rules:
        ws.conditional_formatting.add(rng, CellIsRule(
            operator="equal",
            formula=[f'"{val}"'],
            fill=fill(bg),
            font=Font(name=F, size=10, bold=True, color=fg),
        ))
```

- [ ] **Step 2 : Vérifier l'import des primitives**

```bash
python -c "
import sys; sys.path.insert(0,'projet_TrainSystem')
import design_b
from openpyxl import Workbook
wb = Workbook(); ws = wb.active
ws['A1'].fill = design_b.fill(design_b.NAVY_D)
ws['A1'].font = design_b.fnt(12, bold=True, color=design_b.WHITE)
print('OK primitives')
"
```

Attendu : `OK primitives`

---

### Task 4 : Ajouter banners et composants de mise en page

**Files:**
- Modifier : `projet_TrainSystem/design_b.py`

- [ ] **Step 1 : Ajouter les 3 fonctions de bandeau**

```python
# ── Bandeaux ──────────────────────────────────────────────────────────────────

def banner_B(ws, subtitle, ncols, bg=NAVY_D, chip=NAVY, tint="B5D4F4"):
    """Bandeau coloré lignes 1-4, largeur exacte ncols.
    Couleur de bandeau = couleur d'onglet (cohérence visuelle)."""
    ws.sheet_properties.tabColor = bg
    for rr in range(1, 5):
        for cc in range(1, ncols + 1):
            ws.cell(row=rr, column=cc).fill = fill(bg)
    t = ws.cell(row=1, column=2, value="UO L09U1 — Préparation et passage des IDR")
    t.font = Font(name=F, size=14, bold=True, color=WHITE)
    t.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=1, start_column=2, end_row=1,
                   end_column=max(ncols - 1, 3))
    sp = ws.cell(row=2, column=2, value="CLIMATISATION  —  PROJET DEMO")
    sp.font = Font(name=F, size=12, bold=True, color=WHITE)
    sp.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=2, start_column=2, end_row=2,
                   end_column=min(13, max(ncols - 1, 3)))
    nav = [("⌂ Dashboard", "Dashboard"), ("✎ Activités", "Activites"),
           ("⚑ OIL", "OIL")]
    for i, (label, target) in enumerate(nav):
        col = 2 + i * 2 if ncols >= 8 else 2 + i
        c = ws.cell(row=3, column=col)
        c.value = f'=HYPERLINK("#{target}!A1","{label}")'
        c.font = Font(name=F, size=9, bold=True, color=tint)
        if ncols >= 8:
            ws.merge_cells(start_row=3, start_column=col,
                           end_row=3, end_column=col + 1)
    if ncols >= 12:
        s = ws.cell(row=3, column=ncols - 2, value=subtitle + " · J. Dujardin")
        s.font = Font(name=F, size=9, color=tint)
        s.alignment = Alignment(horizontal="right", vertical="center")
        ws.merge_cells(start_row=3, start_column=ncols - 2, end_row=3,
                       end_column=ncols)
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 6


def banner_teal(ws, subtitle, ncols):
    banner_B(ws, subtitle, ncols, bg=TEAL, chip=TEAL_CHIP, tint=TEAL_TINT)


def banner_amber(ws, subtitle, ncols):
    banner_B(ws, subtitle, ncols, bg=AMB_B, chip=AMB_CHIP, tint=AMB_TINT)
```

- [ ] **Step 2 : Ajouter les 3 composants de mise en page**

```python
# ── Composants de mise en page ────────────────────────────────────────────────

def section_box(ws, title, r1, c1, r2, c2):
    """Zone délimitée : bande de titre fond bleu pâle + cadre fin."""
    for cc in range(c1, c2 + 1):
        ws.cell(row=r1, column=cc).fill = fill(BLUE_L)
    tc = ws.cell(row=r1, column=c1, value=title)
    tc.font = fnt(11, bold=True, color=NAVY_D)
    tc.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[r1].height = 20
    card_border(ws, r1, c1, r2, c2, color=GREY_B)


def kpi_card_B(ws, col, label, value, sub, border_color, value_color):
    """Carte KPI 4 colonnes : label + grande valeur + sous-titre + bordure colorée."""
    r = 6
    card_border(ws, r, col, r + 3, col + 2, color=border_color)
    for rr in range(r, r + 4):
        for cc in range(col, col + 3):
            ws.cell(row=rr, column=cc).fill = fill(WHITE)
    card_border(ws, r, col, r + 3, col + 2, color=border_color)
    lab = ws.cell(row=r + 1, column=col, value=label)
    lab.font = fnt(9, bold=True, color=GREY_D)
    lab.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=r + 1, start_column=col, end_row=r + 1,
                   end_column=col + 2)
    val = ws.cell(row=r + 2, column=col, value=value)
    val.font = Font(name=F, size=20, bold=True, color=value_color)
    val.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=r + 2, start_column=col, end_row=r + 2,
                   end_column=col + 2)
    s = ws.cell(row=r + 3, column=col, value=sub)
    s.font = fnt(8.5, color=GREY_D)
    s.alignment = Alignment(horizontal="center", vertical="top")
    ws.merge_cells(start_row=r + 3, start_column=col, end_row=r + 3,
                   end_column=col + 2)


def make_donut(wb, ws_dash, ws_data, data_row, anchor, label, pct, color):
    """Graphique anneau (jauge %) — données dans feuille _chart_data cachée."""
    ws_data.cell(row=data_row, column=1, value="fait")
    ws_data.cell(row=data_row, column=2, value=pct)
    ws_data.cell(row=data_row + 1, column=1, value="reste")
    ws_data.cell(row=data_row + 1, column=2, value=100 - pct)
    chart = DoughnutChart()
    data = Reference(ws_data, min_col=2, min_row=data_row, max_row=data_row + 1)
    chart.add_data(data, titles_from_data=False)
    chart.holeSize = 60
    serie = chart.series[0]
    p1 = DataPoint(idx=0); p1.graphicalProperties.solidFill = color
    p2 = DataPoint(idx=1); p2.graphicalProperties.solidFill = "EFEDE7"
    serie.data_points = [p1, p2]
    chart.legend = None
    chart.width = 4.6
    chart.height = 4.6
    ws_dash.add_chart(chart, anchor)
```

- [ ] **Step 3 : Vérifier que tout le module s'importe sans erreur**

```bash
python -c "
import sys; sys.path.insert(0,'projet_TrainSystem')
import design_b
fns = ['fnt','fill','card_border','add_table','statut_cf','criticite_cf',
       'banner_B','banner_teal','banner_amber','section_box','kpi_card_B','make_donut']
for f in fns:
    assert hasattr(design_b, f), f'MANQUE: {f}'
print('OK —', len(fns), 'fonctions disponibles')
"
```

Attendu : `OK — 12 fonctions disponibles`

---

### Task 5 : Refactorer `design_demo.py`

**Files:**
- Modifier : `projet_TrainSystem/design_demo.py`

- [ ] **Step 1 : Remplacer les imports et la palette en tête de fichier**

Supprimer les lignes 1-59 de `design_demo.py` (commentaire module + imports openpyxl + palette complète) et les remplacer par :

```python
"""
design_demo.py — Démos de design Excel : option A (Studio clair) et B (Cockpit bandeau).
=========================================================================================
Génère deux classeurs de démonstration poussant les capacités natives d'Excel
(sans VBA), pour choisir la direction visuelle des fichiers UO :

  Design_A_Studio.xlsx   minimalisme clair, filets d'accent, data bars vivantes
  Design_B_Cockpit.xlsx  bandeau marine, navigation hyperliens, jauge anneau

Usage : python projet_TrainSystem/design_demo.py
"""
import sys
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from design_b import (
    NAVY_D, NAVY, BLUE, BLUE_L, GREY_D, GREY_L, GREY_B,
    GREEN, GREEN_L, GREEN_D, AMBER, AMBER_L, AMBER_D,
    RED, RED_L, RED_D, WHITE,
    TEAL, TEAL_CHIP, TEAL_TINT, AMB_B, AMB_CHIP, AMB_TINT,
    F, THIN_G, HAIR,
    fnt, fill, card_border, add_table, statut_cf, criticite_cf,
    banner_B, banner_teal, banner_amber,
    section_box, kpi_card_B, make_donut,
)
```

- [ ] **Step 2 : Supprimer les définitions dupliquées dans `design_demo.py`**

Supprimer de `design_demo.py` les blocs suivants (qui sont maintenant dans `design_b.py`) :
- La section `# ── Palette ──` avec toutes les constantes couleur, `F`, `THIN_G`, `HAIR`
- `def fnt(...)`
- `def fill(...)`
- `def card_border(...)`
- `def add_table(...)`
- `def statut_cf(...)`
- `def criticite_cf(...)`
- La section `# DESIGN B — COCKPIT BANDEAU` avec `TEAL`, `AMB_B` et variantes
- `def banner_B(...)`
- `def banner_teal(...)`
- `def banner_amber(...)`
- `def section_box(...)`
- `def make_donut(...)`
- `def kpi_card_B(...)`

Conserver impérativement :
- `HERE = Path(__file__).parent`
- Les constantes de démo : `ACTIVITES`, `OIL`
- `def activites_sheet(...)`
- `def oil_sheet(...)`
- `def kpi_card_A(...)`
- `def build_design_A()`
- `def build_design_B()`
- Le bloc `if __name__ == "__main__":`

- [ ] **Step 3 : Vérifier que design_demo.py s'importe sans erreur**

```bash
python -c "import sys; sys.path.insert(0,'projet_TrainSystem'); import design_demo; print('OK import')"
```

Attendu : `OK import`

---

### Task 6 : Validation finale et commit

**Files:**
- Aucun fichier modifié dans cette task

- [ ] **Step 1 : Générer les deux fichiers Excel après refactor**

```bash
python projet_TrainSystem/design_demo.py
```

Attendu (identique à la baseline Task 1) :
```
[OK] Design_A_Studio.xlsx
[OK] Design_B_Cockpit.xlsx
```

- [ ] **Step 2 : Vérifier les propriétés clés (même commande qu'en Task 1)**

```bash
python - << 'EOF'
from openpyxl import load_workbook
wb = load_workbook("projet_TrainSystem/Design_B_Cockpit.xlsx")
print("Feuilles :", wb.sheetnames)
for name in wb.sheetnames:
    ws = wb[name]
    print(f"  {name} tabColor={ws.sheet_properties.tabColor} gridLines={ws.sheet_view.showGridLines}")
print("Tableaux Activites :", list(wb["Activites"].tables.keys()))
print("Tableaux OIL :", list(wb["OIL"].tables.keys()))
print("B1 Dashboard fill :", wb["Dashboard"]["B1"].fill.fgColor.rgb)
EOF
```

Les valeurs doivent être **identiques** à celles notées en Task 1. Si une propriété diffère, investiguer avant de commiter.

- [ ] **Step 3 : Commit**

```bash
git add projet_TrainSystem/design_b.py projet_TrainSystem/design_demo.py
git commit -m "refactor(design): extraire design_b.py — palette, helpers, banners, composants

Phase 1 : design_demo.py importe depuis design_b.py.
Aucun changement visuel. creer_uo.py non touché.

Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>"
```

Attendu : commit créé sur `master`.

---

## Résumé des fichiers après phase 1

| Fichier | Rôle |
|---------|------|
| `projet_TrainSystem/design_b.py` | Bibliothèque de style : 12 fonctions + palette |
| `projet_TrainSystem/design_demo.py` | Démo uniquement : données statiques + builders A/B |
| `projet_TrainSystem/creer_uo.py` | Inchangé — sera la cible de la phase 2 |
