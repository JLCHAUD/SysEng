# Cockpits Ingénieur & Dashboard Métier — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Créer deux nouveaux générateurs Excel — cockpit ingénieur agenda (Agenda + Mes UOs + _Manifeste) et dashboard métier consolidé (Synthèse + Par Ingénieur + Alertes + _Manifeste) — avec push/pull via le store JSON ExoSync natif.

**Architecture:** Les cockpits ingénieurs pushent avancement/heures vers `store.json` via leur `_Manifeste`. Le dashboard métier pulle ces valeurs depuis `store.json`. Le filtrage des ingénieurs visibles par le pilote métier utilise `ProfilActeur.filtre_valeur` (liste CSV de noms). Aucun fichier existant n'est modifié.

**Tech Stack:** Python 3.11+, openpyxl, pytest, `src/store.py` (JsonStore), `src/styles.py`, `src/models.py` (UOInstance, ProfilActeur, TypeFiltre)

---

## Structure des fichiers

| Fichier | Action | Responsabilité |
|---------|--------|----------------|
| `src/generators/cockpit_ingenieur_generator.py` | Créer | Génère Agenda + Mes UOs + _Manifeste pour un ingénieur |
| `src/generators/dashboard_metier_generator.py` | Créer | Génère Synthèse + Par Ingénieur + Alertes + _Manifeste pour un pilote métier |
| `tests/test_cockpit_ingenieur.py` | Créer | Tests TDD pour le cockpit ingénieur |
| `tests/test_dashboard_metier.py` | Créer | Tests TDD pour le dashboard métier |

Fichiers existants **non modifiés** : `cockpit_generator.py`, `consolidation_generator.py`, tous les tests existants.

---

## Task 1 : Tests cockpit ingénieur (TDD — écrire d'abord)

**Files:**
- Create: `tests/test_cockpit_ingenieur.py`

- [ ] **Step 1 : Écrire le fichier de tests complet**

```python
# tests/test_cockpit_ingenieur.py
"""Tests TDD pour cockpit_ingenieur_generator."""
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

from src.models import UOInstance, UOType, Activity, System, Project, StatutUO


def _make_uo(uid: str, engineer: str, hours: float, end: date,
             uo_type_id: str = "spec_fonctionnelle") -> UOInstance:
    activities = [
        Activity(id="ACT-1", name="Analyse", default_hours=hours * 0.4),
        Activity(id="ACT-2", name="Rédaction", default_hours=hours * 0.6),
    ]
    uo_type = UOType(id=uo_type_id, name=f"Type {uo_type_id}", activities=activities)
    system = System(id="clim", name="Climatisation")
    project = Project(id="MI20", name="MI20 RATP")
    return UOInstance(
        id=uid, uo_type_id=uo_type_id, system_id="clim", project_id="MI20",
        engineer_name=engineer, total_hours=hours,
        start_date=date(2026, 4, 1), end_date=end,
        statut=StatutUO.EN_COURS,
        uo_type=uo_type, system=system, project=project,
    )


ALL_UOS = [
    _make_uo("UO-001", "Alice Dubois",  32, date(2026, 6, 30)),
    _make_uo("UO-002", "Alice Dubois",  48, date(2026, 7, 15)),
    _make_uo("UO-003", "Bruno Lecomte", 40, date(2026, 6, 20)),
]


class TestCockpitIngenieurFichier:
    def test_fichier_cree(self, tmp_path):
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        assert path.exists()
        assert path.name == "Cockpit_Alice_Dubois.xlsx"

    def test_trois_onglets_presents(self, tmp_path):
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        assert "Agenda" in wb.sheetnames
        assert "Mes UOs" in wb.sheetnames
        assert "_Manifeste" in wb.sheetnames


class TestCockpitMesUOs:
    def test_seules_les_uo_de_alice(self, tmp_path):
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["Mes UOs"]
        # Les IDs dans la colonne A (données à partir de la ligne 6)
        ids = [ws.cell(row=r, column=1).value for r in range(6, 20) if ws.cell(row=r, column=1).value]
        assert "UO-001" in ids
        assert "UO-002" in ids
        assert "UO-003" not in ids  # Bruno, pas Alice

    def test_en_tetes_onglet_mes_uo(self, tmp_path):
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["Mes UOs"]
        headers = [ws.cell(row=5, column=c).value for c in range(1, 10)]
        assert "UO ID" in headers
        assert "% Avancement" in headers
        assert "H réalisées" in headers
        assert "Alerte" in headers

    def test_zone_saisie_avancement_col_f(self, tmp_path):
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["Mes UOs"]
        # En-tête colonne F = "% Avancement"
        assert ws.cell(row=5, column=6).value == "% Avancement"
        # En-tête colonne G = "H réalisées"
        assert ws.cell(row=5, column=7).value == "H réalisées"

    def test_formule_alerte_presente(self, tmp_path):
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["Mes UOs"]
        # Colonne I = Alerte, ligne 6 = première donnée
        alerte_cell = ws.cell(row=6, column=9).value
        assert alerte_cell is not None
        assert str(alerte_cell).startswith("=IF(")


class TestCockpitAgenda:
    def test_en_tetes_onglet_agenda(self, tmp_path):
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["Agenda"]
        # Cherche "UO ID" et "Activité" dans les premières lignes
        all_values = [ws.cell(row=r, column=c).value for r in range(1, 15) for c in range(1, 7)]
        assert "UO ID" in all_values
        assert "Activité" in all_values

    def test_section_semaine_presente(self, tmp_path):
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["Agenda"]
        all_values = [ws.cell(row=r, column=c).value for r in range(1, 30) for c in range(1, 4)]
        assert any("Semaine" in str(v) for v in all_values if v)

    def test_section_prochaines_echeances_presente(self, tmp_path):
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["Agenda"]
        all_values = [ws.cell(row=r, column=c).value for r in range(1, 50) for c in range(1, 4)]
        assert any("Prochaines" in str(v) or "échéance" in str(v).lower() for v in all_values if v)


class TestCockpitManifeste:
    def test_version_manifeste_a1(self, tmp_path):
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        assert str(ws["A1"].value).startswith("MANIFESTE_V=")

    def test_colonne_commentaire_presente(self, tmp_path):
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        headers = [ws.cell(row=2, column=c).value for c in range(1, 15)]
        assert "COMMENTAIRE" in headers

    def test_regles_push_avancement_presentes(self, tmp_path):
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        nom_globals = [ws.cell(row=r, column=3).value for r in range(3, 20) if ws.cell(row=r, column=3).value]
        assert "uo.UO-001.avancement" in nom_globals
        assert "uo.UO-002.avancement" in nom_globals

    def test_commentaires_non_vides(self, tmp_path):
        from src.generators.cockpit_ingenieur_generator import generate_cockpit_ingenieur
        path = generate_cockpit_ingenieur("Alice Dubois", ALL_UOS, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        # Trouver index de la colonne COMMENTAIRE
        headers = {ws.cell(row=2, column=c).value: c for c in range(1, 15)}
        col_c = headers.get("COMMENTAIRE")
        assert col_c is not None
        # Toutes les lignes de données doivent avoir un commentaire
        for r in range(3, 20):
            if ws.cell(row=r, column=1).value:  # ligne non vide
                comment = ws.cell(row=r, column=col_c).value
                assert comment and len(str(comment)) > 10, f"Commentaire manquant ligne {r}"
```

- [ ] **Step 2 : Vérifier que les tests échouent (module absent)**

```
pytest tests/test_cockpit_ingenieur.py -q
```

Résultat attendu : `ModuleNotFoundError: No module named 'src.generators.cockpit_ingenieur_generator'`

---

## Task 2 : Générateur cockpit ingénieur — onglet `Mes UOs`

**Files:**
- Create: `src/generators/cockpit_ingenieur_generator.py`

- [ ] **Step 1 : Créer le générateur avec la fonction principale et l'onglet `Mes UOs`**

```python
# src/generators/cockpit_ingenieur_generator.py
"""Cockpit ingénieur système — Vue Agenda + Mes UOs + _Manifeste."""
from datetime import date, timedelta
from pathlib import Path
from typing import List

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule

from src.models import UOInstance
from src.styles import (
    BLUE_DARK, BLUE_MID, BLUE_LIGHT, GREEN_LIGHT, ORANGE_LIGHT, RED_LIGHT,
    YELLOW_LIGHT, WHITE, GREY_LIGHT, THIN_BORDER,
    solid_fill, header_font, body_font, center, left,
    style_header_row, style_data_row, set_column_widths, freeze_top_row,
)

OUTPUT_DIR = Path(__file__).parent.parent.parent / "output" / "cockpits"
UO_DIR = Path(__file__).parent.parent.parent / "output" / "UOs"


def generate_cockpit_ingenieur(
    engineer_name: str,
    all_uo_instances: List[UOInstance],
    output_dir: Path = OUTPUT_DIR,
) -> Path:
    """Génère le cockpit agenda pour un ingénieur système.

    Args:
        engineer_name: Nom complet de l'ingénieur (ex: "Alice Dubois")
        all_uo_instances: Toutes les UOInstances — sera filtré sur engineer_name
        output_dir: Répertoire de sortie

    Returns:
        Chemin du fichier Excel généré
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    uo_list = [u for u in all_uo_instances if u.engineer_name == engineer_name]

    wb = Workbook()
    wb.remove(wb.active)  # supprime la feuille par défaut

    _sheet_mes_uos(wb, engineer_name, uo_list)
    _sheet_agenda(wb, engineer_name, uo_list)
    _sheet_manifeste(wb, uo_list)

    safe = engineer_name.replace(" ", "_")
    filepath = output_dir / f"Cockpit_{safe}.xlsx"
    wb.save(filepath)
    return filepath


def _sheet_mes_uos(wb: Workbook, engineer_name: str, uo_list: List[UOInstance]):
    ws = wb.create_sheet("Mes UOs")
    ws.sheet_view.showGridLines = False

    # ── Titre ──────────────────────────────────────────────────────────────────
    ws.merge_cells("A1:I1")
    t = ws["A1"]
    t.value = f"Mes UOs — {engineer_name}   |   {date.today().strftime('%d/%m/%Y')}"
    t.fill = solid_fill(BLUE_DARK)
    t.font = header_font(size=13)
    t.alignment = center()
    ws.row_dimensions[1].height = 30

    # ── KPI ────────────────────────────────────────────────────────────────────
    ws.merge_cells("A2:C2")
    ws["A2"].value = f"UOs actives : {len(uo_list)}"
    ws["A2"].fill = solid_fill(BLUE_LIGHT)
    ws["A2"].font = body_font(bold=True)
    ws["A2"].alignment = center()
    ws["A2"].border = THIN_BORDER

    ws.merge_cells("D2:F2")
    ws["D2"].value = f"Charge totale : {sum(u.total_hours for u in uo_list)}h"
    ws["D2"].fill = solid_fill(BLUE_LIGHT)
    ws["D2"].font = body_font(bold=True)
    ws["D2"].alignment = center()
    ws["D2"].border = THIN_BORDER

    # ── Section titre ──────────────────────────────────────────────────────────
    ws.merge_cells("A4:I4")
    sec = ws["A4"]
    sec.value = "Toutes mes UO"
    sec.fill = solid_fill(BLUE_MID)
    sec.font = header_font()
    sec.alignment = center()

    # ── En-têtes colonnes ──────────────────────────────────────────────────────
    # A=UO ID  B=Type  C=Système  D=Projet  E=Charge(h)  F=% Avancement  G=H réalisées  H=Date fin  I=Alerte
    headers = ["UO ID", "Type UO", "Système", "Projet", "Charge (h)",
               "% Avancement", "H réalisées", "Date fin", "Alerte"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=5, column=col, value=h)
    style_header_row(ws, 5, 1, 9, color=BLUE_MID)

    # ── Données ────────────────────────────────────────────────────────────────
    for i, uo in enumerate(uo_list):
        row = 6 + i
        type_name = uo.uo_type.name if uo.uo_type else uo.uo_type_id
        sys_name = uo.system.name if uo.system else uo.system_id
        proj_name = uo.project.name if uo.project else uo.project_id

        # UO ID avec hyperlink
        id_cell = ws.cell(row=row, column=1, value=uo.id)
        id_cell.font = body_font(color="0563C1")
        id_cell.alignment = left()
        id_cell.border = THIN_BORDER
        uo_path = UO_DIR / f"{uo.id}.xlsx"
        id_cell.hyperlink = str(uo_path)

        ws.cell(row=row, column=2, value=type_name)
        ws.cell(row=row, column=3, value=sys_name)
        ws.cell(row=row, column=4, value=proj_name)
        ws.cell(row=row, column=5, value=uo.total_hours)

        # Zone de saisie — col F (% Avancement) et G (H réalisées) — fond jaune
        avanc_cell = ws.cell(row=row, column=6, value=0.0)
        avanc_cell.number_format = "0%"
        avanc_cell.fill = solid_fill(YELLOW_LIGHT)
        avanc_cell.border = THIN_BORDER
        avanc_cell.alignment = center()

        h_cell = ws.cell(row=row, column=7, value=0.0)
        h_cell.fill = solid_fill(YELLOW_LIGHT)
        h_cell.border = THIN_BORDER
        h_cell.alignment = center()

        date_cell = ws.cell(row=row, column=8, value=uo.end_date)
        date_cell.number_format = "DD/MM/YYYY"
        date_cell.border = THIN_BORDER
        date_cell.alignment = center()

        # Formule alerte : dérive heures OU échéance proche
        alert_formula = (
            f'=IF(G{row}>E{row},"⚠ Dérive heures",'
            f'IF(AND(H{row}<>"",H{row}<TODAY()+7),"⏰ Échéance proche","✅ OK"))'
        )
        ws.cell(row=row, column=9, value=alert_formula)

        style_data_row(ws, row, 2, 5, alternate=(i % 2 == 1))
        style_data_row(ws, row, 8, 9, alternate=(i % 2 == 1))

    # ── Mise en forme conditionnelle Alerte ───────────────────────────────────
    last_row = 5 + len(uo_list)
    if uo_list:
        alert_range = f"I6:I{last_row}"
        ws.conditional_formatting.add(
            alert_range,
            CellIsRule(operator="equal", formula=['"⚠ Dérive heures"'],
                       fill=solid_fill(RED_LIGHT)),
        )
        ws.conditional_formatting.add(
            alert_range,
            CellIsRule(operator="equal", formula=['"⏰ Échéance proche"'],
                       fill=solid_fill(ORANGE_LIGHT)),
        )
        ws.conditional_formatting.add(
            alert_range,
            CellIsRule(operator="equal", formula=['"✅ OK"'],
                       fill=solid_fill(GREEN_LIGHT)),
        )

    set_column_widths(ws, {
        "A": 12, "B": 30, "C": 18, "D": 20, "E": 13,
        "F": 16, "G": 14, "H": 14, "I": 22,
    })
    ws.freeze_panes = "A6"
```

- [ ] **Step 2 : Lancer les tests `Mes UOs` pour vérifier qu'ils passent**

```
pytest tests/test_cockpit_ingenieur.py::TestCockpitMesUOs -v
```

Résultat attendu : 4 tests PASS

- [ ] **Step 3 : Lancer les tests fichier**

```
pytest tests/test_cockpit_ingenieur.py::TestCockpitIngenieurFichier -v
```

Résultat attendu : `test_fichier_cree` PASS, `test_trois_onglets_presents` FAIL (Agenda et _Manifeste pas encore créés)

---

## Task 3 : Générateur cockpit ingénieur — onglet `Agenda`

**Files:**
- Modify: `src/generators/cockpit_ingenieur_generator.py`

- [ ] **Step 1 : Ajouter la fonction `_sheet_agenda` dans le générateur**

Ajouter après `_sheet_mes_uos` dans le fichier :

```python
def _sheet_agenda(wb: Workbook, engineer_name: str, uo_list: List[UOInstance]):
    ws = wb.create_sheet("Agenda")
    ws.sheet_view.showGridLines = False
    today = date.today()

    # ── Titre ──────────────────────────────────────────────────────────────────
    ws.merge_cells("A1:F1")
    t = ws["A1"]
    t.value = f"Agenda — {engineer_name}   |   Semaine du {today.strftime('%d/%m/%Y')}"
    t.fill = solid_fill(BLUE_DARK)
    t.font = header_font(size=13)
    t.alignment = center()
    ws.row_dimensions[1].height = 30

    current_row = 3

    # ── Section : Semaine en cours ────────────────────────────────────────────
    current_row = _agenda_section(
        ws, "📅  Semaine en cours", current_row,
        uo_list, today, today + timedelta(days=7),
        color=BLUE_MID,
    )

    current_row += 1

    # ── Section : Prochaines échéances (8-30j) ────────────────────────────────
    current_row = _agenda_section(
        ws, "📋  Prochaines échéances (30 jours)", current_row,
        uo_list, today + timedelta(days=8), today + timedelta(days=30),
        color=BLUE_LIGHT, header_text_color="1F3864",
    )

    current_row += 1

    # ── Section : Points ouverts ──────────────────────────────────────────────
    current_row = _agenda_points_ouverts(ws, current_row, uo_list)

    set_column_widths(ws, {
        "A": 12, "B": 35, "C": 12, "D": 16, "E": 18, "F": 30,
    })


def _agenda_section(
    ws, title: str, start_row: int,
    uo_list: List[UOInstance], date_from: date, date_to: date,
    color: str = BLUE_MID, header_text_color: str = "FFFFFF",
) -> int:
    """Affiche une section agenda filtrée par horizon de dates. Retourne la prochaine ligne libre."""
    # Titre de section
    ws.merge_cells(f"A{start_row}:F{start_row}")
    sec = ws[f"A{start_row}"]
    sec.value = title
    sec.fill = solid_fill(color)
    sec.font = header_font(color=header_text_color)
    sec.alignment = left()
    start_row += 1

    # En-têtes
    headers = ["UO ID", "Activité", "Priorité", "Date échéance", "Statut", "Action"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=start_row, column=col, value=h)
    style_header_row(ws, start_row, 1, 6, color=BLUE_LIGHT)
    for col in range(1, 7):
        ws.cell(row=start_row, column=col).font = body_font(bold=True, color="1F3864")
    start_row += 1

    # Activités dans la fenêtre de dates
    row = start_row
    for uo in uo_list:
        activities = uo.uo_type.activities if uo.uo_type else []
        for idx, act in enumerate(activities):
            # Utilise uo.end_date comme fallback si l'activité n'a pas de date
            act_end = act.end_date if act.end_date else uo.end_date
            if not act_end or not (date_from <= act_end <= date_to):
                continue

            ws.cell(row=row, column=1, value=uo.id)
            ws.cell(row=row, column=2, value=act.name)

            # Priorité : Haute si echéance < 3j, Normale sinon
            priorite = "🔴 Haute" if act_end <= date.today() + timedelta(days=3) else "🟡 Normale"
            ws.cell(row=row, column=3, value=priorite)

            date_cell = ws.cell(row=row, column=4, value=act_end)
            date_cell.number_format = "DD/MM/YYYY"

            statut_val = act.statut.value if hasattr(act.statut, "value") else str(act.statut)
            ws.cell(row=row, column=5, value=statut_val)
            ws.cell(row=row, column=6, value="")  # zone saisie libre

            style_data_row(ws, row, 1, 6, alternate=(row % 2 == 0))
            row += 1

    # Si aucune activité dans cette fenêtre
    if row == start_row:
        ws.merge_cells(f"A{row}:F{row}")
        ws[f"A{row}"].value = "Aucune activité dans cette période"
        ws[f"A{row}"].fill = solid_fill("F9F9F9")
        ws[f"A{row}"].font = body_font(color="999999")
        ws[f"A{row}"].alignment = center()
        row += 1

    return row


def _agenda_points_ouverts(ws, start_row: int, uo_list: List[UOInstance]) -> int:
    """Section Points ouverts en bas de l'Agenda."""
    ws.merge_cells(f"A{start_row}:F{start_row}")
    sec = ws[f"A{start_row}"]
    sec.value = "⚡  Points ouverts / Actions"
    sec.fill = solid_fill(ORANGE_LIGHT)
    sec.font = header_font(color="C00000")
    sec.alignment = left()
    start_row += 1

    headers = ["UO ID", "Description action", "Responsable", "Date limite", "Nb points", "Statut"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=start_row, column=col, value=h)
    style_header_row(ws, start_row, 1, 6, color=BLUE_LIGHT)
    for col in range(1, 7):
        ws.cell(row=start_row, column=col).font = body_font(bold=True, color="1F3864")
    start_row += 1

    # Ligne vide saisissable par UO
    for uo in uo_list:
        ws.cell(row=start_row, column=1, value=uo.id)
        for col in range(2, 7):
            c = ws.cell(row=start_row, column=col, value="")
            c.fill = solid_fill(YELLOW_LIGHT)
            c.border = THIN_BORDER
        ws.cell(row=start_row, column=1).border = THIN_BORDER
        ws.cell(row=start_row, column=1).alignment = left()
        start_row += 1

    return start_row
```

- [ ] **Step 2 : Lancer les tests Agenda**

```
pytest tests/test_cockpit_ingenieur.py::TestCockpitAgenda -v
```

Résultat attendu : 3 tests PASS

---

## Task 4 : Générateur cockpit ingénieur — onglet `_Manifeste`

**Files:**
- Modify: `src/generators/cockpit_ingenieur_generator.py`

- [ ] **Step 1 : Ajouter la fonction `_sheet_manifeste`**

Ajouter après `_agenda_points_ouverts` dans le fichier :

```python
MANIFESTE_HEADERS = [
    "TYPE", "SCOPE", "NOM_GLOBAL", "NOM_LOCAL",
    "FEUILLE", "TABLEAU", "CLE", "COLONNES",
    "CELLULE", "DIRECTION", "FORMULE", "COMMENTAIRE",
]


def _sheet_manifeste(wb: Workbook, uo_list: List[UOInstance]):
    ws = wb.create_sheet("_Manifeste")
    ws.sheet_view.showGridLines = False

    # ── Version ────────────────────────────────────────────────────────────────
    ws["A1"] = "MANIFESTE_V=1"
    ws["A1"].font = body_font(bold=True, color="1F3864")

    # ── En-têtes ───────────────────────────────────────────────────────────────
    for col, h in enumerate(MANIFESTE_HEADERS, 1):
        ws.cell(row=2, column=col, value=h)
    style_header_row(ws, 2, 1, len(MANIFESTE_HEADERS), color=BLUE_MID)

    # ── Règles : une ligne PUSH par UO × 2 champs (avancement + heures)
    #            une ligne PULL par UO × 2 champs (charge_allouee + date_fin)
    row = 3
    for i, uo in enumerate(uo_list):
        mes_uos_data_row = 6 + i  # correspond à la ligne dans "Mes UOs"

        rules = [
            # TYPE         SCOPE     NOM_GLOBAL                          NOM_LOCAL  FEUILLE    TAB  CLE  COL  CELLULE                 DIRECTION  FORMULE  COMMENTAIRE
            ("CELL_PCT",  "GLOBAL", f"uo.{uo.id}.avancement",           "",        "Mes UOs", "",  "",  "",  f"F{mes_uos_data_row}",  "PUSH",    "",
             "Remonte le % d'avancement saisi par l'ingénieur vers le store central"),
            ("CELL_NUM",  "GLOBAL", f"uo.{uo.id}.heures_realisees",     "",        "Mes UOs", "",  "",  "",  f"G{mes_uos_data_row}",  "PUSH",    "",
             "Remonte les heures réalisées saisies par l'ingénieur vers le store central"),
            ("CELL_NUM",  "GLOBAL", f"uo.{uo.id}.charge_allouee",       "",        "Mes UOs", "",  "",  "",  f"E{mes_uos_data_row}",  "PULL",    "",
             "Injecte la charge allouée depuis le store (valeur de référence, lecture seule)"),
            ("CELL_DATE", "GLOBAL", f"uo.{uo.id}.date_fin",             "",        "Mes UOs", "",  "",  "",  f"H{mes_uos_data_row}",  "PULL",    "",
             "Injecte la date de fin planifiée depuis le store (lecture seule)"),
        ]

        for rule in rules:
            for col, val in enumerate(rule, 1):
                c = ws.cell(row=row, column=col, value=val)
                c.font = body_font(size=10)
                c.border = THIN_BORDER
                c.alignment = left()
            # Colonne COMMENTAIRE en italique gris
            ws.cell(row=row, column=12).font = body_font(size=10, color="666666")
            ws.cell(row=row, column=12).fill = solid_fill("F7F7F7")
            row += 1

        # Ligne vide entre UOs pour lisibilité
        row += 1

    set_column_widths(ws, {
        "A": 12, "B": 8, "C": 35, "D": 14, "E": 12, "F": 10,
        "G": 8, "H": 10, "I": 10, "J": 10, "K": 10, "L": 55,
    })
```

- [ ] **Step 2 : Lancer tous les tests cockpit ingénieur**

```
pytest tests/test_cockpit_ingenieur.py -v
```

Résultat attendu : **tous les tests PASS** (environ 12 tests)

- [ ] **Step 3 : Vérifier pas de régression sur la suite complète**

```
pytest -q
```

Résultat attendu : tous les tests passent (334+ tests)

- [ ] **Step 4 : Commiter**

```bash
git add src/generators/cockpit_ingenieur_generator.py tests/test_cockpit_ingenieur.py
git commit -m "feat: cockpit ingénieur agenda — Mes UOs + Agenda + _Manifeste avec COMMENTAIRE"
```

---

## Task 5 : Tests dashboard métier (TDD — écrire d'abord)

**Files:**
- Create: `tests/test_dashboard_metier.py`

- [ ] **Step 1 : Écrire le fichier de tests complet**

```python
# tests/test_dashboard_metier.py
"""Tests TDD pour dashboard_metier_generator."""
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

from src.models import (
    UOInstance, UOType, Activity, System, Project, StatutUO,
    ProfilActeur, Role, TypeFiltre, NiveauAcces,
)
from src.store import JsonStore


def _make_uo(uid: str, engineer: str, hours: float, end: date) -> UOInstance:
    uo_type = UOType(id="spec_fonctionnelle", name="Spec Fonctionnelle", activities=[
        Activity(id="A1", name="Analyse", default_hours=hours * 0.5),
    ])
    return UOInstance(
        id=uid, uo_type_id="spec_fonctionnelle", system_id="clim", project_id="MI20",
        engineer_name=engineer, total_hours=hours,
        start_date=date(2026, 4, 1), end_date=end,
        statut=StatutUO.EN_COURS,
        uo_type=uo_type,
        system=System(id="clim", name="Climatisation"),
        project=Project(id="MI20", name="MI20 RATP"),
    )


def _make_pilote_metier() -> ProfilActeur:
    return ProfilActeur(
        id="USR004", nom="Jean-Luc Bernard",
        role=Role.PILOTE_METIER,
        filtre_type=TypeFiltre.INGENIEUR,
        filtre_valeur="Alice Dubois,Bruno Lecomte",
        acces=NiveauAcces.READ,
    )


ALL_UOS = [
    _make_uo("UO-001", "Alice Dubois",  32, date(2026, 6, 30)),
    _make_uo("UO-002", "Alice Dubois",  48, date(2026, 7, 15)),
    _make_uo("UO-003", "Bruno Lecomte", 40, date(2026, 6, 20)),
    _make_uo("UO-004", "Denis Renard",  24, date(2026, 8, 1)),   # hors périmètre
]


class TestDashboardFichier:
    def test_fichier_cree(self, tmp_path):
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        assert path.exists()
        assert path.name == "Dashboard_Jean-Luc_Bernard.xlsx"

    def test_quatre_onglets_presents(self, tmp_path):
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        wb = load_workbook(path)
        assert "Synthèse" in wb.sheetnames
        assert "Par Ingénieur" in wb.sheetnames
        assert "Alertes" in wb.sheetnames
        assert "_Manifeste" in wb.sheetnames


class TestDashboardFiltrage:
    def test_filtre_ingenieur_respecte(self, tmp_path):
        """Denis Renard ne doit pas apparaître dans le dashboard de Jean-Luc."""
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["Synthèse"]
        all_values = [ws.cell(row=r, column=c).value
                      for r in range(1, 50) for c in range(1, 12)]
        assert "UO-004" not in all_values   # UO de Denis
        assert "Denis Renard" not in all_values

    def test_uo_alice_et_bruno_presents(self, tmp_path):
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["Synthèse"]
        all_values = [ws.cell(row=r, column=c).value
                      for r in range(1, 50) for c in range(1, 12)]
        assert "UO-001" in all_values
        assert "UO-002" in all_values
        assert "UO-003" in all_values


class TestDashboardKPIs:
    def test_charge_totale_correcte(self, tmp_path):
        """32 + 48 + 40 = 120h pour Alice + Bruno (Denis exclu)."""
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["Synthèse"]
        all_values = [ws.cell(row=r, column=c).value
                      for r in range(1, 10) for c in range(1, 15)]
        assert 120 in all_values or "120h" in [str(v) for v in all_values if v]

    def test_nb_uo_kpi_correct(self, tmp_path):
        """3 UOs dans le périmètre (Alice×2 + Bruno×1)."""
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["Synthèse"]
        all_values = [ws.cell(row=r, column=c).value
                      for r in range(1, 10) for c in range(1, 15)]
        assert 3 in all_values or "3 UOs" in [str(v) for v in all_values if v]


class TestDashboardAlertes:
    def test_alerte_depassement_heures(self, tmp_path):
        """UO avec heures_realisees > charge doit apparaître dans Alertes."""
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        store.set("uo.UO-001.heures_realisees", 50.0)  # > 32h allouées
        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["Alertes"]
        all_values = [ws.cell(row=r, column=c).value
                      for r in range(1, 30) for c in range(1, 8)]
        assert "UO-001" in all_values

    def test_pas_alerte_si_heures_ok(self, tmp_path):
        """UO sans dépassement ne doit pas apparaître comme alerte dérive."""
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        store.set("uo.UO-001.heures_realisees", 10.0)  # < 32h — OK
        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        wb = load_workbook(path, data_only=True)
        ws = wb["Alertes"]
        # Cherche "Dépassement H" associé à UO-001
        found = False
        for r in range(3, 30):
            uid = ws.cell(row=r, column=2).value
            type_alerte = ws.cell(row=r, column=3).value
            if uid == "UO-001" and type_alerte and "Dépassement" in str(type_alerte):
                found = True
        assert not found


class TestDashboardManifeste:
    def test_version_manifeste(self, tmp_path):
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        assert str(ws["A1"].value).startswith("MANIFESTE_V=")

    def test_colonne_commentaire_presente(self, tmp_path):
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        headers = [ws.cell(row=2, column=c).value for c in range(1, 15)]
        assert "COMMENTAIRE" in headers

    def test_regles_pull_avancement_presentes(self, tmp_path):
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        wb = load_workbook(path)
        ws = wb["_Manifeste"]
        nom_globals = [ws.cell(row=r, column=3).value for r in range(3, 30)
                       if ws.cell(row=r, column=3).value]
        assert "uo.UO-001.avancement" in nom_globals
        assert "uo.UO-003.avancement" in nom_globals
        # Denis Renard hors périmètre — sa clé ne doit pas être là
        assert "uo.UO-004.avancement" not in nom_globals
```

- [ ] **Step 2 : Vérifier que les tests échouent (module absent)**

```
pytest tests/test_dashboard_metier.py -q
```

Résultat attendu : `ModuleNotFoundError: No module named 'src.generators.dashboard_metier_generator'`

---

## Task 6 : Générateur dashboard métier — `Synthèse` + `Par Ingénieur`

**Files:**
- Create: `src/generators/dashboard_metier_generator.py`

- [ ] **Step 1 : Créer le fichier avec filtrage, `Synthèse` et `Par Ingénieur`**

```python
# src/generators/dashboard_metier_generator.py
"""Dashboard métier — Vue consolidée équipe : Synthèse + Par Ingénieur + Alertes + _Manifeste."""
from datetime import date, timedelta
from pathlib import Path
from typing import List, Tuple

from openpyxl import Workbook

from src.models import ProfilActeur, TypeFiltre, UOInstance
from src.store import JsonStore
from src.styles import (
    BLUE_DARK, BLUE_MID, BLUE_LIGHT, GREEN_LIGHT, ORANGE_LIGHT, RED_LIGHT,
    YELLOW_LIGHT, WHITE, GREY_LIGHT, THIN_BORDER,
    solid_fill, header_font, body_font, center, left,
    style_header_row, style_data_row, set_column_widths, freeze_top_row,
)

OUTPUT_DIR = Path(__file__).parent.parent.parent / "output" / "cockpits"
COCKPIT_DIR = Path(__file__).parent.parent.parent / "output" / "cockpits"


def _filter_instances(acteur: ProfilActeur, instances: List[UOInstance]) -> List[UOInstance]:
    """Filtre les UOs selon le filtre_type et filtre_valeur de l'acteur."""
    if acteur.filtre_valeur == "ALL":
        return list(instances)
    if acteur.filtre_type == TypeFiltre.INGENIEUR:
        noms = {n.strip() for n in acteur.filtre_valeur.split(",")}
        return [u for u in instances if u.engineer_name in noms]
    if acteur.filtre_type == TypeFiltre.PROJET:
        projets = {p.strip() for p in acteur.filtre_valeur.split(",")}
        return [u for u in instances if u.project_id in projets]
    return list(instances)


def generate_dashboard_metier(
    acteur: ProfilActeur,
    all_instances: List[UOInstance],
    store: JsonStore,
    output_dir: Path = OUTPUT_DIR,
) -> Path:
    """Génère le dashboard pour un pilote métier.

    Args:
        acteur: ProfilActeur du pilote métier (filtre_valeur utilisé pour restreindre les UOs)
        all_instances: Toutes les UOInstances disponibles
        store: JsonStore pour lire avancement/heures réalisées
        output_dir: Répertoire de sortie

    Returns:
        Chemin du fichier Excel généré
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    uo_list = _filter_instances(acteur, all_instances)

    wb = Workbook()
    wb.remove(wb.active)

    _sheet_synthese(wb, acteur, uo_list, store)
    _sheet_par_ingenieur(wb, acteur, uo_list, store)
    _sheet_alertes(wb, uo_list, store)
    _sheet_manifeste_dashboard(wb, uo_list)

    safe = acteur.nom.replace(" ", "_")
    filepath = output_dir / f"Dashboard_{safe}.xlsx"
    wb.save(filepath)
    return filepath


def _get_store_float(store: JsonStore, key: str, default: float = 0.0) -> float:
    """Lit une valeur float du store, retourne default si absente."""
    val = store.get(key)
    if val is None:
        return default
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


def _sheet_synthese(wb: Workbook, acteur: ProfilActeur,
                    uo_list: List[UOInstance], store: JsonStore):
    ws = wb.create_sheet("Synthèse")
    ws.sheet_view.showGridLines = False

    # ── Titre ──────────────────────────────────────────────────────────────────
    ws.merge_cells("A1:J1")
    t = ws["A1"]
    t.value = (f"Dashboard Métier — {acteur.nom}   |   "
               f"{date.today().strftime('%d/%m/%Y')}")
    t.fill = solid_fill(BLUE_DARK)
    t.font = header_font(size=14)
    t.alignment = center()
    ws.row_dimensions[1].height = 32

    # ── KPIs ───────────────────────────────────────────────────────────────────
    total_h = sum(u.total_hours for u in uo_list)
    avancement_vals = [_get_store_float(store, f"uo.{u.id}.avancement") for u in uo_list]
    avg_avanc = sum(avancement_vals) / len(avancement_vals) if avancement_vals else 0

    kpis = [
        ("A2:B2", f"{len(uo_list)} UOs", "Périmètre équipe"),
        ("C2:D2", f"{total_h}h", "Charge totale"),
        ("E2:F2", f"{avg_avanc:.0%}", "Avancement moyen"),
    ]
    for cell_range, value, label in kpis:
        ws.merge_cells(cell_range)
        first_cell = cell_range.split(":")[0]
        c = ws[first_cell]
        c.value = f"{value}  ({label})"
        c.fill = solid_fill(BLUE_LIGHT)
        c.font = body_font(bold=True, color="1F3864")
        c.alignment = center()
        c.border = THIN_BORDER
    ws.row_dimensions[2].height = 24

    # ── Tableau consolidé ──────────────────────────────────────────────────────
    ws.merge_cells("A4:J4")
    sec = ws["A4"]
    sec.value = "Toutes les UOs de l'équipe"
    sec.fill = solid_fill(BLUE_MID)
    sec.font = header_font()
    sec.alignment = center()

    headers = ["UO ID", "Ingénieur", "Type UO", "Système", "Projet",
               "Charge (h)", "% Avancement", "H réalisées", "Date fin", "Alerte"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=5, column=col, value=h)
    style_header_row(ws, 5, 1, 10, color=BLUE_MID)

    for i, uo in enumerate(uo_list):
        row = 6 + i
        avancement = _get_store_float(store, f"uo.{uo.id}.avancement")
        h_real = _get_store_float(store, f"uo.{uo.id}.heures_realisees")
        type_name = uo.uo_type.name if uo.uo_type else uo.uo_type_id
        sys_name = uo.system.name if uo.system else uo.system_id
        proj_name = uo.project.name if uo.project else uo.project_id

        ws.cell(row=row, column=1, value=uo.id)
        ws.cell(row=row, column=2, value=uo.engineer_name)
        ws.cell(row=row, column=3, value=type_name)
        ws.cell(row=row, column=4, value=sys_name)
        ws.cell(row=row, column=5, value=proj_name)
        ws.cell(row=row, column=6, value=uo.total_hours)

        avanc_cell = ws.cell(row=row, column=7, value=avancement)
        avanc_cell.number_format = "0%"

        ws.cell(row=row, column=8, value=h_real)

        date_cell = ws.cell(row=row, column=9, value=uo.end_date)
        date_cell.number_format = "DD/MM/YYYY"

        # Alerte calculée (Python, pas formule, car valeurs viennent du store)
        today = date.today()
        if h_real > uo.total_hours:
            alerte = "⚠ Dérive heures"
        elif uo.end_date and uo.end_date <= today + timedelta(days=7):
            alerte = "⏰ Échéance proche"
        else:
            alerte = "✅ OK"
        ws.cell(row=row, column=10, value=alerte)

        style_data_row(ws, row, 1, 10, alternate=(i % 2 == 1))

    set_column_widths(ws, {
        "A": 12, "B": 22, "C": 28, "D": 18, "E": 22,
        "F": 13, "G": 16, "H": 14, "I": 14, "J": 22,
    })
    ws.freeze_panes = "A6"


def _sheet_par_ingenieur(wb: Workbook, acteur: ProfilActeur,
                         uo_list: List[UOInstance], store: JsonStore):
    ws = wb.create_sheet("Par Ingénieur")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:H1")
    t = ws["A1"]
    t.value = f"Détail par Ingénieur — {acteur.nom}"
    t.fill = solid_fill(BLUE_DARK)
    t.font = header_font(size=13)
    t.alignment = center()
    ws.row_dimensions[1].height = 28

    engineers = sorted(set(u.engineer_name for u in uo_list))
    current_row = 3

    for eng in engineers:
        eng_uo = [u for u in uo_list if u.engineer_name == eng]
        total_h = sum(u.total_hours for u in eng_uo)
        avancs = [_get_store_float(store, f"uo.{u.id}.avancement") for u in eng_uo]
        avg_avanc = sum(avancs) / len(avancs) if avancs else 0.0

        # En-tête ingénieur
        ws.merge_cells(f"A{current_row}:H{current_row}")
        eng_cell = ws[f"A{current_row}"]
        eng_cell.value = (f"▶  {eng}   |   {len(eng_uo)} UO   |   "
                          f"{total_h}h   |   {avg_avanc:.0%} avancement moyen")
        eng_cell.fill = solid_fill(BLUE_MID)
        eng_cell.font = header_font()
        eng_cell.alignment = left()
        ws.row_dimensions[current_row].height = 22
        current_row += 1

        # Sous-en-têtes
        sub_headers = ["UO ID", "Type", "Système", "Projet",
                       "Charge (h)", "% Avancement", "H réalisées", "Date fin"]
        for col, h in enumerate(sub_headers, 1):
            ws.cell(row=current_row, column=col, value=h)
        style_header_row(ws, current_row, 1, 8, color=BLUE_LIGHT)
        for col in range(1, 9):
            ws.cell(row=current_row, column=col).font = body_font(bold=True, color="1F3864")
        current_row += 1

        for i, uo in enumerate(eng_uo):
            avancement = _get_store_float(store, f"uo.{uo.id}.avancement")
            h_real = _get_store_float(store, f"uo.{uo.id}.heures_realisees")
            type_name = uo.uo_type.name if uo.uo_type else uo.uo_type_id
            sys_name = uo.system.name if uo.system else uo.system_id
            proj_name = uo.project.name if uo.project else uo.project_id

            ws.cell(row=current_row, column=1, value=uo.id)
            ws.cell(row=current_row, column=2, value=type_name)
            ws.cell(row=current_row, column=3, value=sys_name)
            ws.cell(row=current_row, column=4, value=proj_name)
            ws.cell(row=current_row, column=5, value=uo.total_hours)
            avc = ws.cell(row=current_row, column=6, value=avancement)
            avc.number_format = "0%"
            ws.cell(row=current_row, column=7, value=h_real)
            dc = ws.cell(row=current_row, column=8, value=uo.end_date)
            dc.number_format = "DD/MM/YYYY"
            style_data_row(ws, current_row, 1, 8, alternate=(i % 2 == 1))
            current_row += 1

        # Lien vers cockpit
        cockpit_path = COCKPIT_DIR / f"Cockpit_{eng.replace(' ', '_')}.xlsx"
        ws.merge_cells(f"A{current_row}:H{current_row}")
        link_cell = ws[f"A{current_row}"]
        link_cell.value = f"→ Ouvrir cockpit {eng}"
        link_cell.hyperlink = str(cockpit_path)
        link_cell.font = body_font(color="0563C1")
        link_cell.alignment = left()
        link_cell.fill = solid_fill(GREY_LIGHT)
        link_cell.border = THIN_BORDER
        current_row += 2

    set_column_widths(ws, {
        "A": 12, "B": 28, "C": 18, "D": 22, "E": 13, "F": 16, "G": 14, "H": 14,
    })
```

- [ ] **Step 2 : Lancer les tests filtrage et KPIs**

```
pytest tests/test_dashboard_metier.py::TestDashboardFichier tests/test_dashboard_metier.py::TestDashboardFiltrage tests/test_dashboard_metier.py::TestDashboardKPIs -v
```

Résultat attendu : tests `TestDashboardFichier` et `TestDashboardFiltrage` PASS, `TestDashboardKPIs` potentiellement PASS

---

## Task 7 : Dashboard métier — onglet `Alertes`

**Files:**
- Modify: `src/generators/dashboard_metier_generator.py`

- [ ] **Step 1 : Ajouter la fonction `_sheet_alertes`**

Ajouter après `_sheet_par_ingenieur` dans le fichier :

```python
def _compute_alerts(uo_list: List[UOInstance], store: JsonStore) -> List[Tuple]:
    """Calcule les alertes à partir des données du store. Retourne liste triée par criticité."""
    today = date.today()
    alerts = []

    for uo in uo_list:
        h_real = _get_store_float(store, f"uo.{uo.id}.heures_realisees")

        if h_real > uo.total_hours:
            alerts.append((
                uo.engineer_name, uo.id,
                "Dépassement H",
                f"{h_real:.0f}h réalisées / {uo.total_hours}h allouées",
                "🔴 Critique", 0,
            ))

        if uo.end_date and uo.end_date <= today + timedelta(days=7):
            jours = (uo.end_date - today).days
            label = f"Échéance dans {jours}j" if jours >= 0 else f"Dépassée de {-jours}j"
            criticite = "🔴 Critique" if jours < 0 else "🟠 Élevée"
            prio = 0 if jours < 0 else 1
            alerts.append((
                uo.engineer_name, uo.id,
                "Échéance critique",
                label,
                criticite, prio,
            ))

    # Tri par criticité (0=Critique, 1=Élevée, 2=Normale)
    alerts.sort(key=lambda x: x[5])
    return alerts


def _sheet_alertes(wb: Workbook, uo_list: List[UOInstance], store: JsonStore):
    ws = wb.create_sheet("Alertes")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:E1")
    t = ws["A1"]
    t.value = f"Alertes & Risques — {date.today().strftime('%d/%m/%Y')}"
    t.fill = solid_fill("C00000")
    t.font = header_font(size=13)
    t.alignment = center()
    ws.row_dimensions[1].height = 28

    headers = ["Ingénieur", "UO ID", "Type alerte", "Détail", "Criticité"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=2, column=col, value=h)
    style_header_row(ws, 2, 1, 5, color=BLUE_MID)

    alerts = _compute_alerts(uo_list, store)

    if not alerts:
        ws.merge_cells("A3:E3")
        c = ws["A3"]
        c.value = "✅ Aucune alerte active"
        c.fill = solid_fill(GREEN_LIGHT)
        c.font = body_font(bold=True, color="1F3864")
        c.alignment = center()
        c.border = THIN_BORDER
    else:
        for i, (eng, uid, type_a, detail, criticite, _) in enumerate(alerts):
            row = 3 + i
            ws.cell(row=row, column=1, value=eng)
            ws.cell(row=row, column=2, value=uid)
            ws.cell(row=row, column=3, value=type_a)
            ws.cell(row=row, column=4, value=detail)
            ws.cell(row=row, column=5, value=criticite)

            fill_color = RED_LIGHT if "Critique" in criticite else ORANGE_LIGHT
            for col in range(1, 6):
                c = ws.cell(row=row, column=col)
                c.fill = solid_fill(fill_color)
                c.border = THIN_BORDER
                c.alignment = left()
                c.font = body_font()

    set_column_widths(ws, {"A": 22, "B": 12, "C": 20, "D": 40, "E": 15})
    ws.freeze_panes = "A3"
```

- [ ] **Step 2 : Lancer les tests alertes**

```
pytest tests/test_dashboard_metier.py::TestDashboardAlertes -v
```

Résultat attendu : 2 tests PASS

---

## Task 8 : Dashboard métier — onglet `_Manifeste`

**Files:**
- Modify: `src/generators/dashboard_metier_generator.py`

- [ ] **Step 1 : Ajouter la constante et la fonction `_sheet_manifeste_dashboard`**

Ajouter en bas du fichier (définir la constante localement — pas d'import entre générateurs) :

```python
# Même constante que dans cockpit_ingenieur_generator — définie ici pour éviter le couplage
_MANIFESTE_HEADERS = [
    "TYPE", "SCOPE", "NOM_GLOBAL", "NOM_LOCAL",
    "FEUILLE", "TABLEAU", "CLE", "COLONNES",
    "CELLULE", "DIRECTION", "FORMULE", "COMMENTAIRE",
]


def _sheet_manifeste_dashboard(wb: Workbook, uo_list: List[UOInstance]):
    """_Manifeste du dashboard : règles PULL pour chaque UO du périmètre."""
    ws = wb.create_sheet("_Manifeste")
    ws.sheet_view.showGridLines = False

    ws["A1"] = "MANIFESTE_V=1"
    ws["A1"].font = body_font(bold=True, color="1F3864")

    for col, h in enumerate(_MANIFESTE_HEADERS, 1):
        ws.cell(row=2, column=col, value=h)
    style_header_row(ws, 2, 1, len(_MANIFESTE_HEADERS), color=BLUE_MID)

    row = 3
    for i, uo in enumerate(uo_list):
        # Ligne de données Synthèse : avancement en colonne G, heures en H
        synth_row = 6 + i

        rules = [
            ("CELL_PCT",  "GLOBAL", f"uo.{uo.id}.avancement",       "", "Synthèse",      "", "", "", f"G{synth_row}", "PULL", "",
             f"Récupère l'avancement de {uo.id} poussé par le cockpit de {uo.engineer_name}"),
            ("CELL_NUM",  "GLOBAL", f"uo.{uo.id}.heures_realisees",  "", "Synthèse",      "", "", "", f"H{synth_row}", "PULL", "",
             f"Récupère les heures réalisées de {uo.id} poussées par le cockpit de {uo.engineer_name}"),
            ("CELL_NUM",  "GLOBAL", f"uo.{uo.id}.points_ouverts",   "", "Alertes",       "", "", "", "",              "PULL", "",
             f"Récupère le nb de points ouverts de {uo.id} pour alimentation des alertes"),
        ]

        for rule in rules:
            for col, val in enumerate(rule, 1):
                c = ws.cell(row=row, column=col, value=val)
                c.font = body_font(size=10)
                c.border = THIN_BORDER
                c.alignment = left()
            ws.cell(row=row, column=12).font = body_font(size=10, color="666666")
            ws.cell(row=row, column=12).fill = solid_fill("F7F7F7")
            row += 1

        row += 1  # ligne vide entre UOs

    set_column_widths(ws, {
        "A": 12, "B": 8, "C": 35, "D": 14, "E": 12, "F": 10,
        "G": 8, "H": 10, "I": 10, "J": 10, "K": 10, "L": 60,
    })
```

- [ ] **Step 2 : Lancer tous les tests dashboard**

```
pytest tests/test_dashboard_metier.py -v
```

Résultat attendu : **tous les tests PASS**

- [ ] **Step 3 : Lancer la suite complète — vérifier pas de régression**

```
pytest -q
```

Résultat attendu : tous les tests PASS (334+ tests)

- [ ] **Step 4 : Commiter**

```bash
git add src/generators/dashboard_metier_generator.py tests/test_dashboard_metier.py
git commit -m "feat: dashboard métier — Synthèse + Par Ingénieur + Alertes + _Manifeste avec COMMENTAIRE"
```

---

## Task 9 : Test end-to-end du cycle push/pull

**Files:**
- Modify: `tests/test_dashboard_metier.py`

- [ ] **Step 1 : Ajouter la classe de test end-to-end à la fin de `test_dashboard_metier.py`**

```python
class TestPushPullCycle:
    """Vérifie le cycle complet : store → dashboard (simulation du push ingénieur)."""

    def test_avancement_store_visible_dans_synthese(self, tmp_path):
        """Simule un push de 80% d'avancement → vérifie que Synthèse affiche 0.8."""
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")

        # Simule ce que ferait ExoSync après le push du cockpit d'Alice
        store.set_many({
            "uo.UO-001.avancement": 0.8,
            "uo.UO-001.heures_realisees": 25.0,
            "uo.UO-002.avancement": 0.5,
            "uo.UO-002.heures_realisees": 24.0,
            "uo.UO-003.avancement": 0.3,
            "uo.UO-003.heures_realisees": 12.0,
        })

        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        wb = load_workbook(path, data_only=True)
        ws = wb["Synthèse"]

        # Cherche UO-001 dans la col A, lit avancement en col G (même ligne)
        for row in range(6, 20):
            if ws.cell(row=row, column=1).value == "UO-001":
                avanc = ws.cell(row=row, column=7).value
                assert avanc == pytest.approx(0.8, abs=0.01), \
                    f"Attendu 0.8, obtenu {avanc}"
                break
        else:
            pytest.fail("UO-001 non trouvé dans Synthèse")

    def test_alerte_generee_apres_depassement(self, tmp_path):
        """Simule un dépassement : heures_realisees > charge → alerte dans Alertes."""
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")

        # UO-003 : 40h allouées, Bruno en dépassement à 55h
        store.set("uo.UO-003.heures_realisees", 55.0)

        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        wb = load_workbook(path, data_only=True)
        ws = wb["Alertes"]

        found = any(
            ws.cell(row=r, column=2).value == "UO-003"
            for r in range(3, 20)
        )
        assert found, "UO-003 devrait apparaître dans les Alertes après dépassement"

    def test_store_vide_ne_plante_pas(self, tmp_path):
        """Store vide → dashboard généré avec des 0 (pas d'exception)."""
        from src.generators.dashboard_metier_generator import generate_dashboard_metier
        acteur = _make_pilote_metier()
        store = JsonStore(tmp_path / "store.json")
        # store vide — pas de set_many

        path = generate_dashboard_metier(acteur, ALL_UOS, store, output_dir=tmp_path)
        assert path.exists()
        wb = load_workbook(path, data_only=True)
        ws = wb["Synthèse"]
        # La cellule avancement de UO-001 doit être 0.0 (valeur par défaut)
        for row in range(6, 20):
            if ws.cell(row=row, column=1).value == "UO-001":
                avanc = ws.cell(row=row, column=7).value
                assert avanc == pytest.approx(0.0, abs=0.01)
                break
```

- [ ] **Step 2 : Lancer les tests end-to-end**

```
pytest tests/test_dashboard_metier.py::TestPushPullCycle -v
```

Résultat attendu : 3 tests PASS

- [ ] **Step 3 : Suite complète finale**

```
pytest -q
```

Résultat attendu : tous les tests PASS

- [ ] **Step 4 : Commiter**

```bash
git add tests/test_dashboard_metier.py
git commit -m "test: cycle push/pull end-to-end store → dashboard métier"
```

- [ ] **Step 5 : Push**

```bash
git push origin master
```

---

## Résumé des fichiers créés

| Fichier | Lignes estimées |
|---------|----------------|
| `src/generators/cockpit_ingenieur_generator.py` | ~200 |
| `src/generators/dashboard_metier_generator.py` | ~250 |
| `tests/test_cockpit_ingenieur.py` | ~100 |
| `tests/test_dashboard_metier.py` | ~170 |

**Nouveaux tests ajoutés :** ~35 tests  
**Fichiers existants modifiés :** aucun  
**Régressions possibles :** aucune (ajouts purs)
