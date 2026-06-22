# Assembleur d'instanciation UO (C4) — Plan d'implémentation

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Un script `assembler.py` transforme une seule commande CLI en UO créée + cockpit mis à jour + sync, sans jamais écraser les saisies ingénieur (colonnes jaunes 5-6).

**Architecture:** `assembler.py` orchestre trois actions indépendantes — appel de `creer_uo.py` (inchangé), patch chirurgical du cockpit existant via `ajouter_uo_au_cockpit()`, création automatique si absent via `creer_cockpit_vide()`. Chaque fonction est testée isolément avant intégration.

**Tech Stack:** Python 3.12, openpyxl 3.1, pytest. Tous les fichiers sont dans `projet_TrainSystem/` sauf les tests (`tests/`).

---

## Fichiers concernés

| Action  | Fichier |
|---------|---------|
| Créer   | `projet_TrainSystem/assembler.py` |
| Créer   | `tests/test_assembler.py` |
| Modifier | `projet_TrainSystem/creer_cockpit_se.py` (onglet Agenda + table toujours créée + flag `--safe`) |

---

## Contexte technique à lire avant de commencer

### Structure du cockpit (`Cockpit_Alice_Dubois.xlsx` — feuille "Mes UOs")

```
Ligne 1 : Bannière (fusion A1:J1)
Ligne 2 : vide
Ligne 3 : En-têtes de tbl_mes_uos — cols A..H :
          [UO ID | Système | Projet | Charge (h) | % Avancement | H réalisées | Date fin | Alerte]
Ligne 4+ : données (une ligne par UO)
```

- `tbl_mes_uos` ref = `A3:H{last_row}` (openpyxl : `ws.tables["tbl_mes_uos"].ref`)
- Colonnes **5** (% Avancement) et **6** (H réalisées) = zones jaunes (fill `FFF2CC`) — saisie ingénieur, **ne jamais écraser**
- Lorsque la liste UO est vide, `creer_cockpit_se.py` n'enregistre PAS la table (`if uo_list: ...`) → bug à corriger en Tâche 1

### Import de `creer_uo.py` depuis `assembler.py`

`creer_uo.py` expose `build_instance(code, uo_type, projet_code, systeme_code, args)` et `main()`. Pour l'appeler depuis l'assembleur sans fork de processus :

```python
# projet_TrainSystem/assembler.py
import sys, types
from pathlib import Path
HERE = Path(__file__).parent
sys.path.insert(0, str(HERE.parent))
sys.path.insert(0, str(HERE))
```

Puis `import creer_uo` et appel de `creer_uo.build_instance(...)` ou subprocess.

### Sync

```python
from src.sync import synchroniser_repertoire
synchroniser_repertoire(HERE)   # HERE = projet_TrainSystem/
```

---

## Tâche 0 — Ajouter l'onglet Agenda dans `creer_cockpit_se.py`

**Fichiers :**
- Modifier : `projet_TrainSystem/creer_cockpit_se.py` (ajouter `_sheet_agenda()` + appel dans `generer_cockpit()`)

### Contexte

L'ancienne version du cockpit (commit `6cf48a9`, `src/generators/cockpit_ingenieur_generator.py`) avait un onglet "Agenda" avec 3 sections : *Semaine en cours*, *Prochaines échéances (30 jours)*, *Points ouverts / Actions*. Le générateur actuel (`creer_cockpit_se.py`) ne le produit pas.

Adaptation : le nouveau Agenda ne peut pas lire dynamiquement les dates des activités depuis les UO (cela demanderait d'ouvrir chaque fichier). Il génère à la place une ligne de saisie **jaune** par UO dans chaque section — l'ingénieur remplit lui-même ses échéances de la semaine. Structure identique à l'original, données vides.

- [ ] **Étape 0.1 : Écrire le test qui échoue**

  Ajouter dans `tests/test_assembler.py` :

  ```python
  def test_cockpit_has_agenda_sheet(tmp_path):
      """Le cockpit doit avoir un onglet 'Agenda' avec les 3 sections."""
      from creer_cockpit_se import generer_cockpit

      out = generer_cockpit("Alice Dubois", [
          {"file_id": "L09U1-CFL2400-CLIM", "systeme": "Clim",
           "projet": "CFL", "heures": 200}
      ], "USR004", tmp_path)

      wb = load_workbook(str(out))
      assert "Agenda" in wb.sheetnames

      ws = wb["Agenda"]
      # Vérifier que les 3 titres de section sont présents dans la feuille
      all_values = [ws.cell(row=r, column=c).value
                    for r in range(1, ws.max_row + 1)
                    for c in range(1, 3)]
      all_str = [str(v) for v in all_values if v]
      assert any("Semaine" in s for s in all_str)
      assert any("30" in s for s in all_str)
      assert any("ouverts" in s.lower() or "action" in s.lower() for s in all_str)
  ```

- [ ] **Étape 0.2 : Vérifier que le test échoue**

  ```
  python -m pytest tests/test_assembler.py::test_cockpit_has_agenda_sheet -v
  ```

  Attendu : `FAILED — AssertionError: 'Agenda' not in ['Mes UOs', '_Manifeste', '_Log']`.

- [ ] **Étape 0.3 : Ajouter `_sheet_agenda()` dans `creer_cockpit_se.py`**

  Ajouter les imports manquants en tête du fichier (après les imports existants) :

  ```python
  from datetime import date, timedelta
  ```

  Ajouter la fonction `_sheet_agenda` avant `generer_cockpit` :

  ```python
  def _sheet_agenda(wb: Workbook, se_name: str, uo_list: list[dict]):
      """Onglet Agenda — vue semaine + 30j + points ouverts. Lignes jaunes à remplir par le SE."""
      ws = wb.create_sheet("Agenda")
      ws.sheet_view.showGridLines = False
      today = date.today()

      # Bannière
      ws.merge_cells("A1:F1")
      c = ws["A1"]
      c.value = f"Agenda — {se_name}   |   Semaine du {today.strftime('%d/%m/%Y')}"
      c.fill = _fill("0C447C")
      c.font = _fnt(13, bold=True, color="FFFFFF")
      c.alignment = _center()
      ws.row_dimensions[1].height = 30

      row = 3
      row = _agenda_section(ws, "📅  Semaine en cours", row, uo_list, bg="185FA5")
      row += 1
      row = _agenda_section(ws, "📋  Prochaines échéances (30 jours)", row, uo_list, bg="2E75B6")
      row += 1
      row = _agenda_po(ws, row, uo_list)

      for col, w in zip("ABCDEF", [14, 38, 12, 16, 18, 30]):
          ws.column_dimensions[col].width = w


  def _agenda_section(ws, title: str, start_row: int,
                      uo_list: list[dict], bg: str = "185FA5") -> int:
      """Section agenda avec 1 ligne jaune saisie par UO. Retourne la prochaine ligne libre."""
      ws.merge_cells(f"A{start_row}:F{start_row}")
      c = ws[f"A{start_row}"]
      c.value = title
      c.fill = _fill(bg)
      c.font = _fnt(11, bold=True, color="FFFFFF")
      c.alignment = _left()
      ws.row_dimensions[start_row].height = 22
      start_row += 1

      headers = ["UO ID", "Activité", "Priorité", "Date échéance", "Statut", "Action"]
      for col, h in enumerate(headers, 1):
          c = ws.cell(row=start_row, column=col, value=h)
          c.fill = _fill("D6E4F0")
          c.font = _fnt(9.5, bold=True, color="1F3864")
          c.alignment = _center()
          c.border = _thin_border()
      ws.row_dimensions[start_row].height = 20
      start_row += 1

      for uo in uo_list:
          ws.cell(row=start_row, column=1, value=uo["file_id"]).font = _fnt(9.5, color="0563C1")
          ws.cell(row=start_row, column=1).alignment = _left()
          ws.cell(row=start_row, column=1).border = _thin_border()
          for col in range(2, 7):
              c = ws.cell(row=start_row, column=col, value="")
              c.fill = _fill("FFF2CC")   # jaune — saisie ingénieur
              c.border = _thin_border()
              c.alignment = _center()
          ws.row_dimensions[start_row].height = 18
          start_row += 1

      if not uo_list:
          ws.merge_cells(f"A{start_row}:F{start_row}")
          ws[f"A{start_row}"].value = "Aucune UO"
          ws[f"A{start_row}"].font = _fnt(9, color="999999")
          ws[f"A{start_row}"].fill = _fill("F9F9F9")
          ws[f"A{start_row}"].alignment = _center()
          start_row += 1

      return start_row


  def _agenda_po(ws, start_row: int, uo_list: list[dict]) -> int:
      """Section Points ouverts / Actions. Retourne la prochaine ligne libre."""
      ws.merge_cells(f"A{start_row}:F{start_row}")
      c = ws[f"A{start_row}"]
      c.value = "⚡  Points ouverts / Actions"
      c.fill = _fill("FAEEDA")
      c.font = _fnt(11, bold=True, color="854F0B")
      c.alignment = _left()
      ws.row_dimensions[start_row].height = 22
      start_row += 1

      headers = ["UO ID", "Description action", "Responsable", "Date limite",
                 "Nb points", "Statut"]
      for col, h in enumerate(headers, 1):
          c = ws.cell(row=start_row, column=col, value=h)
          c.fill = _fill("D6E4F0")
          c.font = _fnt(9.5, bold=True, color="1F3864")
          c.alignment = _center()
          c.border = _thin_border()
      ws.row_dimensions[start_row].height = 20
      start_row += 1

      for uo in uo_list:
          ws.cell(row=start_row, column=1, value=uo["file_id"]).font = _fnt(9.5, color="0563C1")
          ws.cell(row=start_row, column=1).alignment = _left()
          ws.cell(row=start_row, column=1).border = _thin_border()
          for col in range(2, 7):
              c = ws.cell(row=start_row, column=col, value="")
              c.fill = _fill("FFF2CC")
              c.border = _thin_border()
              c.alignment = _center()
          ws.row_dimensions[start_row].height = 18
          start_row += 1

      return start_row
  ```

- [ ] **Étape 0.4 : Appeler `_sheet_agenda` dans `generer_cockpit()`**

  Dans `generer_cockpit()`, après l'appel à `_sheet_mes_uos(wb, se_name, uo_list)` et avant `_sheet_manifeste(...)`, ajouter :

  ```python
  _sheet_agenda(wb, se_name, uo_list)
  ```

- [ ] **Étape 0.5 : Vérifier que le test passe**

  ```
  python -m pytest tests/test_assembler.py::test_cockpit_has_agenda_sheet -v
  ```

  Attendu : `PASSED`.

- [ ] **Étape 0.6 : Régression globale**

  ```
  python -m pytest tests/ -q
  ```

  Attendu : `382+ passed, 0 failed`.

- [ ] **Étape 0.7 : Commit**

  ```bash
  git add projet_TrainSystem/creer_cockpit_se.py tests/test_assembler.py
  git commit -m "feat: onglet Agenda dans cockpit SE — 3 sections, lignes jaunes par UO"
  ```

---

## Tâche 1 — Corriger `creer_cockpit_se.py` : table toujours créée + flag `--safe`

**Fichiers :**
- Modifier : `projet_TrainSystem/creer_cockpit_se.py:163-174`

### Pourquoi cette tâche en premier

`creer_cockpit_vide()` (Tâche 3) appelle `generer_cockpit(se_name, [], pilote_id, output_dir)`. Si `uo_list` est vide, la table n'est pas créée → `ajouter_uo_au_cockpit()` plante car `tbl_mes_uos` est absent.

- [ ] **Étape 1.1 : Lire la section concernée**

  Ouvrir `projet_TrainSystem/creer_cockpit_se.py`, repérer la fonction `_sheet_mes_uos()`.
  Bloc à modifier (lignes ~163-174) :

  ```python
  # AVANT
  last_row = row_h + len(uo_list)
  if uo_list:
      tbl = Table(displayName="tbl_mes_uos",
                  ref=f"A{row_h}:H{last_row}")
      tbl.tableStyleInfo = TableStyleInfo(
          name="TableStyleMedium2", showRowStripes=True)
      ws.add_table(tbl)
  ```

- [ ] **Étape 1.2 : Écrire le test qui échoue (table absente pour liste vide)**

  Ajouter dans `tests/test_assembler.py` :

  ```python
  import sys
  from pathlib import Path
  sys.path.insert(0, str(Path(__file__).parent.parent))
  sys.path.insert(0, str(Path(__file__).parent.parent / "projet_TrainSystem"))

  import pytest
  from openpyxl import load_workbook

  from creer_cockpit_se import generer_cockpit


  def test_cockpit_vide_has_table(tmp_path):
      """Un cockpit généré sans UO doit quand même avoir tbl_mes_uos."""
      out = generer_cockpit("Test Ingenieur", [], "USR004", tmp_path)
      wb = load_workbook(str(out))
      ws = wb["Mes UOs"]
      assert "tbl_mes_uos" in ws.tables
      tbl = ws.tables["tbl_mes_uos"]
      assert tbl.ref == "A3:H3"   # header seul, 0 lignes de données
  ```

- [ ] **Étape 1.3 : Vérifier que le test échoue**

  ```
  cd C:\Users\fabie\Documents\JLC\Python\SysEng
  python -m pytest tests/test_assembler.py::test_cockpit_vide_has_table -v
  ```

  Attendu : `FAILED — KeyError: tbl_mes_uos` ou `AssertionError`.

- [ ] **Étape 1.4 : Corriger `_sheet_mes_uos` dans `creer_cockpit_se.py`**

  Remplacer le bloc `if uo_list:` par :

  ```python
  last_row = row_h + max(len(uo_list), 0)
  tbl_ref = f"A{row_h}:H{last_row}" if uo_list else f"A{row_h}:H{row_h}"
  tbl = Table(displayName="tbl_mes_uos", ref=tbl_ref)
  tbl.tableStyleInfo = TableStyleInfo(
      name="TableStyleMedium2", showRowStripes=True)
  ws.add_table(tbl)
  ```

- [ ] **Étape 1.5 : Vérifier que le test passe**

  ```
  python -m pytest tests/test_assembler.py::test_cockpit_vide_has_table -v
  ```

  Attendu : `PASSED`.

- [ ] **Étape 1.6 : Ajouter le flag `--safe` dans `main()` de `creer_cockpit_se.py`**

  Après la ligne `p.add_argument("--output", ...)`, ajouter :

  ```python
  p.add_argument("--safe", action="store_true",
                 help="Ne régénère que les cockpits sans saisies (cols 5-6 nulles)")
  ```

  Et dans la boucle `for se_name, uos in sorted(par_ingenieur.items()):`, entourer l'appel :

  ```python
  for se_name, uos in sorted(par_ingenieur.items()):
      if args.safe:
          out_path = output_dir / f"Cockpit_{se_name.replace(' ', '_')}.xlsx"
          if out_path.exists() and _cockpit_has_saisies(out_path):
              print(f"  [SKIP --safe] {out_path.name} (saisies détectées)")
              continue
      out = generer_cockpit(se_name, uos, args.pilote, output_dir)
      print(f"  [OK] {out.name}  ({len(uos)} UO, pilote_id={args.pilote})")
  ```

  Et ajouter la fonction `_cockpit_has_saisies` **avant** `main()` :

  ```python
  def _cockpit_has_saisies(xlsx_path: Path) -> bool:
      """Retourne True si une cellule des colonnes 5 ou 6 de tbl_mes_uos est non nulle."""
      try:
          wb = load_workbook(str(xlsx_path), data_only=True, read_only=True)
          if "Mes UOs" not in wb.sheetnames:
              return False
          ws = wb["Mes UOs"]
          for row in ws.iter_rows(min_row=4, max_row=ws.max_row,
                                  min_col=5, max_col=6, values_only=True):
              if any(v is not None and v != 0 for v in row):
                  return True
          return False
      except Exception:
          return False
  ```

- [ ] **Étape 1.7 : Vérifier que les 382 tests passent toujours**

  ```
  python -m pytest tests/ -q
  ```

  Attendu : `382+ passed, 0 failed`.

- [ ] **Étape 1.8 : Commit**

  ```bash
  git add projet_TrainSystem/creer_cockpit_se.py tests/test_assembler.py
  git commit -m "fix: creer_cockpit_se toujours créer tbl_mes_uos + flag --safe"
  ```

---

## Tâche 2 — `ajouter_uo_au_cockpit()` dans `assembler.py`

**Fichiers :**
- Créer : `projet_TrainSystem/assembler.py` (section 1 : imports + cette fonction)
- Modifier : `tests/test_assembler.py` (ajouter les tests)

### Signature

```python
def ajouter_uo_au_cockpit(cockpit_path: Path, uo: dict) -> str:
    """
    Insère une ligne UO dans tbl_mes_uos sans écraser les saisies existantes.

    uo = {"file_id": "L09U1-CFL2400-CLIM", "systeme": "Climatisation",
          "projet": "CFL 2400", "heures": 200}

    Retourne "added" | "skipped" | "error:<msg>".
    Fait un backup .bak avant modification.
    """
```

### Contraintes

- Ne JAMAIS écrire dans les colonnes 5 et 6 de lignes existantes
- Si la ligne data à insérer atterrit sur une nouvelle row, écrire cols 1-4 uniquement + laisser 5-6 vides (l'ingénieur les remplira)
- Étendre la table ref d'une ligne
- Si file_id déjà dans col A des data rows → retourner `"skipped"`
- Faire `shutil.copy2(cockpit_path, cockpit_path.with_suffix('.bak'))` avant toute écriture

- [ ] **Étape 2.1 : Écrire les tests qui échouent**

  Ajouter dans `tests/test_assembler.py` :

  ```python
  import shutil
  from openpyxl.styles import PatternFill
  from creer_cockpit_se import generer_cockpit


  # ── Helpers tests ──────────────────────────────────────────────────────────────

  def _make_cockpit_with_saisies(tmp_path: Path, se_name: str = "Test SE") -> Path:
      """Crée un cockpit avec 1 UO et simule une saisie ingénieur (col 5 = 25%)."""
      out = generer_cockpit(se_name, [
          {"file_id": "L09U1-CFL2400-CLIM", "systeme": "Clim",
           "projet": "CFL", "heures": 200}
      ], "USR004", tmp_path)
      # Ouvrir et forcer col 5 de la data row à 0.25 (saisie ingénieur)
      wb = load_workbook(str(out))
      ws = wb["Mes UOs"]
      ws.cell(row=4, column=5, value=0.25)   # data row 1 = ligne 4
      wb.save(str(out))
      return out


  # ── Test 1 : colonnes jaunes préservées ────────────────────────────────────────

  def test_ajouter_uo_preserve_colonnes_jaunes(tmp_path):
      """ajouter_uo_au_cockpit ne doit pas toucher aux colonnes 5 et 6 existantes."""
      from assembler import ajouter_uo_au_cockpit

      cockpit = _make_cockpit_with_saisies(tmp_path)

      result = ajouter_uo_au_cockpit(cockpit, {
          "file_id": "L11U1-RERNG-FREIN",
          "systeme": "Frein", "projet": "RER NG", "heures": 150
      })

      assert result == "added"

      wb = load_workbook(str(cockpit))
      ws = wb["Mes UOs"]
      # La saisie initiale en ligne 4 doit être préservée
      assert ws.cell(row=4, column=5).value == pytest.approx(0.25)
      # La nouvelle UO est en ligne 5 (2e data row)
      assert ws.cell(row=5, column=1).value == "L11U1-RERNG-FREIN"
      # Cols 5-6 de la nouvelle ligne = vides (pas de saisie encore)
      assert ws.cell(row=5, column=5).value is None or ws.cell(row=5, column=5).value == 0
      assert ws.cell(row=5, column=6).value is None or ws.cell(row=5, column=6).value == 0


  # ── Test 2 : idempotence ──────────────────────────────────────────────────────

  def test_ajouter_uo_idempotent(tmp_path):
      """Appeler deux fois avec la même UO → 1 seule ligne, résultat 'skipped' au 2e appel."""
      from assembler import ajouter_uo_au_cockpit

      cockpit = generer_cockpit("Test SE", [], "USR004", tmp_path)
      uo = {"file_id": "L09U1-CFL2400-CLIM", "systeme": "Clim",
             "projet": "CFL", "heures": 200}

      r1 = ajouter_uo_au_cockpit(cockpit, uo)
      r2 = ajouter_uo_au_cockpit(cockpit, uo)

      assert r1 == "added"
      assert r2 == "skipped"

      wb = load_workbook(str(cockpit))
      ws = wb["Mes UOs"]
      # Compter les lignes avec file_id = L09U1-CFL2400-CLIM
      count = sum(
          1 for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=1, max_col=1)
          if row[0].value == "L09U1-CFL2400-CLIM"
      )
      assert count == 1


  # ── Test 3 : table étendue ────────────────────────────────────────────────────

  def test_ajouter_uo_etend_table(tmp_path):
      """La ref de tbl_mes_uos doit être étendue d'une ligne après ajout."""
      from assembler import ajouter_uo_au_cockpit

      cockpit = generer_cockpit("Test SE", [], "USR004", tmp_path)

      ajouter_uo_au_cockpit(cockpit, {
          "file_id": "L09U1-CFL2400-CLIM", "systeme": "Clim",
          "projet": "CFL", "heures": 200
      })

      wb = load_workbook(str(cockpit))
      ws = wb["Mes UOs"]
      assert "tbl_mes_uos" in ws.tables
      assert ws.tables["tbl_mes_uos"].ref == "A3:H4"   # header + 1 data row


  # ── Test 4 : backup créé ──────────────────────────────────────────────────────

  def test_ajouter_uo_cree_backup(tmp_path):
      """Un fichier .bak doit exister après modification."""
      from assembler import ajouter_uo_au_cockpit

      cockpit = generer_cockpit("Test SE", [], "USR004", tmp_path)
      ajouter_uo_au_cockpit(cockpit, {
          "file_id": "L09U1-CFL2400-CLIM", "systeme": "Clim",
          "projet": "CFL", "heures": 200
      })

      assert cockpit.with_suffix(".bak").exists()
  ```

- [ ] **Étape 2.2 : Vérifier que les tests échouent**

  ```
  python -m pytest tests/test_assembler.py::test_ajouter_uo_preserve_colonnes_jaunes tests/test_assembler.py::test_ajouter_uo_idempotent tests/test_assembler.py::test_ajouter_uo_etend_table tests/test_assembler.py::test_ajouter_uo_cree_backup -v
  ```

  Attendu : `4 FAILED — ImportError: cannot import name 'ajouter_uo_au_cockpit'`.

- [ ] **Étape 2.3 : Créer `projet_TrainSystem/assembler.py` avec la fonction**

  ```python
  """
  assembler.py — Instanciation industrialisée d'une UO (brique C4).

  Usage :
      python projet_TrainSystem/assembler.py L09U1 \\
          --projet CFL2400 --systeme CLIM \\
          --se "Alice Dubois" --pilote USR004 --heures 200 [--sync]
  """
  import argparse
  import re
  import shutil
  import sys
  from pathlib import Path

  from openpyxl import load_workbook
  from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
  from openpyxl.worksheet.table import Table, TableStyleInfo

  HERE = Path(__file__).parent
  ROOT = HERE.parent
  sys.path.insert(0, str(ROOT))
  sys.path.insert(0, str(HERE))

  RE_CODE = re.compile(r"^(L\d{2}U\d)(?:-([A-Za-z0-9]+))?(?:-([A-Za-z0-9]+))?$")

  # ── Helpers style (reproduits depuis creer_cockpit_se pour les nouvelles cellules) ──

  def _fill(hex_color: str) -> PatternFill:
      return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

  def _fnt(size=10, bold=False, color="000000") -> Font:
      return Font(name="Segoe UI", size=size, bold=bold, color=color)

  def _thin_border() -> Border:
      s = Side(style="thin", color="BFBFBF")
      return Border(left=s, right=s, top=s, bottom=s)

  def _center() -> Alignment:
      return Alignment(horizontal="center", vertical="center", wrap_text=True)

  def _left() -> Alignment:
      return Alignment(horizontal="left", vertical="center", wrap_text=True)


  # ── Fonction principale ──────────────────────────────────────────────────────

  def ajouter_uo_au_cockpit(cockpit_path: Path, uo: dict) -> str:
      """
      Insère une ligne UO dans tbl_mes_uos sans écraser les saisies existantes.

      uo = {"file_id": "L09U1-CFL2400-CLIM", "systeme": "Climatisation",
            "projet": "CFL 2400", "heures": 200}

      Retourne "added" | "skipped" | "error:<msg>".
      """
      try:
          wb = load_workbook(str(cockpit_path))
      except Exception as e:
          return f"error:{e}"

      if "Mes UOs" not in wb.sheetnames:
          return "error:feuille 'Mes UOs' absente"

      ws = wb["Mes UOs"]

      if "tbl_mes_uos" not in ws.tables:
          return "error:table tbl_mes_uos absente"

      tbl = ws.tables["tbl_mes_uos"]

      # Parser la ref : "A3:H4" → header_row=3, last_row=4
      ref = tbl.ref   # ex. "A3:H4" ou "A3:H3"
      top_ref, bot_ref = ref.split(":")
      header_row = int("".join(c for c in top_ref if c.isdigit()))
      last_row   = int("".join(c for c in bot_ref if c.isdigit()))

      # Lignes de données = [header_row+1 .. last_row] (vide si last_row == header_row)
      data_start = header_row + 1

      # Idempotence : chercher file_id dans col A des data rows
      file_id = uo["file_id"]
      for r in range(data_start, last_row + 1):
          if ws.cell(row=r, column=1).value == file_id:
              wb.close()
              return "skipped"

      # Backup avant modification
      shutil.copy2(str(cockpit_path), str(cockpit_path.with_suffix(".bak")))

      # Nouvelle ligne = last_row + 1 (ou data_start si table vide)
      new_row = last_row + 1 if last_row >= data_start else data_start

      # Alternance de couleur de fond (indice 0 = blanc, 1 = gris clair)
      row_idx = new_row - data_start   # 0-based
      bg = "F2F2F2" if row_idx % 2 else "FFFFFF"

      # Écrire cols 1-4 uniquement
      ws.cell(row=new_row, column=1, value=file_id)
      ws.cell(row=new_row, column=2, value=uo.get("systeme", ""))
      ws.cell(row=new_row, column=3, value=uo.get("projet", ""))
      ws.cell(row=new_row, column=4, value=uo.get("heures", 0))

      # Style cols 1-4
      for col in range(1, 5):
          c = ws.cell(row=new_row, column=col)
          c.fill = _fill(bg)
          c.border = _thin_border()
          c.alignment = _left() if col == 1 else _center()
          c.font = _fnt(9.5, color="0563C1" if col == 1 else "000000")

      # Cols 5-6 : cellules vides avec fill jaune (zones ingénieur)
      for col in (5, 6):
          c = ws.cell(row=new_row, column=col)
          c.fill = _fill("FFF2CC")
          c.border = _thin_border()
          c.alignment = _center()
      ws.cell(row=new_row, column=5).number_format = "0%"

      # Cols 7-8 : vides avec fond bg
      for col in (7, 8):
          c = ws.cell(row=new_row, column=col)
          c.fill = _fill(bg)
          c.border = _thin_border()
          c.alignment = _center()

      # Étendre la ref de la table
      new_ref = f"{top_ref}:{bot_ref[:1]}{new_row}"
      tbl.ref = new_ref

      wb.save(str(cockpit_path))
      wb.close()
      return "added"
  ```

- [ ] **Étape 2.4 : Vérifier que les tests passent**

  ```
  python -m pytest tests/test_assembler.py::test_ajouter_uo_preserve_colonnes_jaunes tests/test_assembler.py::test_ajouter_uo_idempotent tests/test_assembler.py::test_ajouter_uo_etend_table tests/test_assembler.py::test_ajouter_uo_cree_backup -v
  ```

  Attendu : `4 PASSED`.

- [ ] **Étape 2.5 : Régression globale**

  ```
  python -m pytest tests/ -q
  ```

  Attendu : `382+ passed, 0 failed` (les 4 nouveaux s'ajoutent).

- [ ] **Étape 2.6 : Commit**

  ```bash
  git add projet_TrainSystem/assembler.py tests/test_assembler.py
  git commit -m "feat: ajouter_uo_au_cockpit — patch chirurgical tbl_mes_uos"
  ```

---

## Tâche 3 — `creer_cockpit_vide()` dans `assembler.py`

**Fichiers :**
- Modifier : `projet_TrainSystem/assembler.py` (ajouter la fonction)
- Modifier : `tests/test_assembler.py` (ajouter 2 tests)

### Signature

```python
def creer_cockpit_vide(se_name: str, pilote_id: str, output_dir: Path) -> Path:
    """
    Génère Cockpit_{se_name}.xlsx avec tbl_mes_uos vide (0 lignes de données).
    Appelé automatiquement par l'assembleur si le cockpit n'existe pas.
    Retourne le chemin du cockpit créé.
    """
```

- [ ] **Étape 3.1 : Écrire les tests qui échouent**

  Ajouter dans `tests/test_assembler.py` :

  ```python
  # ── Test 5 : cockpit vide créé avec table ────────────────────────────────────

  def test_creer_cockpit_vide(tmp_path):
      """creer_cockpit_vide produit un cockpit avec tbl_mes_uos vide."""
      from assembler import creer_cockpit_vide

      out = creer_cockpit_vide("Alice Dubois", "USR004", tmp_path)

      assert out.exists()
      assert out.name == "Cockpit_Alice_Dubois.xlsx"
      wb = load_workbook(str(out))
      ws = wb["Mes UOs"]
      assert "tbl_mes_uos" in ws.tables
      assert ws.tables["tbl_mes_uos"].ref == "A3:H3"   # header seul


  # ── Test 6 : ajouter UO dans cockpit vide crée ──────────────────────────────

  def test_assembler_cockpit_inexistant_cree_automatiquement(tmp_path):
      """L'assembleur crée automatiquement le cockpit s'il n'existe pas."""
      from assembler import ajouter_uo_au_cockpit, creer_cockpit_vide

      cockpit = creer_cockpit_vide("Marie Dupont", "USR004", tmp_path)
      result = ajouter_uo_au_cockpit(cockpit, {
          "file_id": "L09U1-CFL2400-CLIM",
          "systeme": "Climatisation", "projet": "CFL 2400", "heures": 200
      })

      assert result == "added"
      wb = load_workbook(str(cockpit))
      ws = wb["Mes UOs"]
      assert ws.cell(row=4, column=1).value == "L09U1-CFL2400-CLIM"
  ```

- [ ] **Étape 3.2 : Vérifier que les tests échouent**

  ```
  python -m pytest tests/test_assembler.py::test_creer_cockpit_vide tests/test_assembler.py::test_assembler_cockpit_inexistant_cree_automatiquement -v
  ```

  Attendu : `2 FAILED — ImportError: cannot import name 'creer_cockpit_vide'`.

- [ ] **Étape 3.3 : Ajouter `creer_cockpit_vide` dans `assembler.py`**

  Ajouter après la fonction `ajouter_uo_au_cockpit` :

  ```python
  def creer_cockpit_vide(se_name: str, pilote_id: str, output_dir: Path) -> Path:
      """
      Génère Cockpit_{se_name}.xlsx avec tbl_mes_uos vide (0 lignes de données).
      """
      from creer_cockpit_se import generer_cockpit
      return generer_cockpit(se_name, [], pilote_id, output_dir)
  ```

- [ ] **Étape 3.4 : Vérifier que les tests passent**

  ```
  python -m pytest tests/test_assembler.py::test_creer_cockpit_vide tests/test_assembler.py::test_assembler_cockpit_inexistant_cree_automatiquement -v
  ```

  Attendu : `2 PASSED`.

- [ ] **Étape 3.5 : Régression globale**

  ```
  python -m pytest tests/ -q
  ```

  Attendu : `382+ passed, 0 failed`.

- [ ] **Étape 3.6 : Commit**

  ```bash
  git add projet_TrainSystem/assembler.py tests/test_assembler.py
  git commit -m "feat: creer_cockpit_vide — cockpit avec table vide pour nouvel ingénieur"
  ```

---

## Tâche 4 — Script principal `assembler.py` : CLI + test end-to-end

**Fichiers :**
- Modifier : `projet_TrainSystem/assembler.py` (ajouter `instancier_uo()` + `main()`)
- Modifier : `tests/test_assembler.py` (test end-to-end)

### Comportement attendu

```
python projet_TrainSystem/assembler.py L09U1 \
  --projet CFL2400 --systeme CLIM \
  --se "Alice Dubois" --pilote USR004 --heures 200

[OK] L09U1-CFL2400-CLIM.xlsx créé dans projet_TrainSystem/
[OK] Cockpit_Alice_Dubois.xlsx mis à jour (1 UO ajoutée)

python projet_TrainSystem/assembler.py L09U1 \
  --projet CFL2400 --systeme CLIM \
  --se "Alice Dubois" --pilote USR004 --heures 200  ← 2e appel

[SKIP] L09U1-CFL2400-CLIM.xlsx existe déjà — rien à faire.
```

- [ ] **Étape 4.1 : Écrire le test end-to-end qui échoue**

  Ajouter dans `tests/test_assembler.py` :

  ```python
  import subprocess

  # ── Test 7 : end-to-end (UO + cockpit + idempotence) ─────────────────────────

  def test_assembler_end_to_end(tmp_path):
      """
      L'assembleur crée l'UO, crée/met à jour le cockpit, et est idempotent.
      Utilise un répertoire temporaire (tmp_path) pour isoler les effets.
      """
      from assembler import instancier_uo

      # Créer un cockpit préexistant avec une saisie ingénieur (col 5 = 0.5)
      cockpit = generer_cockpit("Alice Dubois", [
          {"file_id": "L09U1-TEST01-CLIM", "systeme": "Clim",
           "projet": "TEST", "heures": 100}
      ], "USR004", tmp_path)
      wb = load_workbook(str(cockpit))
      ws = wb["Mes UOs"]
      ws.cell(row=4, column=5, value=0.5)
      wb.save(str(cockpit))

      # Appel 1 : doit créer l'UO et ajouter la ligne dans le cockpit
      result1 = instancier_uo(
          uo_type="L09U1",
          projet_code="CFL2400", systeme_code="CLIM",
          se_name="Alice Dubois", pilote_id="USR004",
          heures=200, output_dir=tmp_path, sync=False
      )
      assert result1["uo_status"] == "created"
      assert result1["cockpit_status"] == "added"

      uo_file = tmp_path / "L09U1-CFL2400-CLIM.xlsx"
      assert uo_file.exists()

      # La saisie initiale (col 5, ligne 4) doit être préservée
      wb = load_workbook(str(cockpit))
      ws = wb["Mes UOs"]
      assert ws.cell(row=4, column=5).value == pytest.approx(0.5)

      # La nouvelle UO est en ligne 5
      assert ws.cell(row=5, column=1).value == "L09U1-CFL2400-CLIM"

      # Appel 2 : idempotent
      result2 = instancier_uo(
          uo_type="L09U1",
          projet_code="CFL2400", systeme_code="CLIM",
          se_name="Alice Dubois", pilote_id="USR004",
          heures=200, output_dir=tmp_path, sync=False
      )
      assert result2["uo_status"] == "skipped"
      assert result2["cockpit_status"] == "skipped"
  ```

- [ ] **Étape 4.2 : Vérifier que le test échoue**

  ```
  python -m pytest tests/test_assembler.py::test_assembler_end_to_end -v
  ```

  Attendu : `FAILED — ImportError: cannot import name 'instancier_uo'`.

- [ ] **Étape 4.3 : Ajouter `instancier_uo()` et `main()` dans `assembler.py`**

  Ajouter à la fin de `assembler.py` :

  ```python
  def instancier_uo(
      uo_type: str, projet_code: str, systeme_code: str,
      se_name: str, pilote_id: str, heures: float,
      output_dir: Path, sync: bool = False
  ) -> dict:
      """
      Orchestre la création d'une UO et la mise à jour du cockpit.

      Retourne un dict :
        {"uo_status": "created"|"skipped", "cockpit_status": "added"|"skipped"|"created+added",
         "sync_push": int, "sync_errors": int}
      """
      code = f"{uo_type}-{projet_code}-{systeme_code}"
      uo_file = output_dir / f"{code}.xlsx"
      cockpit_file = output_dir / f"Cockpit_{se_name.replace(' ', '_')}.xlsx"

      # Étape 1 : UO
      uo_status = "skipped"
      if not uo_file.exists():
          import creer_uo
          import types

          # Construire un namespace args compatible avec build_instance
          args = types.SimpleNamespace(
              se=se_name, heures=heures,
              projet=projet_code, systeme=systeme_code,
              output=str(output_dir)
          )
          wb = creer_uo.build_instance(code, uo_type, projet_code, systeme_code, args)
          wb.save(str(uo_file))
          uo_status = "created"

      # Étape 2 : Cockpit
      cockpit_status_prefix = ""
      if not cockpit_file.exists():
          creer_cockpit_vide(se_name, pilote_id, output_dir)
          cockpit_status_prefix = "created+"

      uo_dict = {
          "file_id": code,
          "systeme": systeme_code,
          "projet":  projet_code,
          "heures":  heures,
      }
      add_result = ajouter_uo_au_cockpit(cockpit_file, uo_dict)
      cockpit_status = cockpit_status_prefix + add_result  # "added" | "skipped" | "created+added"

      # Étape 3 : Sync optionnelle
      sync_push = sync_errors = 0
      if sync:
          from src.sync import synchroniser_repertoire
          rapport_path = synchroniser_repertoire(output_dir)
          import json
          rapport = json.loads(rapport_path.read_text(encoding="utf-8"))
          sync_push = sum(
              len([l for l in r.get("log", []) if "PUSH=" in l])
              for r in rapport.get("fichiers", [])
          )
          sync_errors = rapport.get("nb_erreur", 0)

      return {
          "uo_status": uo_status,
          "cockpit_status": cockpit_status,
          "sync_push": sync_push,
          "sync_errors": sync_errors,
      }


  def main():
      p = argparse.ArgumentParser(
          description="Instancie une UO et met à jour le cockpit de l'ingénieur")
      p.add_argument("uo_type", help="ex: L09U1")
      p.add_argument("--projet",  required=True, help="Code projet ex: CFL2400")
      p.add_argument("--systeme", required=True, help="Code système ex: CLIM")
      p.add_argument("--se",      required=True, help="Nom ingénieur SE")
      p.add_argument("--pilote",  default="USR004", help="Pilote ID")
      p.add_argument("--heures",  type=float, default=0, help="Heures vendues")
      p.add_argument("--output",  default=str(HERE), help="Répertoire de sortie")
      p.add_argument("--sync",    action="store_true", help="Lancer la sync après création")
      args = p.parse_args()

      output_dir = Path(args.output)

      if not RE_CODE.match(f"{args.uo_type}-{args.projet}-{args.systeme}"):
          sys.exit(f"[ERR] Code invalide : {args.uo_type}-{args.projet}-{args.systeme}")

      code = f"{args.uo_type}-{args.projet}-{args.systeme}"
      uo_file = output_dir / f"{code}.xlsx"

      if uo_file.exists():
          print(f"[SKIP] {code}.xlsx existe déjà — rien à faire.")
          return

      result = instancier_uo(
          uo_type=args.uo_type,
          projet_code=args.projet,
          systeme_code=args.systeme,
          se_name=args.se,
          pilote_id=args.pilote,
          heures=args.heures,
          output_dir=output_dir,
          sync=args.sync,
      )

      print(f"[OK] {code}.xlsx créé dans {output_dir}/")
      cockpit_name = f"Cockpit_{args.se.replace(' ', '_')}.xlsx"
      if "added" in result["cockpit_status"]:
          print(f"[OK] {cockpit_name} mis à jour (1 UO ajoutée)")
      else:
          print(f"[SKIP] {cockpit_name} — UO déjà présente")
      if args.sync:
          print(f"[OK] Sync : {result['sync_push']} PUSH, {result['sync_errors']} erreur(s)")


  if __name__ == "__main__":
      main()
  ```

- [ ] **Étape 4.4 : Vérifier que le test passe**

  ```
  python -m pytest tests/test_assembler.py::test_assembler_end_to_end -v
  ```

  Attendu : `PASSED`.

- [ ] **Étape 4.5 : Régression globale**

  ```
  python -m pytest tests/ -q
  ```

  Attendu : `389+ passed, 0 failed` (7 nouveaux tests).

- [ ] **Étape 4.6 : Test manuel CLI (golden path)**

  ```powershell
  cd C:\Users\fabie\Documents\JLC\Python\SysEng
  # Copier d'abord Cockpit_Alice_Dubois.xlsx quelque part de sûr si déjà modifié
  python projet_TrainSystem/assembler.py L11U1 `
    --projet RERNG --systeme FREIN `
    --se "Alice Dubois" --pilote USR004 --heures 150
  ```

  Attendu :
  ```
  [OK] L11U1-RERNG-FREIN.xlsx créé dans projet_TrainSystem/
  [OK] Cockpit_Alice_Dubois.xlsx mis à jour (1 UO ajoutée)
  ```

- [ ] **Étape 4.7 : Test idempotence CLI**

  ```powershell
  python projet_TrainSystem/assembler.py L11U1 `
    --projet RERNG --systeme FREIN `
    --se "Alice Dubois" --pilote USR004 --heures 150
  ```

  Attendu :
  ```
  [SKIP] L11U1-RERNG-FREIN.xlsx existe déjà — rien à faire.
  ```

- [ ] **Étape 4.8 : Commit**

  ```bash
  git add projet_TrainSystem/assembler.py tests/test_assembler.py
  git commit -m "feat: assembler CLI — instanciation UO en < 5s, idempotent"
  ```

---

## Vérification finale

- [ ] **Test de performance**

  ```powershell
  Measure-Command {
    python projet_TrainSystem/assembler.py L09U2 `
      --projet CFL2400 --systeme TRACTION `
      --se "Bruno Lecomte" --pilote USR004 --heures 300
  }
  ```

  Attendu : `TotalSeconds < 5`.

- [ ] **Régression complète**

  ```
  python -m pytest tests/ -q
  ```

  Attendu : `389+ passed, 0 failed`.

- [ ] **Créer la branche et push**

  ```bash
  git checkout -b feature/assembleur
  git push -u origin feature/assembleur
  ```

---

## Self-review — couverture spec

| Requirement spec | Couvert par |
|------------------|-------------|
| Onglet Agenda dans le cockpit SE | Tâche 0 — `_sheet_agenda()` dans `creer_cockpit_se.py` |
| Génère L09U1-CFL2400-CLIM.xlsx | `instancier_uo()` appelle `creer_uo.build_instance()` |
| Ouvre cockpit existant, ajoute UNE ligne | `ajouter_uo_au_cockpit()` |
| Ne touche pas aux zones jaunes (cols 5-6) | `ajouter_uo_au_cockpit()` n'écrit jamais cols 5-6 existantes |
| Étend la plage de tbl_mes_uos | `tbl.ref = new_ref` |
| Backup .bak avant modification | `shutil.copy2()` |
| Skip si UO déjà dans table | Retourne `"skipped"` |
| Cockpit inexistant → créé auto | `creer_cockpit_vide()` |
| `--sync` lance synchroniser_repertoire() | `instancier_uo(sync=True)` |
| Résumé console | `main()` affiche [OK] / [SKIP] |
| `--safe` dans creer_cockpit_se | `_cockpit_has_saisies()` + flag --safe |
| Tests : colonnes jaunes préservées | `test_ajouter_uo_preserve_colonnes_jaunes` |
| Tests : idempotence | `test_ajouter_uo_idempotent` |
| Tests : cockpit inexistant | `test_assembler_cockpit_inexistant_cree_automatiquement` |
| Tests : end-to-end | `test_assembler_end_to_end` |
| < 5 secondes | Étape performance |
| 382+ tests passent toujours | Régression à chaque commit |
