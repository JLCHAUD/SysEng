"""
tests/test_assembler.py — Tests TDD pour assembler.py et les modifications de creer_cockpit_se.py.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "projet_TrainSystem"))

import pytest
from openpyxl import load_workbook

from creer_cockpit_se import generer_cockpit


# ── Tâche 0 : onglet Agenda ───────────────────────────────────────────────────

def test_cockpit_has_agenda_sheet(tmp_path):
    """Le cockpit doit avoir un onglet 'Agenda' avec les 3 sections."""
    out = generer_cockpit("Alice Dubois", [
        {"file_id": "L09U1-CFL2400-CLIM", "systeme": "Clim",
         "projet": "CFL", "heures": 200}
    ], "USR004", tmp_path)

    wb = load_workbook(str(out))
    assert "Agenda" in wb.sheetnames

    ws = wb["Agenda"]
    all_values = [ws.cell(row=r, column=c).value
                  for r in range(1, ws.max_row + 1)
                  for c in range(1, 3)]
    all_str = [str(v) for v in all_values if v]
    assert any("Semaine" in s for s in all_str)
    assert any("30" in s for s in all_str)
    assert any("ouverts" in s.lower() or "action" in s.lower() for s in all_str)


# ── Tâche 1 : table toujours créée ───────────────────────────────────────────

def test_cockpit_vide_has_table(tmp_path):
    """Un cockpit généré sans UO doit quand même avoir tbl_mes_uos."""
    out = generer_cockpit("Test Ingenieur", [], "USR004", tmp_path)
    wb = load_workbook(str(out))
    ws = wb["Mes UOs"]
    assert "tbl_mes_uos" in ws.tables
    tbl = ws.tables["tbl_mes_uos"]
    assert tbl.ref == "A3:H3"   # header seul, 0 lignes de données
