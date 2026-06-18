# src/generators/cockpit_ingenieur_generator.py
"""Cockpit ingénieur système — Vue Agenda + Mes UOs + _Manifeste."""
from datetime import date, timedelta
from pathlib import Path
from typing import List

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.table import Table, TableStyleInfo

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
    uo_dir: Path = UO_DIR,
) -> Path:
    """Génère le cockpit agenda pour un ingénieur système."""
    output_dir.mkdir(parents=True, exist_ok=True)
    uo_list = [u for u in all_uo_instances if u.engineer_name == engineer_name]

    wb = Workbook()
    wb.remove(wb.active)

    _sheet_mes_uos(wb, engineer_name, uo_list, uo_dir)
    _sheet_agenda(wb, engineer_name, uo_list)
    _sheet_manifeste(wb, engineer_name, uo_list)

    safe = engineer_name.replace(" ", "_")
    filepath = output_dir / f"Cockpit_{safe}.xlsx"
    wb.save(filepath)
    return filepath


def _sheet_mes_uos(wb: Workbook, engineer_name: str, uo_list: List[UOInstance], uo_dir: Path = UO_DIR):
    ws = wb.create_sheet("Mes UOs")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:I1")
    t = ws["A1"]
    t.value = f"Mes UOs — {engineer_name}   |   {date.today().strftime('%d/%m/%Y')}"
    t.fill = solid_fill(BLUE_DARK)
    t.font = header_font(size=13)
    t.alignment = center()
    ws.row_dimensions[1].height = 30

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

    ws.merge_cells("A4:I4")
    sec = ws["A4"]
    sec.value = "Toutes mes UO"
    sec.fill = solid_fill(BLUE_MID)
    sec.font = header_font()
    sec.alignment = center()

    # A=UO ID  B=Type  C=Système  D=Projet  E=Charge(h)  F=% Avancement  G=H réalisées  H=Date fin  I=Alerte
    headers = ["UO ID", "Type UO", "Système", "Projet", "Charge (h)",
               "% Avancement", "H réalisées", "Date fin", "Alerte"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=5, column=col, value=h)
    style_header_row(ws, 5, 1, 9, color=BLUE_MID)

    for i, uo in enumerate(uo_list):
        row = 6 + i
        type_name = uo.uo_type.name if uo.uo_type else uo.uo_type_id
        sys_name = uo.system.name if uo.system else uo.system_id
        proj_name = uo.project.name if uo.project else uo.project_id

        id_cell = ws.cell(row=row, column=1, value=uo.id)
        id_cell.font = body_font(color="0563C1")
        id_cell.alignment = left()
        id_cell.border = THIN_BORDER
        uo_path = uo_dir / f"{uo.id}.xlsx"
        id_cell.hyperlink = str(uo_path)

        ws.cell(row=row, column=2, value=type_name)
        ws.cell(row=row, column=3, value=sys_name)
        ws.cell(row=row, column=4, value=proj_name)
        ws.cell(row=row, column=5, value=uo.total_hours)

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

        alert_formula = (
            f'=IF(G{row}>E{row},"⚠ Dérive heures",'
            f'IF(AND(H{row}<>"",H{row}<TODAY()+7),"⏰ Échéance proche","✅ OK"))'
        )
        ws.cell(row=row, column=9, value=alert_formula)

        style_data_row(ws, row, 2, 5, alternate=(i % 2 == 1))
        style_data_row(ws, row, 8, 9, alternate=(i % 2 == 1))

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

    # Table nommée pour GET_TABLE(Mes UOs, tbl_mes_uos)
    if uo_list:
        tbl_ref = f"A5:I{last_row}"
        tbl = Table(displayName="tbl_mes_uos", ref=tbl_ref)
        tbl.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2", showFirstColumn=False,
            showLastColumn=False, showRowStripes=True, showColumnStripes=False,
        )
        ws.add_table(tbl)

    set_column_widths(ws, {
        "A": 12, "B": 30, "C": 18, "D": 20, "E": 13,
        "F": 16, "G": 14, "H": 14, "I": 22,
    })
    ws.freeze_panes = "A6"


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
    ws.merge_cells(f"A{start_row}:F{start_row}")
    sec = ws[f"A{start_row}"]
    sec.value = title
    sec.fill = solid_fill(color)
    sec.font = header_font(color=header_text_color)
    sec.alignment = left()
    start_row += 1

    headers = ["UO ID", "Activité", "Priorité", "Date échéance", "Statut", "Action"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=start_row, column=col, value=h)
    style_header_row(ws, start_row, 1, 6, color=BLUE_LIGHT)
    for col in range(1, 7):
        ws.cell(row=start_row, column=col).font = body_font(bold=True, color="1F3864")
    start_row += 1

    row = start_row
    for uo in uo_list:
        activities = uo.uo_type.activities if uo.uo_type else []
        for idx, act in enumerate(activities):
            act_end = getattr(act, 'end_date', None) or uo.end_date
            if not act_end or not (date_from <= act_end <= date_to):
                continue

            ws.cell(row=row, column=1, value=uo.id)
            ws.cell(row=row, column=2, value=act.name)

            priorite = "🔴 Haute" if act_end <= date.today() + timedelta(days=3) else "🟡 Normale"
            ws.cell(row=row, column=3, value=priorite)

            date_cell = ws.cell(row=row, column=4, value=act_end)
            date_cell.number_format = "DD/MM/YYYY"

            act_statut = getattr(act, 'statut', None)
            statut_val = act_statut.value if hasattr(act_statut, "value") else str(act_statut) if act_statut else ""
            ws.cell(row=row, column=5, value=statut_val)
            ws.cell(row=row, column=6, value="")

            style_data_row(ws, row, 1, 6, alternate=(row % 2 == 0))
            row += 1

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

    # Ligne 1 : version
    ws["A1"] = "MANIFESTE_V=1"
    ws["A1"].font = body_font(bold=True, color="1F3864")
    # Ligne 2 intentionnellement vide (skippée par parser.py)

    # Métadonnées
    _mxl_row(3, "FILE_TYPE: cockpit_ingenieur",     comment="Type de fichier ExoSync")
    _mxl_row(4, f"FILE_ID: Cockpit_{safe_name}",    comment="Identifiant unique du cockpit")
    _mxl_row(5, f"ingenieur: {engineer_name}",      comment="Nom de l'ingénieur propriétaire")

    # Définition de la table
    _mxl_row(7, "DEF $mes_uos = GET_TABLE(Mes UOs, tbl_mes_uos)",
             comment="Référence à la table des UOs de l'ingénieur")
    _mxl_row(8, "COL $mes_uos.avancement : WRITE=engineer",
             comment="% avancement saisi par l'ingénieur (zone jaune)")
    _mxl_row(9, "COL $mes_uos.heures_realisees : WRITE=engineer",
             comment="Heures réalisées saisies par l'ingénieur (zone jaune)")

    # Export vers store
    _mxl_row(11, f"PUSH $mes_uos -> cockpit.{safe_name}.mes_uos",
             comment="Remonte les saisies ingénieur vers le store central ExoSync")

    set_column_widths(ws, {"A": 60, "B": 18, "C": 55})
