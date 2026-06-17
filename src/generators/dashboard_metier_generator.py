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

    ws.merge_cells("A1:J1")
    t = ws["A1"]
    t.value = (f"Dashboard Métier — {acteur.nom}   |   "
               f"{date.today().strftime('%d/%m/%Y')}")
    t.fill = solid_fill(BLUE_DARK)
    t.font = header_font(size=14)
    t.alignment = center()
    ws.row_dimensions[1].height = 32

    total_h = int(sum(u.total_hours for u in uo_list))
    avancement_vals = [_get_store_float(store, f"uo.{u.id}.avancement") for u in uo_list]
    avg_avanc = sum(avancement_vals) / len(avancement_vals) if avancement_vals else 0

    kpis = [
        ("A2:B2", len(uo_list), "Périmètre équipe"),
        ("C2:D2", total_h, "Charge totale"),
        ("E2:F2", avg_avanc, "Avancement moyen"),
    ]
    for cell_range, value, label in kpis:
        ws.merge_cells(cell_range)
        first_cell = cell_range.split(":")[0]
        c = ws[first_cell]
        c.value = value
        c.fill = solid_fill(BLUE_LIGHT)
        c.font = body_font(bold=True, color="1F3864")
        c.alignment = center()
        c.border = THIN_BORDER
    # Label row under KPIs
    ws.merge_cells("A3:B3")
    ws["A3"].value = "Périmètre équipe"
    ws.merge_cells("C3:D3")
    ws["C3"].value = "Charge totale (h)"
    ws.merge_cells("E3:F3")
    ws["E3"].value = "Avancement moyen"
    for col_label in ("A3", "C3", "E3"):
        ws[col_label].font = body_font(color="555555")
        ws[col_label].alignment = center()
    ws.row_dimensions[2].height = 24
    ws.row_dimensions[3].height = 18

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

        ws.merge_cells(f"A{current_row}:H{current_row}")
        eng_cell = ws[f"A{current_row}"]
        eng_cell.value = (f"▶  {eng}   |   {len(eng_uo)} UO   |   "
                          f"{total_h}h   |   {avg_avanc:.0%} avancement moyen")
        eng_cell.fill = solid_fill(BLUE_MID)
        eng_cell.font = header_font()
        eng_cell.alignment = left()
        ws.row_dimensions[current_row].height = 22
        current_row += 1

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
        synth_row = 6 + i

        rules = [
            ("CELL_PCT",  "GLOBAL", f"uo.{uo.id}.avancement",      "", "Synthèse", "", "", "", f"G{synth_row}", "PULL", "",
             f"Récupère l'avancement de {uo.id} poussé par le cockpit de {uo.engineer_name}"),
            ("CELL_NUM",  "GLOBAL", f"uo.{uo.id}.heures_realisees", "", "Synthèse", "", "", "", f"H{synth_row}", "PULL", "",
             f"Récupère les heures réalisées de {uo.id} poussées par le cockpit de {uo.engineer_name}"),
            ("CELL_NUM",  "GLOBAL", f"uo.{uo.id}.points_ouverts",  "", "Alertes",  "", "", "", "",             "PULL", "",
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

        row += 1

    set_column_widths(ws, {
        "A": 12, "B": 8, "C": 35, "D": 14, "E": 12, "F": 10,
        "G": 8, "H": 10, "I": 10, "J": 10, "K": 10, "L": 60,
    })
