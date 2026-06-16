"""
creer_uo.py — Assembleur d'instanciation d'UO (brique C4).
===========================================================
Cree un fichier UO complet et synchronisable a partir du Catalogue :

    python projet_TrainSystem/creer_uo.py L09U1-CFL2400-CLIM \\
        --se "Jean Dujardin" --heures 200 \\
        [--projet "CFL 2400"] [--systeme "Climatisation"] [--output DIR]

Le code se decompose : L{lot}U{uo}-{PROJET}-{SYSTEME}
  L09U1  -> entree du Catalogue (activites, livrables, donnees d'entree)
  CFL2400 -> code projet (affiche dans General ; --projet pour le libelle)
  CLIM    -> code systeme (--systeme pour le libelle)

L'assembleur :
  1. lit Catalogue_UO_TrainSystem.xlsx (filtre uo_type)
  2. genere le classeur 11 feuilles avec Design B (bandeau marine/teal/amber)
  3. pre-remplit Activites / Livrables / Donnees_Entree depuis le catalogue
  4. genere le _Manifeste MXL avec les bonnes cles uo.<code>.*
"""
import argparse
import datetime
import re
import sys
from pathlib import Path

HERE = Path(__file__).parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(HERE))

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from design_b import (
    NAVY_D, BLUE, GREY_D, GREY_B, AMBER_L, AMBER_D, RED_L, RED_D, WHITE,
    F, HAIR,
    fnt, fill, add_table as _add_table,
    statut_cf, criticite_cf,
    banner_B, banner_teal, banner_amber,
    section_box, kpi_card_B,
)

CATALOGUE = HERE / "Catalogue_UO_TrainSystem.xlsx"
RE_CODE = re.compile(r"^(L\d{2}U\d)(?:-([A-Za-z0-9]+))?(?:-([A-Za-z0-9]+))?$")

# Banner occupies rows 1-4; table headers and content start at row T.
T = 5


# ─── Lecture du catalogue ─────────────────────────────────────────────────────

def read_table(ws, table_name):
    ref = ws.tables[table_name].ref
    rows = list(ws[ref])
    headers = [c.value for c in rows[0]]
    return [dict(zip(headers, (c.value for c in r))) for r in rows[1:]
            if any(c.value is not None for c in r)]


def load_catalogue(uo_type):
    if not CATALOGUE.exists():
        sys.exit(f"[ERR] Catalogue introuvable : {CATALOGUE}\n"
                 "      Lancez d'abord : python projet_TrainSystem/build_catalogue.py")
    wb = load_workbook(CATALOGUE, data_only=False)
    index = {r["uo_type"]: r for r in read_table(wb["Index"], "tbl_index")}
    if uo_type not in index:
        sys.exit(f"[ERR] UO type '{uo_type}' absent du catalogue. "
                 f"Disponibles : {', '.join(sorted(index))}")
    acts = [r for r in read_table(wb["Catalogue_Activites"], "tbl_cat_activites")
            if r["uo_type"] == uo_type]
    livs = [r for r in read_table(wb["Catalogue_Livrables"], "tbl_cat_livrables")
            if r["uo_type"] == uo_type]
    des = [r for r in read_table(wb["Catalogue_DonneesEntree"], "tbl_cat_donnees")
           if r["uo_type"] == uo_type]
    return index[uo_type], acts, livs, des


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _write_table(ws, name, headers, rows, widths=None, start_row=T):
    """Ecrit entetes + donnees, applique le style Design B et enregistre le tableau."""
    hr = start_row
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=hr, column=ci, value=h)
        c.fill = fill(NAVY_D)
        c.font = fnt(10, bold=True, color=WHITE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = HAIR
    ws.row_dimensions[hr].height = 22
    for ri, rd in enumerate(rows, hr + 1):
        for ci, h in enumerate(headers, 1):
            c = ws.cell(row=ri, column=ci, value=rd.get(h))
            c.border = HAIR
            c.font = fnt(10)
        ws.row_dimensions[ri].height = 20
    last = hr + max(len(rows), 1)
    _add_table(ws, name, f"A{hr}:{get_column_letter(len(headers))}{last}")
    if widths:
        for col, w in widths.items():
            ws.column_dimensions[col].width = w
    return hr + 1, last  # (data_start_row, data_end_row)


def named(wb, name, sheet, cell):
    wb.defined_names[name] = DefinedName(
        name, attr_text=f"'{sheet}'!${cell[0]}${cell[1:]}")


def _dv(ws, type_, formula, rng):
    dv = DataValidation(type=type_, formula1=formula, allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(rng)


# ─── Construction de l'instance ───────────────────────────────────────────────

def build_instance(code, uo_type, projet_code, systeme_code, args):
    info, cat_acts, cat_livs, cat_des = load_catalogue(uo_type)
    wb = Workbook()

    uo_title = f"UO {uo_type} — {info['uo_libelle']}"
    parts = [args.projet or projet_code or "", args.systeme or systeme_code or ""]
    proj_line = "  ·  ".join(p for p in parts if p)
    se = args.se or ""
    bkw = dict(title=uo_title, project_line=proj_line, se=se)

    # ── General ──────────────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "General"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 52
    banner_B(ws, "Fiche générale", 8, **bkw)
    kv = [
        ("Pôle",                    "Train Système"),
        ("Lot",                     f"LOT {info['lot']} — {info['lot_libelle']}"),
        ("UO",                      f"UO{uo_type[-1]} — {info['uo_libelle']}"),
        ("Libellé",                 info["lot_libelle"]),
        ("Projet",                  args.projet or projet_code or ""),
        ("Système",                 args.systeme or systeme_code or ""),
        ("Ingénieur Système (SE)",  se),
        ("Heures vendues",          args.heures),
        ("Date de création",        datetime.date.today().isoformat()),
    ]
    for i, (k, v) in enumerate(kv, T + 1):
        ws.cell(row=i, column=1, value=k).font = fnt(10, bold=True, color=NAVY_D)
        ws.cell(row=i, column=2, value=v).font = fnt(10)
        ws.row_dimensions[i].height = 20
    # "Heures vendues" is the 8th item (index 7) → row T+1+7 = T+8
    named(wb, "heures_vendues", "General", f"B{T + 8}")

    # ── Description_Besoin ───────────────────────────────────────────────────
    ws = wb.create_sheet("Description_Besoin")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 110
    banner_B(ws, "Cahier des charges", 6, **bkw)
    r = T + 2
    for act in cat_acts:
        c = ws.cell(row=r, column=1, value=f"{act['id']} : {act['designation']}")
        c.font = fnt(11, bold=True, color=NAVY_D)
        r += 1
        for line in (act.get("description") or "").split("\n"):
            if line.strip():
                c = ws.cell(row=r, column=1, value="   " + line)
                c.font = fnt(10)
                c.alignment = Alignment(wrap_text=True)
                r += 1
        r += 1
    if info.get("criteres_acceptation"):
        ws.cell(row=r, column=1, value="Critères d'acceptation").font = \
            fnt(11, bold=True, color=NAVY_D)
        r += 1
        for line in str(info["criteres_acceptation"]).split("\n"):
            ws.cell(row=r, column=1, value="   " + line).font = fnt(10)
            r += 1

    # ── Donnees_Entree ───────────────────────────────────────────────────────
    ws = wb.create_sheet("Donnees_Entree")
    ws.sheet_view.showGridLines = False
    banner_teal(ws, "Données d'entrée", 10, **bkw)
    _write_table(ws, "tbl_donnees_entree",
        ["id", "designation", "type", "origine", "pic", "statut",
         "date_reception", "maturite", "date_update", "commentaire"],
        [{"id": f"DE-{i:03d}", "designation": d["designation"], "statut": "EN_ATTENTE"}
         for i, d in enumerate(cat_des, 1)],
        widths={"A": 9, "B": 50, "C": 10, "D": 10, "F": 12, "G": 13,
                "H": 9, "I": 12, "J": 28})
    _dv(ws, "list", '"NA,EN_ATTENTE,RECUE"', f"F{T+1}:F{T+80}")

    # ── Activites ────────────────────────────────────────────────────────────
    ws = wb.create_sheet("Activites")
    ws.sheet_view.showGridLines = False
    banner_teal(ws, "Activités", 10, **bkw)
    act_headers = ["id", "designation", "applicable", "poids", "heures_allouees",
                   "statut", "avancement", "heures_consommees", "reste_a_faire",
                   "commentaire"]
    act_rows = [{"id": a["id"], "designation": a["designation"],
                 "applicable": "OUI", "poids": a.get("poids") or 1,
                 "statut": "A_FAIRE", "avancement": 0, "heures_consommees": 0}
                for a in cat_acts]
    _write_table(ws, "tbl_activites", act_headers, act_rows,
                 widths={"A": 10, "B": 46, "C": 11, "D": 8, "E": 15, "F": 11,
                         "G": 12, "H": 17, "I": 13, "J": 32})
    n = len(act_rows)
    dr_s, dr_e = T + 1, T + n
    for r in range(dr_s, dr_e + 1):
        ws.cell(row=r, column=5).value = (
            f'=IF(C{r}="OUI",D{r}/SUMIF($C${dr_s}:$C${dr_e},"OUI"'
            f',$D${dr_s}:$D${dr_e})*heures_vendues,0)')
        ws.cell(row=r, column=9).value = f"=(1-G{r}/100)*E{r}"
        ws.cell(row=r, column=5).number_format = "0.00"
        ws.cell(row=r, column=9).number_format = "0.00"
    _dv(ws, "list", '"OUI,NON"', f"C{T+1}:C{T+60}")
    _dv(ws, "list", '"A_FAIRE,EN_COURS,TERMINEE,STAND_BY"', f"F{T+1}:F{T+60}")
    from openpyxl.formatting.rule import DataBarRule
    ws.conditional_formatting.add(
        f"G{T+1}:G{T+n}",
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100,
                    color=BLUE, showValue=True))
    statut_cf(ws, f"F{T+1}:F{T+n}")
    ws.freeze_panes = f"A{T+1}"

    # ── Livrables ─────────────────────────────────────────────────────────────
    ws = wb.create_sheet("Livrables")
    ws.sheet_view.showGridLines = False
    banner_teal(ws, "Livrables", 10, **bkw)
    liv_rows = [{"id": f"LIV-{i:03d}", "designation": l["designation"],
                 "statut": "A_FAIRE"}
                for i, l in enumerate(cat_livs, 1)]
    _write_table(ws, "tbl_livrables",
        ["id", "designation", "type", "maturite_attendue", "date_attendue",
         "pic", "statut", "date_revisee", "date_livraison", "commentaire"],
        liv_rows,
        widths={"A": 9, "B": 50, "C": 8, "D": 16, "E": 14, "F": 10,
                "G": 11, "H": 13, "I": 13, "J": 28})
    _dv(ws, "list", '"A_FAIRE,EN_COURS,LIVRE,VALIDE"', f"G{T+1}:G{T+60}")
    statut_cf(ws, f"G{T+1}:G{T+len(liv_rows)}")

    # ── OIL ──────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("OIL")
    ws.sheet_view.showGridLines = False
    banner_amber(ws, "Points ouverts — OIL", 10, **bkw)
    _write_table(ws, "tbl_oil",
        ["id", "titre", "description", "en_action", "domaine", "criticite",
         "statut", "date_ouverture", "date_besoin", "journal"],
        [{"id": "PO-000", "titre": "(exemple — remplacer par ton premier point)",
          "description": "", "en_action": "SE", "criticite": "BASSE",
          "statut": "CLOS", "journal": ""}],
        widths={"A": 8, "B": 34, "C": 44, "D": 13, "E": 12, "F": 11,
                "G": 9, "H": 13, "I": 12, "J": 50})
    _dv(ws, "list", '"SE,FOURNISSEUR,EXPERT,AT,CLIENT,AUTRE"', f"D{T+1}:D{T+60}")
    _dv(ws, "list", '"BASSE,MOYENNE,HAUTE"', f"F{T+1}:F{T+60}")
    _dv(ws, "list", '"OUVERT,CLOS"', f"G{T+1}:G{T+60}")
    criticite_cf(ws, f"F{T+1}:F{T+60}")
    from openpyxl.formatting.rule import CellIsRule
    ws.conditional_formatting.add(f"G{T+1}:G{T+60}", CellIsRule(
        operator="equal", formula=['"OUVERT"'],
        fill=fill(RED_L), font=Font(name=F, size=10, bold=True, color=RED_D)))
    ws.conditional_formatting.add(f"G{T+1}:G{T+60}", CellIsRule(
        operator="equal", formula=['"CLOS"'],
        fill=fill("EAF3DE"), font=Font(name=F, size=10, color="27500A")))

    # ── KPI ──────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("KPI")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 64
    banner_B(ws, "KPI", 8, **bkw)

    section_box(ws, "KPI calculés par ExoSync (ne pas saisir manuellement)",
                T, 1, T + 7, 3)
    exo_kpis = [
        ("Avancement UO (%)",           "kpi_avancement",     "0.00",
         "Moyenne des avancements pondérée par les poids (activités applicables)"),
        ("Heures consommées",           "kpi_h_conso",        "0.00",
         "Somme des heures consommées (activités applicables)"),
        ("Points ouverts",              "kpi_po_ouverts",     "0",
         "Nombre de points OIL au statut OUVERT"),
        ("Points fermés",               "kpi_po_fermes",      "0",
         "Nombre de points OIL au statut CLOS"),
        ("Points critiques ouverts",    "kpi_po_critiques",   "0",
         "Points OUVERTS avec criticite = HAUTE"),
        ("Dont balle chez fournisseur", "kpi_po_fournisseur", "0",
         "Points OUVERTS avec en_action = FOURNISSEUR"),
        ("Dont balle chez expert",      "kpi_po_expert",      "0",
         "Points OUVERTS avec en_action = EXPERT"),
    ]
    r = T + 1
    for label, nm, fmt, desc in exo_kpis:
        ws.cell(row=r, column=1, value=label).font = fnt(10, bold=True, color=NAVY_D)
        named(wb, nm, "KPI", f"B{r}")
        ws.cell(row=r, column=2).number_format = fmt
        ws.cell(row=r, column=3, value=desc).font = fnt(9, color=GREY_D, italic=True)
        ws.row_dimensions[r].height = 20
        r += 1
    # r = T + 8 here; leave row T+8 as spacer, Excel section starts at T+9
    r += 1
    section_box(ws, "Indicateurs Excel (formules — s'actualisent à l'ouverture du fichier)",
                r, 1, r + 5, 3)
    r += 1
    excel_kpis = [
        ("Heures vendues",                    "kpi_h_vendues",
         "=heures_vendues",                  "0",
         "Lu depuis General (formule Excel)"),
        ("Reste à faire total (h)",           "kpi_raf",
         f"=SUM(Activites!I{T+1}:I{T+60})", "0.00",
         "Somme colonne reste_a_faire (formule Excel)"),
        ("Heures estimées à terminaison (EAC)", "kpi_eac",
         "=kpi_h_conso+kpi_raf",             "0.00",
         "EAC = consommé + reste à faire (formule Excel)"),
        ("Dérive à terminaison (h)",          "kpi_derive",
         "=kpi_eac-kpi_h_vendues",           "0.00",
         "EAC − heures vendues : >0 = dépassement prévu (formule Excel)"),
        ("Santé",                             "kpi_sante",
         '=IF(kpi_po_critiques>0,"ROUGE",IF(kpi_po_ouverts>0,"ORANGE","VERT"))',
         "General",
         "ROUGE si point critique · ORANGE si point ouvert · VERT sinon"),
    ]
    for label, nm, formula, fmt, desc in excel_kpis:
        ws.cell(row=r, column=1, value=label).font = fnt(10, bold=True, color=NAVY_D)
        ws.cell(row=r, column=2, value=formula)
        ws.cell(row=r, column=2).number_format = fmt
        named(wb, nm, "KPI", f"B{r}")
        ws.cell(row=r, column=3, value=desc).font = fnt(9, color=GREY_D, italic=True)
        ws.row_dimensions[r].height = 20
        r += 1

    # ── Dashboard ─────────────────────────────────────────────────────────────
    ws = wb.create_sheet("Dashboard")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    for i in range(2, 18):
        ws.column_dimensions[get_column_letter(i)].width = 8.5
    banner_B(ws, "Cockpit de pilotage", 16, **bkw)

    ROW1, ROW2 = T + 1, T + 6  # two rows of KPI cards
    for rr in range(ROW1, ROW1 + 4):
        ws.row_dimensions[rr].height = 18
    for rr in range(ROW2, ROW2 + 4):
        ws.row_dimensions[rr].height = 18

    cards_r1 = [
        (2,  "AVANCEMENT UO",  '=ROUND(kpi_avancement,1)&" %"',
         "pondéré par poids",   "B5D4F4", NAVY_D),
        (6,  "HEURES",
         '=ROUND(kpi_h_conso,0)&" / "&heures_vendues&" h"',
         "consommées",          "D3D1C7", "2C2C2A"),
        (10, "POINTS OUVERTS", "=kpi_po_ouverts",
         "au statut OUVERT",    "FAC775", AMBER_D),
        (14, "PTS CRITIQUES",  "=kpi_po_critiques",
         "criticité HAUTE",     "FCEBEB", RED_D),
    ]
    cards_r2 = [
        (2,  "SANTÉ",          "=kpi_sante",
         "",                    "F5F4F0", NAVY_D),
        (6,  "EAC (h)",        "=ROUND(kpi_eac,0)",
         "estimé à terminaison","D3D1C7", "2C2C2A"),
        (10, "DÉRIVE (h)",     "=ROUND(kpi_derive,0)",
         "EAC − heures vendues","F5F4F0", GREY_D),
        (14, "BALLE FOURN.",   "=kpi_po_fournisseur",
         "pts chez fournisseur","FAEEDA", AMBER_D),
    ]
    for col, label, value, sub, bc, vc in cards_r1:
        kpi_card_B(ws, col, label, value, sub, bc, vc, row=ROW1)
    for col, label, value, sub, bc, vc in cards_r2:
        kpi_card_B(ws, col, label, value, sub, bc, vc, row=ROW2)

    note_row = ROW2 + 5
    note = ws.cell(row=note_row, column=2,
                   value="Cette feuille est libre — l'ingénieur la complète "
                         "(elle ne lit que l'onglet KPI).")
    note.font = fnt(9, color=GREY_D, italic=True)
    ws.merge_cells(start_row=note_row, start_column=2,
                   end_row=note_row, end_column=16)

    # ── Planning ──────────────────────────────────────────────────────────────
    ws = wb.create_sheet("Planning")
    ws.sheet_view.showGridLines = False
    banner_B(ws, "Planning", 10, **bkw)
    ws.cell(row=T + 2, column=1,
            value="Réservé : visualisation planning.").font = \
        fnt(10, color=GREY_D, italic=True)

    # ── Orga ──────────────────────────────────────────────────────────────────
    ws = wb.create_sheet("Orga")
    ws.sheet_view.showGridLines = False
    banner_B(ws, "Organisation", 10, **bkw)
    ws.cell(row=T + 2, column=1,
            value="Organisation projet (à renseigner depuis le Catalogue Projets).").font = \
        fnt(10, color=GREY_D, italic=True)

    # ── _Manifeste ────────────────────────────────────────────────────────────
    ws = wb.create_sheet("_Manifeste")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 95
    ws.sheet_properties.tabColor = "888888"
    ws["A1"] = "MANIFESTE_V=1"
    ws["A2"] = "instruction"
    ws["A2"].font = Font(bold=True)
    manifeste = [
        "FILE_TYPE: uo_instance",
        f"FILE_ID: {code}",
        "",
        "# ── Lecture des donnees locales ─────────────────────────────",
        "DEF $act = GET_TABLE(Activites, tbl_activites)",
        "DEF $oil = GET_TABLE(OIL, tbl_oil)",
        "DEF $liv = GET_TABLE(Livrables, tbl_livrables)",
        "",
        "# ── Sous-ensembles ──────────────────────────────────────────",
        'DEF $actifs = COMPUTE(FILTER($act, applicable = "OUI"))',
        'DEF $po_ouv = COMPUTE(FILTER($oil, statut = "OUVERT"))',
        "",
        "# ── KPI ─────────────────────────────────────────────────────",
        "DEF $avancement = COMPUTE(MEAN_WEIGHTED($actifs.avancement, $actifs.poids))",
        "DEF $h_conso = COMPUTE(SUM($actifs.heures_consommees))",
        'DEF $po_ouverts = COMPUTE(COUNT_IF($oil.statut, "OUVERT"))',
        'DEF $po_fermes = COMPUTE(COUNT_IF($oil.statut, "CLOS"))',
        'DEF $po_critiques = COMPUTE(COUNT_IF($po_ouv.criticite, "HAUTE"))',
        'DEF $po_fournisseur = COMPUTE(COUNT_IF($po_ouv.en_action, "FOURNISSEUR"))',
        'DEF $po_expert = COMPUTE(COUNT_IF($po_ouv.en_action, "EXPERT"))',
        "",
        "# ── Regles de qualite ───────────────────────────────────────",
        "VALIDATE $actifs.avancement : RANGE(0, 100)",
        'VALIDATE $act.applicable : IN("OUI", "NON")',
        'VALIDATE $oil.statut : IN("OUVERT", "CLOS")',
        "",
        "# ── Affichage local (onglet KPI) ────────────────────────────",
        "BIND $avancement -> KPI.kpi_avancement",
        "BIND $h_conso -> KPI.kpi_h_conso",
        "BIND $po_ouverts -> KPI.kpi_po_ouverts",
        "BIND $po_fermes -> KPI.kpi_po_fermes",
        "BIND $po_critiques -> KPI.kpi_po_critiques",
        "BIND $po_fournisseur -> KPI.kpi_po_fournisseur",
        "BIND $po_expert -> KPI.kpi_po_expert",
        "",
        "# ── Publication vers l'ecosysteme ───────────────────────────",
        f"PUSH $avancement -> uo.{code}.avancement",
        f"PUSH $h_conso -> uo.{code}.heures_consommees",
        f"PUSH $po_ouverts -> uo.{code}.po_ouverts",
        f"PUSH $po_fermes -> uo.{code}.po_fermes",
        f"PUSH $po_critiques -> uo.{code}.po_critiques",
        f"PUSH $actifs -> uo.{code}.activites",
        f"PUSH $oil -> uo.{code}.points_ouverts",
        f"PUSH $liv -> uo.{code}.livrables",
    ]
    for i, line in enumerate(manifeste, 3):
        ws.cell(row=i, column=1, value=line)

    return wb


def main():
    p = argparse.ArgumentParser(description="Cree une UO depuis le Catalogue")
    p.add_argument("code", help="ex: L09U1-CFL2400-CLIM")
    p.add_argument("--se", default="", help="Nom de l'ingenieur systeme")
    p.add_argument("--heures", type=float, default=0, help="Heures vendues")
    p.add_argument("--projet", default="", help="Libelle du projet")
    p.add_argument("--systeme", default="", help="Libelle du systeme")
    p.add_argument("--output", default=str(HERE), help="Repertoire de sortie")
    args = p.parse_args()

    m = RE_CODE.match(args.code)
    if not m:
        sys.exit(f"[ERR] Code invalide '{args.code}' — attendu : L09U1-PROJET-SYSTEME")
    uo_type, projet_code, systeme_code = m.groups()

    wb = build_instance(args.code, uo_type, projet_code, systeme_code, args)
    out = Path(args.output) / f"{args.code}.xlsx"
    wb.save(out)
    print(f"[OK] {out}")
    print(f"     UO type {uo_type} — SE: {args.se or '<a assigner>'} — "
          f"{args.heures:g} h vendues")
    print(f"     Synchroniser : python scripts/valider_un.py "
          f"{out.relative_to(ROOT) if out.is_relative_to(ROOT) else out}")


if __name__ == "__main__":
    main()
