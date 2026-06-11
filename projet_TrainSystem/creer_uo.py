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
  2. genere le classeur 11 feuilles (memes regles que le fichier UO type v1)
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

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

CATALOGUE = HERE / "Catalogue_UO_TrainSystem.xlsx"
NAVY = "1F4E79"
FONT = "Arial"
RE_CODE = re.compile(r"^(L\d{2}U\d)(?:-([A-Za-z0-9]+))?(?:-([A-Za-z0-9]+))?$")


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


# ─── Helpers de construction ──────────────────────────────────────────────────

def style_header(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", wrap_text=True)


def add_table(ws, name, headers, rows, widths=None):
    for ci, h in enumerate(headers, 1):
        ws.cell(row=1, column=ci, value=h)
    for ri, rd in enumerate(rows, 2):
        for ci, h in enumerate(headers, 1):
            ws.cell(row=ri, column=ci, value=rd.get(h))
    style_header(ws, 1, 1, len(headers))
    end = f"{get_column_letter(len(headers))}{1 + max(len(rows), 1)}"
    t = Table(displayName=name, ref=f"A1:{end}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(t)
    if widths:
        for col, w in widths.items():
            ws.column_dimensions[col].width = w


def named(wb, name, sheet, cell):
    wb.defined_names[name] = DefinedName(name, attr_text=f"'{sheet}'!${cell[0]}${cell[1:]}")


# ─── Construction de l'instance ───────────────────────────────────────────────

def build_instance(code, uo_type, projet_code, systeme_code, args):
    info, cat_acts, cat_livs, cat_des = load_catalogue(uo_type)
    wb = Workbook()

    # General
    ws = wb.active
    ws.title = "General"
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 50
    ws["A1"] = f"UO — {code}"
    ws["A1"].font = Font(name=FONT, bold=True, size=14)
    rows = [
        ("Pôle", "Train Système"),
        ("Lot", f"LOT {info['lot']} — {info['lot_libelle']}"),
        ("UO", f"UO{uo_type[-1]} — {info['uo_libelle']}"),
        ("Libellé", info["lot_libelle"]),
        ("Projet", args.projet or projet_code or ""),
        ("Système", args.systeme or systeme_code or ""),
        ("Ingénieur Système (SE)", args.se),
        ("Heures vendues", args.heures),
        ("Date de création", datetime.date.today().isoformat()),
    ]
    for i, (k, v) in enumerate(rows, 3):
        ws.cell(row=i, column=1, value=k).font = Font(name=FONT, bold=True)
        ws.cell(row=i, column=2, value=v)
    named(wb, "heures_vendues", "General", "B10")

    # Description_Besoin — texte complet du catalogue
    ws = wb.create_sheet("Description_Besoin")
    ws.column_dimensions["A"].width = 110
    ws["A1"] = f"Cahier des charges — Catalogue UO {uo_type} (copié à la création)"
    ws["A1"].font = Font(name=FONT, bold=True, size=12)
    r = 3
    for act in cat_acts:
        c = ws.cell(row=r, column=1, value=f"{act['id']} : {act['designation']}")
        c.font = Font(name=FONT, bold=True)
        r += 1
        for line in (act.get("description") or "").split("\n"):
            if line.strip():
                ws.cell(row=r, column=1, value="   " + line).alignment = \
                    Alignment(wrap_text=True)
                r += 1
        r += 1
    if info.get("criteres_acceptation"):
        ws.cell(row=r, column=1, value="Critères d'acceptation").font = \
            Font(name=FONT, bold=True)
        r += 1
        for line in str(info["criteres_acceptation"]).split("\n"):
            ws.cell(row=r, column=1, value="   " + line)
            r += 1

    # Donnees_Entree
    ws = wb.create_sheet("Donnees_Entree")
    add_table(ws, "tbl_donnees_entree",
        ["id", "designation", "type", "origine", "pic", "statut",
         "date_reception", "maturite", "date_update", "commentaire"],
        [{"id": f"DE-{i:03d}", "designation": d["designation"], "statut": "EN_ATTENTE"}
         for i, d in enumerate(cat_des, 1)],
        widths={"A": 9, "B": 50, "C": 10, "D": 10, "F": 12, "G": 13, "H": 9,
                "I": 12, "J": 28})
    dv = DataValidation(type="list", formula1='"NA,EN_ATTENTE,RECUE"', allow_blank=True)
    ws.add_data_validation(dv); dv.add("F2:F80")

    # Activites
    ws = wb.create_sheet("Activites")
    headers = ["id", "designation", "applicable", "poids", "heures_allouees",
               "statut", "avancement", "heures_consommees", "reste_a_faire",
               "commentaire"]
    rows = [{"id": a["id"], "designation": a["designation"], "applicable": "OUI",
             "poids": a.get("poids") or 1, "statut": "A_FAIRE", "avancement": 0,
             "heures_consommees": 0} for a in cat_acts]
    add_table(ws, "tbl_activites", headers, rows,
              widths={"A": 10, "B": 46, "C": 11, "D": 8, "E": 15, "F": 11,
                      "G": 12, "H": 17, "I": 13, "J": 32})
    n = len(rows)
    for r in range(2, 2 + n):
        ws.cell(row=r, column=5).value = (
            f'=IF(C{r}="OUI",D{r}/SUMIF($C$2:$C${1+n},"OUI",$D$2:$D${1+n})'
            f"*heures_vendues,0)")
        ws.cell(row=r, column=9).value = f"=(1-G{r}/100)*E{r}"
        ws.cell(row=r, column=5).number_format = "0.00"
        ws.cell(row=r, column=9).number_format = "0.00"
    dv = DataValidation(type="list", formula1='"OUI,NON"', allow_blank=True)
    ws.add_data_validation(dv); dv.add("C2:C60")
    dv = DataValidation(type="list", formula1='"A_FAIRE,EN_COURS,TERMINEE,STAND_BY"',
                        allow_blank=True)
    ws.add_data_validation(dv); dv.add("F2:F60")

    # Livrables
    ws = wb.create_sheet("Livrables")
    add_table(ws, "tbl_livrables",
        ["id", "designation", "type", "maturite_attendue", "date_attendue",
         "pic", "statut", "date_revisee", "date_livraison", "commentaire"],
        [{"id": f"LIV-{i:03d}", "designation": l["designation"], "statut": "A_FAIRE"}
         for i, l in enumerate(cat_livs, 1)],
        widths={"A": 9, "B": 50, "C": 8, "D": 16, "E": 14, "F": 10, "G": 11,
                "H": 13, "I": 13, "J": 28})
    dv = DataValidation(type="list", formula1='"A_FAIRE,EN_COURS,LIVRE,VALIDE"',
                        allow_blank=True)
    ws.add_data_validation(dv); dv.add("G2:G60")

    # OIL — vide (1 ligne exemple a remplacer)
    ws = wb.create_sheet("OIL")
    add_table(ws, "tbl_oil",
        ["id", "titre", "description", "en_action", "domaine", "criticite",
         "statut", "date_ouverture", "date_besoin", "journal"],
        [{"id": "PO-000", "titre": "(exemple — remplacer par ton premier point)",
          "description": "", "en_action": "SE", "criticite": "BASSE",
          "statut": "CLOS", "journal": ""}],
        widths={"A": 8, "B": 34, "C": 44, "D": 13, "E": 12, "F": 11, "G": 9,
                "H": 13, "I": 12, "J": 50})
    dv = DataValidation(type="list",
                        formula1='"SE,FOURNISSEUR,EXPERT,AT,CLIENT,AUTRE"',
                        allow_blank=True)
    ws.add_data_validation(dv); dv.add("D2:D60")
    dv = DataValidation(type="list", formula1='"BASSE,MOYENNE,HAUTE"', allow_blank=True)
    ws.add_data_validation(dv); dv.add("F2:F60")
    dv = DataValidation(type="list", formula1='"OUVERT,CLOS"', allow_blank=True)
    ws.add_data_validation(dv); dv.add("G2:G60")

    # KPI
    ws = wb.create_sheet("KPI")
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 64
    ws["A1"] = "KPI — calculés par ExoSync à chaque synchronisation (ne pas saisir)"
    ws["A1"].font = Font(name=FONT, bold=True, size=12)
    kpis = [
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
    r = 3
    for label, name, fmt, desc in kpis:
        ws.cell(row=r, column=1, value=label).font = Font(name=FONT, bold=True)
        named(wb, name, "KPI", f"B{r}")
        ws.cell(row=r, column=2).number_format = fmt
        ws.cell(row=r, column=3, value=desc).font = Font(name=FONT, size=9, italic=True)
        r += 1
    excel_kpis = [
        ("Heures vendues", "kpi_h_vendues", "=heures_vendues", "0",
         "Lu depuis General (formule Excel)"),
        ("Reste à faire total (h)", "kpi_raf", "=SUM(Activites!I2:I60)", "0.00",
         "Somme colonne reste_a_faire (formule Excel)"),
        ("Heures estimées à terminaison (EAC)", "kpi_eac", "=kpi_h_conso+kpi_raf",
         "0.00", "EAC = consommé + reste à faire (formule Excel)"),
        ("Dérive à terminaison (h)", "kpi_derive", "=kpi_eac-kpi_h_vendues", "0.00",
         "EAC − heures vendues : >0 = dépassement prévu (formule Excel)"),
        ("Santé", "kpi_sante",
         '=IF(kpi_po_critiques>0,"ROUGE",IF(kpi_po_ouverts>0,"ORANGE","VERT"))',
         "General",
         "ROUGE si point critique ouvert · ORANGE si point ouvert · VERT sinon"),
    ]
    for label, name, formula, fmt, desc in excel_kpis:
        ws.cell(row=r, column=1, value=label).font = Font(name=FONT, bold=True)
        ws.cell(row=r, column=2, value=formula)
        ws.cell(row=r, column=2).number_format = fmt
        named(wb, name, "KPI", f"B{r}")
        ws.cell(row=r, column=3, value=desc).font = Font(name=FONT, size=9, italic=True)
        r += 1

    # Dashboard
    ws = wb.create_sheet("Dashboard")
    ws["A1"] = f"Dashboard — {code} · {args.se}"
    ws["A1"].font = Font(name=FONT, bold=True, size=14)
    ws.column_dimensions["A"].width = 26
    for col in "BCDEFG":
        ws.column_dimensions[col].width = 16
    cards = [
        ("A4", "AVANCEMENT UO",  "A6", '=ROUND(kpi_avancement,2)&" %"'),
        ("C4", "HEURES",         "C6", '=ROUND(kpi_h_conso,2)&" / "&heures_vendues&" h"'),
        ("E4", "POINTS OUVERTS", "E6", "=kpi_po_ouverts"),
        ("G4", "ATTENTE FOURN.", "G6", "=kpi_po_fournisseur"),
        ("A8", "SANTÉ",          "A10", "=kpi_sante"),
        ("C8", "EAC (h)",        "C10", "=ROUND(kpi_eac,2)"),
        ("E8", "DÉRIVE (h)",     "E10", "=ROUND(kpi_derive,2)"),
        ("G8", "PTS CRITIQUES",  "G10", "=kpi_po_critiques"),
    ]
    for lc, label, vc, formula in cards:
        ws[lc] = label
        ws[lc].font = Font(name=FONT, bold=True, size=10, color=NAVY)
        ws[vc] = formula
        ws[vc].font = Font(name=FONT, bold=True, size=18)
    ws["A13"] = ("Cette feuille est libre : l'ingénieur la customise "
                 "(elle ne fait que LIRE l'onglet KPI).")
    ws["A13"].font = Font(name=FONT, italic=True, size=9)

    # Planning / Orga
    ws = wb.create_sheet("Planning")
    ws["A1"] = "Réservé : visualisation planning."
    ws["A1"].font = Font(name=FONT, italic=True)
    ws = wb.create_sheet("Orga")
    ws["A1"] = "Organisation projet (statique v1 — à renseigner depuis le Catalogue Projets)."
    ws["A1"].font = Font(name=FONT, italic=True)

    # _Manifeste
    ws = wb.create_sheet("_Manifeste")
    ws.column_dimensions["A"].width = 95
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
