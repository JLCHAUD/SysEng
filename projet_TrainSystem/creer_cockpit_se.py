"""
creer_cockpit_se.py — Génère un cockpit ingénieur à partir des UO du répertoire.

Scanne tous les fichiers L*.xlsx, extrait l'ingénieur depuis la feuille General,
groupe les UO par ingénieur et génère un cockpit Excel par ingénieur.

Usage :
    python projet_TrainSystem/creer_cockpit_se.py
    python projet_TrainSystem/creer_cockpit_se.py --pilote USR004
"""
import argparse
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
sys.path.insert(0, str(Path(__file__).parent))

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

HERE = Path(__file__).parent


# ─── Lecture des UO ──────────────────────────────────────────────────────────

def lire_uo_du_repertoire(dossier: Path) -> list[dict]:
    """Scanne les UO xlsx et retourne la liste de leurs métadonnées."""
    uo_list = []
    for xlsx in sorted(dossier.glob("L*.xlsx")):
        if xlsx.name.startswith("~"):
            continue
        try:
            wb = load_workbook(str(xlsx), data_only=True)
        except Exception:
            continue

        if "General" not in wb.sheetnames or "_Manifeste" not in wb.sheetnames:
            wb.close()
            continue

        ws_g = wb["General"]
        se_name = heures = systeme = projet = None
        for row in ws_g.iter_rows(min_row=1, max_row=30, values_only=True):
            if not row[1]:
                continue
            label = str(row[1])
            if "Ing" in label and "SE" in label:
                se_name = row[2]
            elif "Heures" in label:
                heures = row[2]
            elif "Syst" in label:
                systeme = row[2]
            elif "Projet" in label:
                projet = row[2]

        if se_name:
            uo_list.append({
                "file_id": xlsx.stem,
                "se_name": str(se_name),
                "heures":  heures or 0,
                "systeme": systeme or "",
                "projet":  projet or "",
            })
        wb.close()
    return uo_list


# ─── Helpers style ────────────────────────────────────────────────────────────

def _fill(hex_color: str) -> PatternFill:
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

def _fnt(size=10, bold=False, color="000000", italic=False) -> Font:
    return Font(name="Segoe UI", size=size, bold=bold, color=color, italic=italic)

def _center() -> Alignment:
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def _left() -> Alignment:
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def _thin_border() -> Border:
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)


# ─── Génération du cockpit ────────────────────────────────────────────────────

def generer_cockpit(se_name: str, uo_list: list[dict],
                    pilote_id: str, output_dir: Path) -> Path:
    """Génère un fichier cockpit Excel pour un ingénieur."""
    wb = Workbook()
    wb.remove(wb.active)

    _sheet_mes_uos(wb, se_name, uo_list)
    _sheet_manifeste(wb, se_name, uo_list, pilote_id)

    safe = se_name.replace(" ", "_")
    out = output_dir / f"Cockpit_{safe}.xlsx"
    wb.save(str(out))
    return out


def _sheet_mes_uos(wb: Workbook, se_name: str, uo_list: list[dict]):
    ws = wb.create_sheet("Mes UOs")
    ws.sheet_view.showGridLines = False

    # Bannière
    ws.merge_cells("A1:J1")
    c = ws["A1"]
    c.value = f"Cockpit Ingenieur — {se_name}   |   {date.today().strftime('%d/%m/%Y')}"
    c.fill = _fill("0C447C")
    c.font = _fnt(13, bold=True, color="FFFFFF")
    c.alignment = _center()
    ws.row_dimensions[1].height = 30

    # En-têtes table
    headers = ["UO ID", "Système", "Projet", "Charge (h)",
               "% Avancement", "H réalisées", "Date fin", "Alerte"]
    row_h = 3
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row_h, column=col, value=h)
        cell.fill = _fill("1F3864")
        cell.font = _fnt(9.5, bold=True, color="FFFFFF")
        cell.alignment = _center()
        cell.border = _thin_border()
    ws.row_dimensions[row_h].height = 20

    # Lignes UO
    for i, uo in enumerate(uo_list):
        row = row_h + 1 + i
        bg = "F2F2F2" if i % 2 else "FFFFFF"

        ws.cell(row=row, column=1, value=uo["file_id"]).font = _fnt(9.5, color="0563C1")
        ws.cell(row=row, column=2, value=uo["systeme"])
        ws.cell(row=row, column=3, value=uo["projet"])
        ws.cell(row=row, column=4, value=uo["heures"])

        # Colonnes jaunes (saisie ingénieur)
        for col in (5, 6):
            c = ws.cell(row=row, column=col, value=0)
            c.fill = _fill("FFF2CC")
            c.border = _thin_border()
            c.alignment = _center()
        ws.cell(row=row, column=5).number_format = "0%"

        ws.cell(row=row, column=7, value="")   # Date fin
        ws.cell(row=row, column=8, value="")   # Alerte

        for col in range(1, 9):
            c = ws.cell(row=row, column=col)
            if col not in (5, 6):
                c.fill = _fill(bg)
            c.border = _thin_border()
            if col not in (1,):
                c.alignment = _center()
            else:
                c.alignment = _left()

    # Table nommée (requise pour GET_TABLE dans le manifeste)
    last_row = row_h + len(uo_list)
    if uo_list:
        tbl = Table(displayName="tbl_mes_uos",
                    ref=f"A{row_h}:H{last_row}")
        tbl.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tbl)

    # Largeurs colonnes
    for col, w in zip("ABCDEFGH", [20, 18, 22, 12, 16, 14, 14, 22]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = f"A{row_h + 1}"


def _sheet_manifeste(wb: Workbook, se_name: str,
                     uo_list: list[dict], pilote_id: str):
    ws = wb.create_sheet("_Manifeste")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 65
    ws.column_dimensions["C"].width = 60

    def w(row, instr, comment=""):
        c = ws.cell(row=row, column=1, value=instr)
        bold = any(instr.startswith(k) for k in ("DEF ", "PUSH ", "PULL ", "LIST "))
        c.font = Font(name="Calibri", size=9.5, bold=bold, color="1F3864" if bold else "000000")
        c.alignment = Alignment(horizontal="left")
        if comment:
            cc = ws.cell(row=row, column=3, value=comment)
            cc.font = Font(name="Calibri", size=9, color="666666", italic=True)

    safe = se_name.replace(" ", "_")

    ws["A1"] = "MANIFESTE_V=1"
    ws["A1"].font = Font(name="Calibri", size=10, bold=True, color="1F3864")

    r = 3
    w(r, "FILE_TYPE: cockpit_ingenieur",  "Type de fichier ExoSync"); r += 1
    w(r, f"FILE_ID: Cockpit_{safe}",      "Identifiant unique du cockpit"); r += 1
    w(r, f"ingenieur: {se_name}",         "Nom de l'ingénieur propriétaire"); r += 1
    if pilote_id:
        w(r, f"pilote_id: {pilote_id}",
          "Pilote responsable — permet au dashboard de découvrir ce cockpit"); r += 1

    r += 1  # séparateur
    w(r, "# ── Lecture de la table des UOs ────────────────────────────"); r += 1
    w(r, "DEF $mes_uos = GET_TABLE(Mes UOs, tbl_mes_uos)",
      "Lit la table des UOs avec les saisies ingénieur (zones jaunes)"); r += 1

    r += 1
    w(r, "# ── Publication vers le store central ──────────────────────"); r += 1
    w(r, f"PUSH $mes_uos -> cockpit.{safe}.mes_uos",
      "Remonte la table UOs (avec avancements) vers le store central ExoSync")


# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    p = argparse.ArgumentParser(
        description="Génère des cockpits ingénieur depuis les UO du répertoire")
    p.add_argument("--pilote", default="USR004",
                   help="Pilote ID à inscrire dans le manifeste (défaut: USR004)")
    p.add_argument("--ingenieur", default="",
                   help="Filtrer un seul ingénieur (optionnel)")
    p.add_argument("--output", default=str(HERE),
                   help="Répertoire de sortie (défaut: projet_TrainSystem)")
    args = p.parse_args()

    output_dir = Path(args.output)
    uo_list = lire_uo_du_repertoire(HERE)

    if not uo_list:
        sys.exit("[ERR] Aucune UO trouvée dans le répertoire.")

    # Grouper par ingénieur
    par_ingenieur: dict[str, list] = {}
    for uo in uo_list:
        se = uo["se_name"]
        if args.ingenieur and se != args.ingenieur:
            continue
        par_ingenieur.setdefault(se, []).append(uo)

    if not par_ingenieur:
        sys.exit(f"[ERR] Aucun ingénieur trouvé{' pour ' + args.ingenieur if args.ingenieur else ''}.")

    print(f"UOs trouvées : {len(uo_list)}")
    for se_name, uos in sorted(par_ingenieur.items()):
        out = generer_cockpit(se_name, uos, args.pilote, output_dir)
        print(f"  [OK] {out.name}  ({len(uos)} UO, pilote_id={args.pilote})")


if __name__ == "__main__":
    main()
