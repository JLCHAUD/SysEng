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
    GREEN_L = "EAF3DE"
    GREEN_D = "27500A"
    BLUE_L = "E6F1FB"
    NAVY_D = "0C447C"
    AMBER_L = "FAEEDA"
    AMBER_D = "854F0B"
    RED_L = "FCEBEB"
    RED_D = "791F1F"
    GREY_L = "F1EFE8"
    GREY_D = "5F5E5A"
    GREY_B = "D3D1C7"

    # ── Bordure fine (4 côtés) ─────────────────────────────────
    _SIDE = Side(style="thin", color="D3D1C7")
    HAIR = Border(left=_SIDE, right=_SIDE, top=_SIDE, bottom=_SIDE)

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
