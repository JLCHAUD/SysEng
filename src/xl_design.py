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
    GREEN_L = "EAF3DE"; GREEN_D = "27500A"
    BLUE_L = "E6F1FB";  NAVY_D = "0C447C"
    AMBER_L = "FAEEDA"; AMBER_D = "854F0B"
    RED_L = "FCEBEB";   RED_D = "791F1F"
    GREY_L = "F1EFE8";  GREY_D = "5F5E5A"; GREY_B = "D3D1C7"

    # ── Bordure fine (4 côtés) ─────────────────────────────────
    _SIDE = Side(style="thin", color="D3D1C7")
    HAIR = Border(left=_SIDE, right=_SIDE, top=_SIDE, bottom=_SIDE)

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
