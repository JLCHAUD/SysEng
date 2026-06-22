"""Tests TDD pour le module de design system xl_design."""
from openpyxl.styles import Font, PatternFill, Border

from src.xl_design import XD


class TestPrimitives:
    def test_font_family_segoe(self):
        f = XD.fnt(12, bold=True, color="FFFFFF")
        assert isinstance(f, Font)
        assert f.name == "Segoe UI"
        assert f.size == 12
        assert f.bold is True
        assert f.color.rgb.endswith("FFFFFF")

    def test_font_defaut(self):
        f = XD.fnt()
        assert f.size == 10
        assert f.bold is False
        assert f.color.rgb.endswith("2C2C2A")

    def test_fill_solide(self):
        fill = XD.fill("0C447C")
        assert isinstance(fill, PatternFill)
        assert fill.fgColor.rgb.endswith("0C447C")

    def test_input_jaune_constant(self):
        assert XD.INPUT == "FFF2CC"

    def test_hair_border(self):
        assert isinstance(XD.HAIR, Border)
        assert XD.HAIR.left.style == "thin"

    def test_alignements(self):
        assert XD.center().horizontal == "center"
        assert XD.center().wrap_text is True
        assert XD.left().horizontal == "left"


class TestRegistreOnglets:
    def test_onze_familles(self):
        assert len(XD.SHEETS) == 11

    def test_cles_attendues(self):
        attendues = {"general", "dashboard", "description", "planning",
                     "donnees_entree", "activites", "livrables", "oil",
                     "kpi", "orga", "manifeste"}
        assert set(XD.SHEETS) == attendues

    def test_triple_activites(self):
        s = XD.sheet("activites")
        assert s.banner == "084434"
        assert s.header == "0C5E49"
        assert s.accent == "E1F5EE"
        assert s.glyph == "✔"

    def test_triple_oil_rouge(self):
        assert XD.sheet("oil").header == "A32D2D"

    def test_cle_inconnue_leve(self):
        import pytest
        with pytest.raises(KeyError):
            XD.sheet("inexistant")

    def test_tab_colors_mappe_le_ton_moyen(self):
        tc = XD.tab_colors()
        assert tc["general"] == "0C447C"
        assert tc["kpi"] == "534AB7"
