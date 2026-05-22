"""Tests du mot-clé IDENT dans le parser MXL."""
import pytest
from src.parser import parse_lines, ManifestAST


def test_parse_ident_basic():
    """IDENT avec LABEL et valeur col B."""
    lines = [
        ("FILE_TYPE: uo_instance", ""),
        ("FILE_ID:   UO-001", ""),
        ("IDENT nom : LABEL=\"Nom de l'UO\"", "Jean Dupont"),
        ("IDENT responsable : LABEL=\"Responsable\"", "Marie Martin"),
        ("DEF $activites = GET_TABLE(Activites, TabActivites)", ""),
    ]
    ast = parse_lines(lines)
    assert len(ast.idents) == 2
    assert ast.idents[0].name == "nom"
    assert ast.idents[0].label == "Nom de l'UO"
    assert ast.idents[0].value == "Jean Dupont"
    assert ast.idents[1].name == "responsable"
    assert ast.idents[1].value == "Marie Martin"
    assert ast.errors == []


def test_ident_not_captured_as_metadata():
    """IDENT ne doit PAS être stocké dans manifest_metadata."""
    lines = [("IDENT nom : LABEL=\"Nom\"", "Jean")]
    ast = parse_lines(lines)
    assert "nom" not in ast.header.manifest_metadata
    assert "ident" not in ast.header.manifest_metadata
    assert len(ast.idents) == 1


def test_ident_label_fallback_to_name():
    """Sans LABEL=, le label prend la valeur de name."""
    lines = [("IDENT site :", "Paris")]
    ast = parse_lines(lines)
    assert ast.idents[0].name == "site"
    assert ast.idents[0].label == "site"
    assert ast.idents[0].value == "Paris"


def test_no_ident_is_valid():
    """Un manifeste sans IDENT est valide — ast.idents est vide."""
    lines = [("FILE_TYPE: uo_instance", ""), ("DEF $t = GET_TABLE(S, T)", "")]
    ast = parse_lines(lines)
    assert ast.idents == []
    assert ast.errors == []


def test_multiple_idents():
    """Plusieurs IDENT sur un même Post."""
    lines = [
        ("IDENT nom : LABEL=\"Nom\"", "Alice"),
        ("IDENT site : LABEL=\"Site\"", "Paris"),
        ("IDENT region : LABEL=\"Région\"", "Île-de-France"),
    ]
    ast = parse_lines(lines)
    assert len(ast.idents) == 3
    assert [i.name for i in ast.idents] == ["nom", "site", "region"]
    assert [i.value for i in ast.idents] == ["Alice", "Paris", "Île-de-France"]
