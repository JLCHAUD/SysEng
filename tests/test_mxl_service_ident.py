"""Tests de génération IDENT dans build_class_mxl_lines."""
from web.mxl_service import build_class_mxl_lines


def test_is_key_true_generates_ident_line():
    """min_fields avec is_key=True → ligne IDENT dans la sortie."""
    ft = {
        "min_fields": [
            {"name": "nom", "label": "Nom de l'UO", "is_key": True},
            {"name": "statut", "label": "Statut", "is_key": False, "source": "user_input"},
        ]
    }
    lines = build_class_mxl_lines("uo", "UO-001", ft, [])
    ident_lines = [l for l in lines if l.startswith("IDENT")]
    assert len(ident_lines) == 1
    assert ident_lines[0] == 'IDENT nom : LABEL="Nom de l\'UO"'


def test_is_key_false_remains_header_metadata():
    """min_fields avec is_key=False → ligne en-tête classique (nom: # label)."""
    ft = {
        "min_fields": [
            {"name": "nom",    "label": "Nom", "is_key": True},
            {"name": "statut", "label": "Statut", "is_key": False, "source": "user_input"},
        ]
    }
    lines = build_class_mxl_lines("uo", "UO-001", ft, [])
    header_lines = [l for l in lines if l.startswith("statut:")]
    assert len(header_lines) == 1
    assert "# Statut" in header_lines[0]


def test_no_min_fields_no_ident():
    """Classe sans min_fields → aucune ligne IDENT."""
    ft = {}
    lines = build_class_mxl_lines("uo", "UO-001", ft, [])
    assert not any(l.startswith("IDENT") for l in lines)


def test_multiple_is_key_fields():
    """Plusieurs is_key=True → plusieurs lignes IDENT dans l'ordre."""
    ft = {
        "min_fields": [
            {"name": "nom",  "label": "Nom",  "is_key": True},
            {"name": "site", "label": "Site", "is_key": True},
        ]
    }
    lines = build_class_mxl_lines("uo", "UO-001", ft, [])
    ident_lines = [l for l in lines if l.startswith("IDENT")]
    assert len(ident_lines) == 2
    assert ident_lines[0] == 'IDENT nom : LABEL="Nom"'
    assert ident_lines[1] == 'IDENT site : LABEL="Site"'
