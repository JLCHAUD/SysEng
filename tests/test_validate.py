"""
Tests — instruction VALIDATE
==============================
Couvre :
  1. Parsing de _parse_validate() et parse_lines()
  2. _validate_rule() — toutes les règles
  3. execute_validates() — intégration avec le contexte
  4. Sévérité error vs warning
  5. Scalaire vs colonne de table
"""
import pytest
from src.parser import _parse_validate, ValidateNode, parse_lines
from src.executor import _validate_rule, execute_validates, ExecutionResult
from src.parser import PasserelleAST


# ─── 1. Parsing ───────────────────────────────────────────────────────────────

class TestParseValidate:
    def test_not_null(self):
        node = _parse_validate("VALIDATE $activites.id : NOT_NULL")
        assert node is not None
        assert node.var_ref  == "$activites.id"
        assert node.rule     == "NOT_NULL"
        assert node.severity == "error"

    def test_range(self):
        node = _parse_validate("VALIDATE $activites.avancement : RANGE(0, 100)")
        assert node.rule     == "RANGE(0, 100)"
        assert node.severity == "error"

    def test_in(self):
        node = _parse_validate('VALIDATE $activites.statut : IN("EN_COURS", "CLOTUREE")')
        assert node.rule.startswith("IN(")
        assert node.severity == "error"

    def test_severity_warning(self):
        node = _parse_validate("VALIDATE $total_heures : NON_NEGATIVE  SEVERITY=warning")
        assert node.var_ref  == "$total_heures"
        assert node.rule     == "NON_NEGATIVE"
        assert node.severity == "warning"

    def test_severity_casse_insensible(self):
        node = _parse_validate("VALIDATE $x.v : POSITIVE  SEVERITY=WARNING")
        assert node.severity == "warning"

    def test_unique(self):
        node = _parse_validate("VALIDATE $activites.id : UNIQUE")
        assert node.rule == "UNIQUE"

    def test_scalaire(self):
        node = _parse_validate("VALIDATE $total_heures : NON_NEGATIVE")
        assert node.var_ref == "$total_heures"
        assert "." not in node.var_ref

    def test_syntaxe_invalide(self):
        assert _parse_validate("VALIDATE mal forme") is None
        assert _parse_validate("VALIDATE $x") is None

    def test_parse_lines_integre_validate(self):
        lines = [
            ("FILE_TYPE: uo_instance", ""),
            ("FILE_ID: UO-001", ""),
            ("VALIDATE $activites.avancement : RANGE(0, 100)", ""),
            ("VALIDATE $activites.id         : NOT_NULL", ""),
            ("VALIDATE $total_heures         : NON_NEGATIVE  SEVERITY=warning", ""),
        ]
        ast = parse_lines(lines)
        assert len(ast.validates) == 3
        assert ast.errors == []
        assert ast.validates[0].rule == "RANGE(0, 100)"
        assert ast.validates[2].severity == "warning"


# ─── 2. Règles — _validate_rule ───────────────────────────────────────────────

class TestValidateRule:

    # NOT_NULL
    def test_not_null_ok(self):
        assert _validate_rule("NOT_NULL", [1, "a", 0]) == []

    def test_not_null_violation(self):
        v = _validate_rule("NOT_NULL", [1, None, 3])
        assert len(v) == 1
        assert "NOT_NULL" in v[0]

    # POSITIVE
    def test_positive_ok(self):
        assert _validate_rule("POSITIVE", [1, 2, 100]) == []

    def test_positive_zero_violation(self):
        assert _validate_rule("POSITIVE", [0]) != []

    def test_positive_negatif_violation(self):
        assert _validate_rule("POSITIVE", [-5]) != []

    # NON_NEGATIVE
    def test_non_negative_zero_ok(self):
        assert _validate_rule("NON_NEGATIVE", [0, 1, 100]) == []

    def test_non_negative_negatif_violation(self):
        v = _validate_rule("NON_NEGATIVE", [-1])
        assert len(v) == 1

    # UNIQUE
    def test_unique_ok(self):
        assert _validate_rule("UNIQUE", ["A1", "A2", "A3"]) == []

    def test_unique_doublon(self):
        v = _validate_rule("UNIQUE", ["A1", "A2", "A1"])
        assert len(v) == 1
        assert "UNIQUE" in v[0]

    def test_unique_ignore_none(self):
        # Les None ne comptent pas comme doublons
        assert _validate_rule("UNIQUE", [None, None, "A1"]) == []

    # RANGE
    def test_range_ok(self):
        assert _validate_rule("RANGE(0, 100)", [0, 50, 100]) == []

    def test_range_hors_borne_haute(self):
        v = _validate_rule("RANGE(0, 100)", [101])
        assert len(v) == 1

    def test_range_hors_borne_basse(self):
        v = _validate_rule("RANGE(0, 100)", [-1])
        assert len(v) == 1

    def test_range_none_est_violation(self):
        v = _validate_rule("RANGE(0, 100)", [None])
        assert len(v) == 1

    # IN
    def test_in_ok(self):
        assert _validate_rule('IN("EN_COURS", "CLOTUREE")', ["EN_COURS", "CLOTUREE"]) == []

    def test_in_valeur_absente(self):
        v = _validate_rule('IN("EN_COURS", "CLOTUREE")', ["INCONNU"])
        assert len(v) == 1

    def test_in_plusieurs_violations(self):
        v = _validate_rule('IN("A", "B")', ["A", "C", "D"])
        assert len(v) == 1
        assert "2" in v[0]  # 2 violations mentionnées

    # Règle inconnue
    def test_regle_inconnue(self):
        v = _validate_rule("INEXISTANT", [1])
        assert len(v) == 1
        assert "inconnue" in v[0].lower()


# ─── 3. execute_validates — intégration ───────────────────────────────────────

class TestExecuteValidates:

    def _ast(self, validates):
        ast = PasserelleAST()
        ast.validates = validates
        return ast

    # Colonne de table
    def test_colonne_ok(self):
        ast = self._ast([ValidateNode("$activites.avancement", "RANGE(0, 100)")])
        ctx = {"$activites": [
            {"avancement": 60}, {"avancement": 80}, {"avancement": 100}
        ]}
        result = ExecutionResult()
        execute_validates(ast, ctx, result)
        assert result.errors   == []
        assert result.warnings == []

    def test_colonne_violation_error(self):
        ast = self._ast([ValidateNode("$activites.avancement", "RANGE(0, 100)")])
        ctx = {"$activites": [{"avancement": 150}]}
        result = ExecutionResult()
        execute_validates(ast, ctx, result)
        assert len(result.errors) == 1
        assert "$activites.avancement" in result.errors[0]

    # Scalaire
    def test_scalaire_ok(self):
        ast = self._ast([ValidateNode("$total_heures", "NON_NEGATIVE")])
        ctx = {"$total_heures": 430}
        result = ExecutionResult()
        execute_validates(ast, ctx, result)
        assert result.errors == []

    def test_scalaire_violation(self):
        ast = self._ast([ValidateNode("$total_heures", "NON_NEGATIVE")])
        ctx = {"$total_heures": -10}
        result = ExecutionResult()
        execute_validates(ast, ctx, result)
        assert len(result.errors) == 1

    # Sévérité warning
    def test_severity_warning_va_dans_warnings(self):
        ast = self._ast([ValidateNode("$total_heures", "NON_NEGATIVE", severity="warning")])
        ctx = {"$total_heures": -5}
        result = ExecutionResult()
        execute_validates(ast, ctx, result)
        assert result.errors   == []        # pas dans errors
        assert len(result.warnings) == 1    # dans warnings

    def test_severity_error_va_dans_errors(self):
        ast = self._ast([ValidateNode("$activites.id", "NOT_NULL", severity="error")])
        ctx = {"$activites": [{"id": None}]}
        result = ExecutionResult()
        execute_validates(ast, ctx, result)
        assert len(result.errors) == 1
        assert result.warnings == []

    # Variable absente
    def test_variable_indefinie_erreur(self):
        ast = self._ast([ValidateNode("$inconnu.col", "NOT_NULL")])
        ctx = {}
        result = ExecutionResult()
        execute_validates(ast, ctx, result)
        assert len(result.errors) == 1
        assert "non définie" in result.errors[0]

    # Plusieurs règles
    def test_plusieurs_regles_independantes(self):
        ast = self._ast([
            ValidateNode("$activites.avancement", "RANGE(0, 100)"),
            ValidateNode("$activites.id",         "NOT_NULL"),
            ValidateNode("$activites.statut",     'IN("EN_COURS", "CLOTUREE")'),
        ])
        ctx = {"$activites": [
            {"avancement": 60, "id": "A1", "statut": "EN_COURS"},
            {"avancement": 80, "id": "A2", "statut": "CLOTUREE"},
        ]}
        result = ExecutionResult()
        execute_validates(ast, ctx, result)
        assert result.errors == []

    def test_plusieurs_violations_independantes(self):
        ast = self._ast([
            ValidateNode("$activites.avancement", "RANGE(0, 100)"),
            ValidateNode("$activites.id",         "NOT_NULL"),
        ])
        ctx = {"$activites": [
            {"avancement": 150, "id": None},  # deux violations
        ]}
        result = ExecutionResult()
        execute_validates(ast, ctx, result)
        assert len(result.errors) == 2  # une par règle

    # has_blocking_errors
    def test_has_blocking_errors_true(self):
        result = ExecutionResult(errors=["oops"])
        assert result.has_blocking_errors is True

    def test_has_blocking_errors_false(self):
        result = ExecutionResult(warnings=["attention"])
        assert result.has_blocking_errors is False

    def test_warnings_only_ne_bloquent_pas(self):
        ast = self._ast([
            ValidateNode("$total_heures", "NON_NEGATIVE", severity="warning"),
        ])
        ctx = {"$total_heures": -1}
        result = ExecutionResult()
        execute_validates(ast, ctx, result)
        assert not result.has_blocking_errors
        assert len(result.warnings) == 1


# ─── 4. Scénario complet — UO typique ─────────────────────────────────────────

class TestScenarioUO:
    def test_uo_valide(self):
        """Données propres → aucune violation."""
        ast = PasserelleAST()
        ast.validates = [
            ValidateNode("$activites.avancement", "RANGE(0, 100)"),
            ValidateNode("$activites.id",          "NOT_NULL"),
            ValidateNode("$activites.id",          "UNIQUE"),
            ValidateNode("$activites.statut",      'IN("EN_COURS", "CLOTUREE", "SUSPENDUE")'),
            ValidateNode("$total_heures",          "NON_NEGATIVE", severity="warning"),
        ]
        ctx = {
            "$activites": [
                {"id": "A1", "avancement": 60,  "statut": "EN_COURS"},
                {"id": "A2", "avancement": 100, "statut": "CLOTUREE"},
                {"id": "A3", "avancement": 0,   "statut": "SUSPENDUE"},
            ],
            "$total_heures": 430,
        }
        result = ExecutionResult()
        execute_validates(ast, ctx, result)
        assert result.errors   == []
        assert result.warnings == []

    def test_uo_avec_donnees_corrompues(self):
        """Avancement hors range + doublon d'ID → 2 erreurs."""
        ast = PasserelleAST()
        ast.validates = [
            ValidateNode("$activites.avancement", "RANGE(0, 100)"),
            ValidateNode("$activites.id",          "UNIQUE"),
        ]
        ctx = {"$activites": [
            {"id": "A1", "avancement": 110},   # hors range
            {"id": "A1", "avancement": 50},    # doublon
        ]}
        result = ExecutionResult()
        execute_validates(ast, ctx, result)
        assert len(result.errors) == 2
