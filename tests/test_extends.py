"""
Tests — instruction EXTENDS
============================
Couvre :
  1. Parsing de EXTENDS dans parse_lines()
  2. parse_mxl_file() — lecture + substitution de variables
  3. merge_asts() — règles de fusion par type d'instruction
  4. resolve_extends() — pipeline complet avec template réel
  5. Cas d'erreur (template absent, syntaxe invalide)
"""
import textwrap
from pathlib import Path

import pytest

from src.parser import (
    ExtendsNode, FileHeader, DefNode, PullNode, BindNode,
    PushNode, ValidateNode, ColNode, ParseError,
    PasserelleAST, parse_lines, parse_mxl_file,
    merge_asts, resolve_extends,
)


# ─── Helpers ──────────────────────────────────────────────────────────────────

def _make_ast(
    file_type="uo_instance", file_id="UO-001", version="1", doc="",
    extends=None, defs=None, pulls=None, cols=None,
    validates=None, binds=None, pushes=None,
) -> PasserelleAST:
    ast = PasserelleAST()
    ast.header = FileHeader(file_type=file_type, file_id=file_id,
                            version=version, doc=doc)
    ast.extends   = extends
    ast.defs      = defs      or []
    ast.pulls     = pulls     or []
    ast.cols      = cols      or []
    ast.validates = validates or []
    ast.binds     = binds     or []
    ast.pushes    = pushes    or []
    ast._defs_index = {d.var_name: d for d in ast.defs}
    return ast


def _write_mxl(tmp_path: Path, name: str, content: str) -> Path:
    """Écrit un fichier .mxl dans tmp_path et retourne son chemin."""
    p = tmp_path / f"{name}.mxl"
    p.write_text(textwrap.dedent(content), encoding="utf-8")
    return p


# ─── 1. Parsing EXTENDS ───────────────────────────────────────────────────────

class TestParseExtends:
    def test_extends_parse(self):
        lines = [
            ("FILE_TYPE: uo_instance", ""),
            ("FILE_ID:   UO-001", ""),
            ("EXTENDS uo_generique", ""),
        ]
        ast = parse_lines(lines)
        assert ast.extends is not None
        assert ast.extends.template_name == "uo_generique"
        assert ast.errors == []

    def test_sans_extends(self):
        lines = [("FILE_TYPE: uo_instance", ""), ("FILE_ID: UO-001", "")]
        ast = parse_lines(lines)
        assert ast.extends is None

    def test_extends_syntaxe_invalide(self):
        lines = [("EXTENDS", "")]   # manque le nom
        ast = parse_lines(lines)
        assert len(ast.errors) == 1
        assert "EXTENDS" in ast.errors[0].message

    def test_doc_parse(self):
        lines = [
            ("FILE_ID: UO-001", ""),
            ('DOC: "Suivi UO Signalisation"', ""),
        ]
        ast = parse_lines(lines)
        assert ast.header.doc == "Suivi UO Signalisation"

    def test_doc_sans_guillemets(self):
        lines = [("DOC: Suivi UO Signalisation", "")]
        ast = parse_lines(lines)
        assert ast.header.doc == "Suivi UO Signalisation"


# ─── 2. parse_mxl_file ────────────────────────────────────────────────────────

class TestParseMxlFile:
    def test_lit_instructions(self, tmp_path):
        _write_mxl(tmp_path, "t", """
            FILE_TYPE: uo_instance
            FILE_ID:   {{FILE_ID}}
            DEF $activites = GET_TABLE(Activites, TabActivites)
            PUSH $avancement -> uo.{{FILE_ID}}.avancement
        """)
        p = tmp_path / "t.mxl"
        ast = parse_mxl_file(p, {"FILE_ID": "UO-001"})
        assert ast.header.file_type == "uo_instance"
        assert ast.header.file_id   == "UO-001"
        assert len(ast.defs)   == 1
        assert len(ast.pushes) == 1
        assert ast.pushes[0].global_name == "uo.UO-001.avancement"

    def test_ignore_commentaires(self, tmp_path):
        _write_mxl(tmp_path, "t", """
            # Ceci est un commentaire
            FILE_TYPE: uo_instance
            # Un autre commentaire
        """)
        ast = parse_mxl_file(tmp_path / "t.mxl")
        assert ast.errors == []
        assert ast.header.file_type == "uo_instance"

    def test_substitution_file_id(self, tmp_path):
        _write_mxl(tmp_path, "t", "PUSH $x -> uo.{{FILE_ID}}.x")
        ast = parse_mxl_file(tmp_path / "t.mxl", {"FILE_ID": "UO-042"})
        assert ast.pushes[0].global_name == "uo.UO-042.x"

    def test_substitution_doc(self, tmp_path):
        _write_mxl(tmp_path, "t", 'DOC: {{DOC}}')
        ast = parse_mxl_file(tmp_path / "t.mxl", {"DOC": "Mon UO"})
        assert ast.header.doc == "Mon UO"

    def test_sans_substitution(self, tmp_path):
        _write_mxl(tmp_path, "t", "FILE_TYPE: referentiel")
        ast = parse_mxl_file(tmp_path / "t.mxl")
        assert ast.header.file_type == "referentiel"


# ─── 3. merge_asts ────────────────────────────────────────────────────────────

class TestMergeAsts:

    # Header : l'enfant prime
    def test_header_enfant_prime(self):
        child    = _make_ast(file_id="UO-001", file_type="uo_instance")
        template = _make_ast(file_id="{{FILE_ID}}", file_type="uo_instance", doc="Doc template")
        merged = merge_asts(child, template)
        assert merged.header.file_id   == "UO-001"
        assert merged.header.file_type == "uo_instance"

    def test_header_fallback_template(self):
        """Si l'enfant n'a pas de file_type, il prend celui du template."""
        child    = _make_ast(file_type="", file_id="UO-001")
        template = _make_ast(file_type="uo_instance", file_id="")
        merged = merge_asts(child, template)
        assert merged.header.file_type == "uo_instance"

    def test_doc_enfant_prime(self):
        child    = _make_ast(doc="Doc enfant")
        template = _make_ast(doc="Doc template")
        merged = merge_asts(child, template)
        assert merged.header.doc == "Doc enfant"

    def test_doc_fallback_template(self):
        child    = _make_ast(doc="")
        template = _make_ast(doc="Doc template")
        merged = merge_asts(child, template)
        assert merged.header.doc == "Doc template"

    # DEF : l'enfant remplace par même nom
    def test_def_enfant_remplace(self):
        d_template = DefNode("$avancement", "COMPUTE", formula="SUM($t.h)")
        d_child    = DefNode("$avancement", "COMPUTE", formula="AVG($t.h)")
        child    = _make_ast(defs=[d_child])
        template = _make_ast(defs=[d_template])
        merged = merge_asts(child, template)
        assert len(merged.defs) == 1
        assert merged.defs[0].formula == "AVG($t.h)"

    def test_def_nouveau_enfant_ajoute(self):
        d_template = DefNode("$total",      "COMPUTE", formula="SUM($t.h)")
        d_child    = DefNode("$par_systeme","COMPUTE", formula="GROUP_BY($t, s)")
        child    = _make_ast(defs=[d_child])
        template = _make_ast(defs=[d_template])
        merged = merge_asts(child, template)
        assert len(merged.defs) == 2
        names = [d.var_name for d in merged.defs]
        assert "$total" in names
        assert "$par_systeme" in names

    def test_def_ordre_template_puis_enfant(self):
        """Les DEF du template viennent en premier, les nouveaux DEF enfant à la fin."""
        child    = _make_ast(defs=[DefNode("$nouveau", "COMPUTE", formula="SUM($t.v)")])
        template = _make_ast(defs=[
            DefNode("$a", "COMPUTE", formula="SUM($t.a)"),
            DefNode("$b", "COMPUTE", formula="SUM($t.b)"),
        ])
        merged = merge_asts(child, template)
        assert [d.var_name for d in merged.defs] == ["$a", "$b", "$nouveau"]

    # Listes additives
    def test_pull_additif(self):
        child    = _make_ast(pulls=[PullNode("ref.specifique", "FILL_TABLE", "S", "T", "APPEND_NEW")])
        template = _make_ast(pulls=[PullNode("projet.acteurs",  "FILL_TABLE", "O", "T", "OVERWRITE")])
        merged = merge_asts(child, template)
        assert len(merged.pulls) == 2

    def test_push_additif(self):
        child    = _make_ast(pushes=[PushNode("$x", "uo.UO-001.x")])
        template = _make_ast(pushes=[PushNode("$avancement", "uo.UO-001.avancement")])
        merged = merge_asts(child, template)
        assert len(merged.pushes) == 2

    def test_bind_additif(self):
        child    = _make_ast(binds=[BindNode("$x", "Dashboard", "x")])
        template = _make_ast(binds=[BindNode("$avancement", "Dashboard", "avancement_global")])
        merged = merge_asts(child, template)
        assert len(merged.binds) == 2

    def test_validate_additif(self):
        child    = _make_ast(validates=[ValidateNode("$t.statut", "UNIQUE")])
        template = _make_ast(validates=[ValidateNode("$t.avancement", "RANGE(0, 100)")])
        merged = merge_asts(child, template)
        assert len(merged.validates) == 2

    def test_col_additif(self):
        child    = _make_ast(cols=[ColNode("$t.specifique", "$t", "specifique")])
        template = _make_ast(cols=[ColNode("$t.id", "$t", "id", is_key=True)])
        merged = merge_asts(child, template)
        assert len(merged.cols) == 2

    def test_erreurs_combinees(self):
        child    = _make_ast()
        child.errors    = [ParseError(1, "x", "erreur enfant")]
        template = _make_ast()
        template.errors = [ParseError(2, "y", "erreur template")]
        merged = merge_asts(child, template)
        assert len(merged.errors) == 2


# ─── 4. resolve_extends — pipeline complet ────────────────────────────────────

class TestResolveExtends:
    def test_sans_extends_retourne_ast_inchange(self):
        ast = _make_ast(file_id="UO-001")
        result = resolve_extends(ast)
        assert result is ast   # même objet

    def test_template_absent_ajoute_erreur(self, tmp_path):
        ast = _make_ast(file_id="UO-001")
        ast.extends = ExtendsNode("template_inexistant")
        result = resolve_extends(ast, templates_dir=tmp_path)
        assert len(result.errors) == 1
        assert "introuvable" in result.errors[0].message.lower()

    def test_fusion_avec_vrai_template(self, tmp_path):
        """Teste resolve_extends avec un vrai fichier .mxl."""
        _write_mxl(tmp_path, "uo_generique", """
            FILE_TYPE: uo_instance
            FILE_ID:   {{FILE_ID}}
            DEF $activites = GET_TABLE(Activites, TabActivites)
            DEF $avancement = COMPUTE(MEAN_WEIGHTED($activites.avancement, $activites.heures))
            PUSH $avancement -> uo.{{FILE_ID}}.avancement
            VALIDATE $activites.avancement : RANGE(0, 100)
        """)

        child = _make_ast(
            file_id="UO-042",
            doc="Mon UO",
            extends=ExtendsNode("uo_generique"),
        )
        merged = resolve_extends(child, templates_dir=tmp_path)

        assert merged.errors == []
        assert merged.header.file_id   == "UO-042"
        assert merged.header.file_type == "uo_instance"
        assert len(merged.defs)        == 2
        assert len(merged.pushes)      == 1
        # Substitution : {{FILE_ID}} → UO-042
        assert merged.pushes[0].global_name == "uo.UO-042.avancement"
        assert len(merged.validates) == 1

    def test_enfant_override_def_template(self, tmp_path):
        """Un DEF enfant remplace la version du template."""
        _write_mxl(tmp_path, "t", """
            DEF $avancement = COMPUTE(SUM($t.h))
            PUSH $avancement -> store.avancement
        """)
        child = _make_ast(
            file_id="UO-001",
            extends=ExtendsNode("t"),
            defs=[DefNode("$avancement", "COMPUTE", formula="AVG($t.h)")],
        )
        merged = resolve_extends(child, templates_dir=tmp_path)
        assert merged.errors == []
        # Le DEF enfant (AVG) remplace celui du template (SUM)
        assert merged.defs[0].formula == "AVG($t.h)"
        assert len(merged.defs) == 1

    def test_enfant_pull_additionne(self, tmp_path):
        """Un PULL enfant s'ajoute au PULL du template."""
        _write_mxl(tmp_path, "t", "PULL projet.acteurs -> FILL_TABLE(Org, TabA) MODE=OVERWRITE")
        child = _make_ast(
            file_id="UO-001",
            extends=ExtendsNode("t"),
            pulls=[PullNode("ref.specifique", "FILL_TABLE", "S", "TabS", "APPEND_NEW")],
        )
        merged = resolve_extends(child, templates_dir=tmp_path)
        assert len(merged.pulls) == 2

    def test_avec_template_uo_generique_reel(self):
        """Teste avec le vrai uo_generique.mxl du projet."""
        real_templates = Path(__file__).parent.parent / "config" / "templates"
        if not (real_templates / "uo_generique.mxl").exists():
            pytest.skip("uo_generique.mxl absent")

        child = _make_ast(file_id="UO-TEST", doc="Test UO", extends=ExtendsNode("uo_generique"))
        merged = resolve_extends(child, templates_dir=real_templates)

        assert merged.header.file_id == "UO-TEST"
        assert len(merged.errors) == 0
        # Le template doit apporter des DEF, PUSH et BIND
        assert len(merged.defs)   > 0
        assert len(merged.pushes) > 0
        # Vérifier la substitution {{FILE_ID}}
        push_keys = [p.global_name for p in merged.pushes]
        assert any("UO-TEST" in k for k in push_keys)
