"""
Tests — clause ONLY_IF sur PUSH
================================
Vérifie le parsing de ONLY_IF et son évaluation dans execute_pushes().
"""
import pytest
from src.parser import _parse_push, PushNode, parse_lines
from src.executor import execute_pushes, ExecutionResult


# ─── Helpers ──────────────────────────────────────────────────────────────────

class MemStore:
    def __init__(self):
        self._data = {}
    def get(self, k):
        return self._data.get(k)
    def set(self, k, v):
        self._data[k] = v
    def set_many(self, d):
        self._data.update(d)


# ─── Tests parser — _parse_push ───────────────────────────────────────────────

class TestParsePush:
    def test_push_sans_only_if(self):
        node = _parse_push("PUSH $avancement -> uo.UO-001.avancement")
        assert node is not None
        assert node.var_name   == "$avancement"
        assert node.global_name == "uo.UO-001.avancement"
        assert node.only_if    == ""

    def test_push_only_if_numerique(self):
        node = _parse_push("PUSH $avancement -> uo.UO-001.avancement  ONLY_IF $total_heures > 0")
        assert node is not None
        assert node.only_if == "$total_heures > 0"

    def test_push_only_if_egal_string(self):
        node = _parse_push('PUSH $statut -> uo.UO-001.statut  ONLY_IF $statut = "VERT"')
        assert node is not None
        assert node.only_if == '$statut = "VERT"'

    def test_push_only_if_different_null(self):
        node = _parse_push("PUSH $val -> store.val  ONLY_IF $val != NULL")
        assert node is not None
        assert node.only_if == "$val != NULL"

    def test_push_only_if_superieur_egal(self):
        node = _parse_push("PUSH $taux -> store.taux  ONLY_IF $taux >= 0.5")
        assert node is not None
        assert node.only_if == "$taux >= 0.5"

    def test_push_only_if_casse_insensible(self):
        node = _parse_push("PUSH $x -> store.x  only_if $x > 0")
        assert node is not None
        assert node.only_if == "$x > 0"

    def test_push_syntaxe_invalide(self):
        node = _parse_push("PUSH mal formé")
        assert node is None


# ─── Tests parser — parse_lines intégré ───────────────────────────────────────

class TestParseLines:
    def test_parse_lines_avec_only_if(self):
        lines = [
            ("FILE_TYPE: uo_instance", ""),
            ("FILE_ID: UO-001", ""),
            ("PUSH $avancement -> uo.UO-001.avancement  ONLY_IF $total_heures > 0", ""),
            ("PUSH $statut -> uo.UO-001.statut", ""),
        ]
        ast = parse_lines(lines)
        assert len(ast.pushes) == 2
        assert ast.pushes[0].only_if == "$total_heures > 0"
        assert ast.pushes[1].only_if == ""
        assert ast.errors == []


# ─── Tests executor — execute_pushes ──────────────────────────────────────────

class TestExecutePushesOnlyIf:

    def _ast_with_pushes(self, pushes):
        """Crée un AST minimal avec une liste de PushNode."""
        from src.parser import PasserelleAST
        ast = PasserelleAST()
        ast.pushes = pushes
        return ast

    def test_sans_only_if_pousse_toujours(self):
        ast = self._ast_with_pushes([
            PushNode("$val", "store.val"),
        ])
        ctx = {"$val": 42}
        store = MemStore()
        result = ExecutionResult()
        execute_pushes(ast, ctx, store, result)
        assert store.get("store.val") == 42
        assert "store.val" in result.pushed

    def test_only_if_vrai_pousse(self):
        ast = self._ast_with_pushes([
            PushNode("$avancement", "uo.UO-001.avancement", only_if="$total_heures > 0"),
        ])
        ctx = {"$avancement": 65.0, "$total_heures": 430}
        store = MemStore()
        result = ExecutionResult()
        execute_pushes(ast, ctx, store, result)
        assert store.get("uo.UO-001.avancement") == 65.0
        assert result.skipped == []

    def test_only_if_faux_ne_pousse_pas(self):
        ast = self._ast_with_pushes([
            PushNode("$avancement", "uo.UO-001.avancement", only_if="$total_heures > 0"),
        ])
        ctx = {"$avancement": 0.0, "$total_heures": 0}  # ← condition fausse
        store = MemStore()
        result = ExecutionResult()
        execute_pushes(ast, ctx, store, result)
        assert store.get("uo.UO-001.avancement") is None  # pas écrit
        assert len(result.skipped) == 1
        assert "ONLY_IF" in result.skipped[0]

    def test_only_if_ne_pas_ecraser_valeur_existante(self):
        """Quand ONLY_IF est faux, la valeur précédente du store est préservée."""
        ast = self._ast_with_pushes([
            PushNode("$avancement", "uo.UO-001.avancement", only_if="$total_heures > 0"),
        ])
        store = MemStore()
        store.set("uo.UO-001.avancement", 72.0)  # valeur précédente

        ctx = {"$avancement": 0.0, "$total_heures": 0}
        result = ExecutionResult()
        execute_pushes(ast, ctx, store, result)
        assert store.get("uo.UO-001.avancement") == 72.0  # toujours l'ancienne valeur

    def test_only_if_different_null_pousse_si_non_null(self):
        ast = self._ast_with_pushes([
            PushNode("$statut", "store.statut", only_if="$statut != NULL"),
        ])
        ctx = {"$statut": "VERT"}
        store = MemStore()
        result = ExecutionResult()
        execute_pushes(ast, ctx, store, result)
        assert store.get("store.statut") == "VERT"

    def test_only_if_different_null_skippe_si_null(self):
        ast = self._ast_with_pushes([
            PushNode("$statut", "store.statut", only_if="$statut != NULL"),
        ])
        ctx = {"$statut": None}
        store = MemStore()
        result = ExecutionResult()
        execute_pushes(ast, ctx, store, result)
        assert store.get("store.statut") is None
        assert len(result.skipped) == 1

    def test_only_if_egalite_string(self):
        ast = self._ast_with_pushes([
            PushNode("$statut", "store.statut", only_if='$statut = "VERT"'),
        ])
        ctx = {"$statut": "ROUGE"}   # ← condition fausse
        store = MemStore()
        result = ExecutionResult()
        execute_pushes(ast, ctx, store, result)
        assert store.get("store.statut") is None

    def test_mix_with_et_sans_only_if(self):
        """PUSH avec et sans ONLY_IF dans le même AST."""
        ast = self._ast_with_pushes([
            PushNode("$avancement", "uo.UO-001.avancement", only_if="$total_heures > 0"),
            PushNode("$nb_total",   "uo.UO-001.nb_activites"),   # toujours poussé
            PushNode("$statut",     "uo.UO-001.statut",  only_if="$total_heures > 0"),
        ])
        ctx = {
            "$avancement":  0.0,
            "$nb_total":    4,
            "$statut":      "ROUGE",
            "$total_heures": 0,    # ← condition fausse pour avancement et statut
        }
        store = MemStore()
        result = ExecutionResult()
        execute_pushes(ast, ctx, store, result)

        assert store.get("uo.UO-001.avancement")  is None   # skippé
        assert store.get("uo.UO-001.nb_activites") == 4      # poussé
        assert store.get("uo.UO-001.statut")       is None   # skippé
        assert len(result.pushed)  == 1
        assert len(result.skipped) == 2

    def test_only_if_superieur_egal(self):
        ast = self._ast_with_pushes([
            PushNode("$taux", "store.taux", only_if="$taux >= 0.5"),
        ])
        store = MemStore()
        result = ExecutionResult()

        # Cas vrai
        execute_pushes(ast, {"$taux": 0.5}, store, result)
        assert store.get("store.taux") == 0.5

        # Cas faux
        store2 = MemStore()
        result2 = ExecutionResult()
        execute_pushes(ast, {"$taux": 0.3}, store2, result2)
        assert store2.get("store.taux") is None
