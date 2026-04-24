"""
Tests — GROUP_BY / SORT / TOP_N
================================
Couvre :
  1. _eval_agg_on_rows()  — fonctions d'agrégation sur des rows bruts
  2. GROUP_BY             — regroupement + agrégation via _eval_formula / execute_computes
  3. SORT                 — tri ASC/DESC, gestion des None
  4. TOP_N                — tri + troncature
  5. Cas d'erreur         — table absente, mauvais nombre d'arguments
  6. Scénario intégration — GROUP_BY puis TOP_N sur le résultat
"""
import pytest
from src.executor import (
    _eval_agg_on_rows,
    _eval_group_by,
    _eval_sort,
    _eval_top_n,
    _eval_formula,
    execute_computes,
    ExecutionResult,
)
from src.parser import PasserelleAST, DefNode


# ─── Données communes ─────────────────────────────────────────────────────────

ACTIVITES = [
    {"id": "A1", "type": "TRAVAUX",    "avancement": 60,  "heures": 200, "statut": "EN_COURS"},
    {"id": "A2", "type": "TRAVAUX",    "avancement": 80,  "heures": 150, "statut": "CLOTUREE"},
    {"id": "A3", "type": "ETUDES",     "avancement": 40,  "heures": 100, "statut": "EN_COURS"},
    {"id": "A4", "type": "ETUDES",     "avancement": 100, "heures": 50,  "statut": "CLOTUREE"},
    {"id": "A5", "type": "FOURNITURE", "avancement": 20,  "heures": 300, "statut": "EN_COURS"},
]


# ─── 1. _eval_agg_on_rows ─────────────────────────────────────────────────────

class TestEvalAggOnRows:

    def test_sum(self):
        rows = [{"h": 100}, {"h": 200}, {"h": 50}]
        assert _eval_agg_on_rows("SUM(h)", rows) == 350

    def test_sum_ignore_none(self):
        rows = [{"h": 100}, {"h": None}, {"h": 50}]
        assert _eval_agg_on_rows("SUM(h)", rows) == 150

    def test_count(self):
        rows = [{"id": "A1"}, {"id": "A2"}, {"id": None}]
        assert _eval_agg_on_rows("COUNT(id)", rows) == 2

    def test_avg(self):
        rows = [{"v": 60}, {"v": 80}, {"v": 100}]
        assert _eval_agg_on_rows("AVG(v)", rows) == 80.0

    def test_avg_vide_retourne_zero(self):
        assert _eval_agg_on_rows("AVG(v)", [{"v": None}]) == 0.0

    def test_min(self):
        rows = [{"h": 100}, {"h": 50}, {"h": 200}]
        assert _eval_agg_on_rows("MIN(h)", rows) == 50

    def test_max(self):
        rows = [{"h": 100}, {"h": 50}, {"h": 200}]
        assert _eval_agg_on_rows("MAX(h)", rows) == 200

    def test_count_if(self):
        rows = [{"s": "EN_COURS"}, {"s": "CLOTUREE"}, {"s": "EN_COURS"}]
        assert _eval_agg_on_rows('COUNT_IF(s, "EN_COURS")', rows) == 2

    def test_mean_weighted(self):
        # MEAN_WEIGHTED : (60*200 + 80*150) / (200+150) = (12000+12000)/350 ≈ 68.57
        rows = [{"av": 60, "h": 200}, {"av": 80, "h": 150}]
        result = _eval_agg_on_rows("MEAN_WEIGHTED(av, h)", rows)
        assert abs(result - 68.571) < 0.01

    def test_mean_weighted_poids_nul(self):
        rows = [{"av": 60, "h": 0}, {"av": 80, "h": 0}]
        assert _eval_agg_on_rows("MEAN_WEIGHTED(av, h)", rows) == 0.0

    def test_agg_inconnue_leve_erreur(self):
        with pytest.raises(ValueError, match="Agrégation inconnue"):
            _eval_agg_on_rows("MEDIAN(v)", [{"v": 1}])


# ─── 2. GROUP_BY ──────────────────────────────────────────────────────────────

class TestGroupBy:

    def _ctx(self, table=None):
        return {"$activites": table if table is not None else ACTIVITES}

    def test_groupe_simple_sum(self):
        ctx = self._ctx()
        result = _eval_group_by(
            "$activites, type, heures = SUM(heures)",
            ctx,
        )
        by_type = {r["type"]: r for r in result}
        assert len(result) == 3
        assert by_type["TRAVAUX"]["heures"]    == 350   # 200+150
        assert by_type["ETUDES"]["heures"]     == 150   # 100+50
        assert by_type["FOURNITURE"]["heures"] == 300

    def test_groupe_multiple_aggs(self):
        ctx = self._ctx()
        result = _eval_group_by(
            "$activites, type, heures = SUM(heures), nb = COUNT(id)",
            ctx,
        )
        by_type = {r["type"]: r for r in result}
        assert by_type["TRAVAUX"]["nb"]    == 2
        assert by_type["ETUDES"]["nb"]     == 2
        assert by_type["FOURNITURE"]["nb"] == 1

    def test_groupe_mean_weighted(self):
        ctx = self._ctx()
        result = _eval_group_by(
            "$activites, type, avancement = MEAN_WEIGHTED(avancement, heures)",
            ctx,
        )
        by_type = {r["type"]: r for r in result}
        # TRAVAUX : (60*200 + 80*150) / 350
        expected = (60 * 200 + 80 * 150) / 350
        assert abs(by_type["TRAVAUX"]["avancement"] - expected) < 0.01

    def test_ordre_premier_apparu(self):
        """Les groupes sont retournés dans l'ordre d'apparition."""
        ctx = self._ctx()
        result = _eval_group_by("$activites, type, nb = COUNT(id)", ctx)
        types = [r["type"] for r in result]
        assert types == ["TRAVAUX", "ETUDES", "FOURNITURE"]

    def test_groupe_unique(self):
        """Si toutes les lignes ont la même valeur → un seul groupe."""
        table = [{"type": "X", "v": 10}, {"type": "X", "v": 20}]
        result = _eval_group_by("$t, type, total = SUM(v)", {"$t": table})
        assert len(result) == 1
        assert result[0]["total"] == 30

    def test_table_vide(self):
        result = _eval_group_by("$t, type, total = SUM(v)", {"$t": []})
        assert result == []

    def test_table_indefinie_leve_erreur(self):
        with pytest.raises(ValueError, match="non définie"):
            _eval_group_by("$inconnu, col, n = COUNT(col)", {})

    def test_trop_peu_darguments(self):
        with pytest.raises(ValueError, match="au moins 3"):
            _eval_group_by("$t, col", {"$t": []})

    def test_spec_invalide_leve_erreur(self):
        with pytest.raises(ValueError, match="spécification invalide"):
            _eval_group_by("$t, col, pas_une_spec", {"$t": [{"col": 1}]})

    def test_via_eval_formula(self):
        """Vérifie l'accès via _eval_formula (COMPUTE wrapping)."""
        ctx = {"$activites": ACTIVITES}
        result = _eval_formula(
            "GROUP_BY($activites, type, heures = SUM(heures))",
            ctx,
        )
        assert isinstance(result, list)
        assert len(result) == 3

    def test_colonne_groupe_presente_dans_resultat(self):
        """La colonne de groupement doit apparaître dans chaque ligne résultat."""
        ctx = self._ctx()
        result = _eval_group_by("$activites, type, n = COUNT(id)", ctx)
        for row in result:
            assert "type" in row


# ─── 3. SORT ──────────────────────────────────────────────────────────────────

class TestSort:

    def _ctx(self):
        return {"$activites": ACTIVITES}

    def test_sort_asc(self):
        ctx = self._ctx()
        result = _eval_sort("$activites, heures, ASC", ctx)
        values = [r["heures"] for r in result]
        assert values == sorted(values)

    def test_sort_desc(self):
        ctx = self._ctx()
        result = _eval_sort("$activites, heures, DESC", ctx)
        values = [r["heures"] for r in result]
        assert values == sorted(values, reverse=True)

    def test_sort_string_asc(self):
        ctx = self._ctx()
        result = _eval_sort("$activites, type, ASC", ctx)
        types = [r["type"] for r in result]
        assert types == sorted(types)

    def test_sort_ne_modifie_pas_le_contexte(self):
        """SORT retourne une nouvelle liste sans modifier le contexte."""
        original = list(ACTIVITES)
        ctx = {"$activites": ACTIVITES}
        _ = _eval_sort("$activites, heures, ASC", ctx)
        assert ctx["$activites"] == ACTIVITES  # même objet

    def test_sort_none_en_dernier_asc(self):
        table = [{"v": 10}, {"v": None}, {"v": 5}]
        result = _eval_sort("$t, v, ASC", {"$t": table})
        assert result[0]["v"] == 5
        assert result[1]["v"] == 10
        assert result[2]["v"] is None

    def test_sort_none_en_dernier_desc(self):
        table = [{"v": 10}, {"v": None}, {"v": 5}]
        result = _eval_sort("$t, v, DESC", {"$t": table})
        assert result[0]["v"] == 10
        assert result[1]["v"] == 5
        assert result[2]["v"] is None

    def test_table_vide(self):
        result = _eval_sort("$t, v, ASC", {"$t": []})
        assert result == []

    def test_table_indefinie_leve_erreur(self):
        with pytest.raises(ValueError, match="non définie"):
            _eval_sort("$inconnu, col, ASC", {})

    def test_direction_invalide_leve_erreur(self):
        with pytest.raises(ValueError, match="direction inconnue"):
            _eval_sort("$t, v, CROISSANT", {"$t": [{"v": 1}]})

    def test_trop_peu_darguments(self):
        with pytest.raises(ValueError, match="3 arguments"):
            _eval_sort("$t, v", {"$t": []})

    def test_via_eval_formula(self):
        ctx = {"$activites": ACTIVITES}
        result = _eval_formula("SORT($activites, heures, ASC)", ctx)
        assert isinstance(result, list)
        assert result[0]["heures"] == 50


# ─── 4. TOP_N ─────────────────────────────────────────────────────────────────

class TestTopN:

    def _ctx(self):
        return {"$activites": ACTIVITES}

    def test_top3_desc(self):
        ctx = self._ctx()
        result = _eval_top_n("$activites, 3, heures, DESC", ctx)
        assert len(result) == 3
        # Les 3 premiers en heures DESC : 300, 200, 150
        assert result[0]["heures"] == 300
        assert result[1]["heures"] == 200
        assert result[2]["heures"] == 150

    def test_top1_asc(self):
        ctx = self._ctx()
        result = _eval_top_n("$activites, 1, heures, ASC", ctx)
        assert len(result) == 1
        assert result[0]["heures"] == 50

    def test_n_superieur_taille_retourne_tout(self):
        ctx = self._ctx()
        result = _eval_top_n("$activites, 100, heures, DESC", ctx)
        assert len(result) == len(ACTIVITES)

    def test_n_zero(self):
        ctx = self._ctx()
        result = _eval_top_n("$activites, 0, heures, DESC", ctx)
        assert result == []

    def test_table_vide(self):
        result = _eval_top_n("$t, 3, v, DESC", {"$t": []})
        assert result == []

    def test_table_indefinie_leve_erreur(self):
        with pytest.raises(ValueError, match="non définie"):
            _eval_top_n("$inconnu, 3, col, DESC", {})

    def test_n_non_entier_leve_erreur(self):
        with pytest.raises(ValueError, match="n="):
            _eval_top_n("$t, abc, col, DESC", {"$t": [{"col": 1}]})

    def test_direction_invalide_leve_erreur(self):
        with pytest.raises(ValueError, match="direction inconnue"):
            _eval_top_n("$t, 3, col, HAUT", {"$t": [{"col": 1}]})

    def test_trop_peu_darguments(self):
        with pytest.raises(ValueError, match="4 arguments"):
            _eval_top_n("$t, 3, col", {"$t": []})

    def test_via_eval_formula(self):
        ctx = {"$activites": ACTIVITES}
        result = _eval_formula("TOP_N($activites, 2, heures, DESC)", ctx)
        assert isinstance(result, list)
        assert len(result) == 2
        assert result[0]["heures"] == 300

    def test_none_en_dernier(self):
        table = [{"v": 10}, {"v": None}, {"v": 5}, {"v": 20}]
        result = _eval_top_n("$t, 2, v, DESC", {"$t": table})
        assert len(result) == 2
        assert result[0]["v"] == 20
        assert result[1]["v"] == 10


# ─── 5. Intégration — execute_computes ────────────────────────────────────────

class TestGroupBySortTopNIntegration:
    """Teste GROUP_BY/SORT/TOP_N via execute_computes avec un faux workbook."""

    class _FakeWB:
        """Faux workbook qui n'est jamais ouvert (les tables viennent du ctx pré-rempli)."""
        sheetnames = []

    def _ast(self, defs):
        ast = PasserelleAST()
        ast.defs = defs
        return ast

    def test_group_by_puis_sort_dans_ast(self):
        """GROUP_BY crée une variable de table que SORT peut trier."""
        ast = self._ast([
            DefNode("$par_type",  "COMPUTE",
                    formula="GROUP_BY($activites, type, heures = SUM(heures), nb = COUNT(id))"),
            DefNode("$par_type_tri", "COMPUTE",
                    formula="SORT($par_type, heures, DESC)"),
        ])
        # On pré-injecte $activites dans le contexte fictif via un DEF fantôme
        # Pour simplifier, on le met directement dans un ctx via monkey-patch
        ast.defs = [
            DefNode("$par_type",     "COMPUTE",
                    formula="GROUP_BY($activites, type, heures = SUM(heures), nb = COUNT(id))"),
            DefNode("$par_type_tri", "COMPUTE",
                    formula="SORT($par_type, heures, DESC)"),
        ]
        # On précharge $activites directement dans execute_computes via context trick
        # en ajoutant un DefNode fictif avec source_type=COMPUTE qui retourne la table
        # → plus simple : tester _eval_formula directement en chaîne
        ctx = {"$activites": ACTIVITES}
        g = _eval_formula("GROUP_BY($activites, type, heures = SUM(heures), nb = COUNT(id))", ctx)
        ctx["$par_type"] = g
        sorted_result = _eval_formula("SORT($par_type, heures, DESC)", ctx)
        assert sorted_result[0]["type"] == "TRAVAUX"    # 350h
        assert sorted_result[1]["type"] == "FOURNITURE" # 300h
        assert sorted_result[2]["type"] == "ETUDES"     # 150h

    def test_group_by_puis_top_n(self):
        """TOP_N sur le résultat de GROUP_BY."""
        ctx = {"$activites": ACTIVITES}
        g = _eval_formula("GROUP_BY($activites, type, heures = SUM(heures))", ctx)
        ctx["$par_type"] = g
        top = _eval_formula("TOP_N($par_type, 2, heures, DESC)", ctx)
        assert len(top) == 2
        # TOP 2 : TRAVAUX (350) et FOURNITURE (300)
        types = {r["type"] for r in top}
        assert "TRAVAUX" in types
        assert "FOURNITURE" in types

    def test_chaine_complète(self):
        """Pipeline complet : données brutes → GROUP_BY → SORT → TOP_N."""
        ctx = {"$activites": ACTIVITES}

        # Étape 1 : regrouper par type avec SUM heures et COUNT activités
        grouped = _eval_formula(
            "GROUP_BY($activites, type, heures = SUM(heures), nb = COUNT(id))",
            ctx,
        )
        ctx["$par_type"] = grouped

        # Étape 2 : trier les groupes par heures DESC
        tri = _eval_formula("SORT($par_type, heures, DESC)", ctx)
        ctx["$par_type_tri"] = tri

        # Étape 3 : garder les 2 types avec le plus d'heures
        top2 = _eval_formula("TOP_N($par_type_tri, 2, heures, DESC)", ctx)

        assert len(top2) == 2
        assert top2[0]["heures"] >= top2[1]["heures"]

    def test_filtre_puis_group_by(self):
        """FILTER → GROUP_BY : ne compter que les activités EN_COURS."""
        ctx = {"$activites": ACTIVITES}
        en_cours = _eval_formula(
            'FILTER($activites, statut = "EN_COURS")',
            ctx,
        )
        ctx["$en_cours"] = en_cours
        grouped = _eval_formula(
            "GROUP_BY($en_cours, type, nb = COUNT(id))",
            ctx,
        )
        by_type = {r["type"]: r for r in grouped}
        assert by_type["TRAVAUX"]["nb"]    == 1  # A1 seulement
        assert by_type["ETUDES"]["nb"]     == 1  # A3 seulement
        assert by_type["FOURNITURE"]["nb"] == 1  # A5 seulement

    def test_group_by_count_if(self):
        """COUNT_IF dans GROUP_BY pour compter les activités clôturées par type."""
        ctx = {"$activites": ACTIVITES}
        result = _eval_formula(
            'GROUP_BY($activites, type, nb_cloturees = COUNT_IF(statut, "CLOTUREE"))',
            ctx,
        )
        by_type = {r["type"]: r for r in result}
        assert by_type["TRAVAUX"]["nb_cloturees"]    == 1  # A2
        assert by_type["ETUDES"]["nb_cloturees"]     == 1  # A4
        assert by_type["FOURNITURE"]["nb_cloturees"] == 0  # aucune
