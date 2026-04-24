"""
Tests unitaires — Phase COMPUTE de executor.py
===============================================
Teste toutes les fonctions MXL sans ouvrir de fichier Excel.
Le contexte est un simple dict Python.
"""
import pytest
from src.executor import _eval_formula, _eval_condition


# ─── Fixtures communes ────────────────────────────────────────────────────────

@pytest.fixture
def ctx_activites():
    """Table d'activités standard pour les tests."""
    return {
        "$activites": [
            {"id": "ACT-001", "libelle": "Pose câbles",     "statut": "EN_COURS",  "avancement": 60,  "heures": 100},
            {"id": "ACT-002", "libelle": "Tests unitaires", "statut": "EN_COURS",  "avancement": 30,  "heures": 200},
            {"id": "ACT-003", "libelle": "Mise en service", "statut": "CLOTUREE",  "avancement": 100, "heures": 50},
            {"id": "ACT-004", "libelle": "Documentation",   "statut": "SUSPENDUE", "avancement": 10,  "heures": 80},
        ]
    }


# ─── SUM ─────────────────────────────────────────────────────────────────────

class TestSum:
    def test_sum_basique(self, ctx_activites):
        result = _eval_formula("SUM($activites.heures)", ctx_activites)
        assert result == 430  # 100+200+50+80

    def test_sum_ignore_none(self):
        ctx = {"$t": [{"v": 10}, {"v": None}, {"v": 20}]}
        assert _eval_formula("SUM($t.v)", ctx) == 30

    def test_sum_table_vide(self):
        ctx = {"$t": []}
        assert _eval_formula("SUM($t.v)", ctx) == 0


# ─── COUNT ───────────────────────────────────────────────────────────────────

class TestCount:
    def test_count_basique(self, ctx_activites):
        result = _eval_formula("COUNT($activites.id)", ctx_activites)
        assert result == 4

    def test_count_ignore_none(self):
        ctx = {"$t": [{"id": 1}, {"id": None}, {"id": 3}]}
        assert _eval_formula("COUNT($t.id)", ctx) == 2

    def test_count_if_egal(self, ctx_activites):
        result = _eval_formula('COUNT_IF($activites.statut, "EN_COURS")', ctx_activites)
        assert result == 2

    def test_count_if_zero(self, ctx_activites):
        result = _eval_formula('COUNT_IF($activites.statut, "INEXISTANT")', ctx_activites)
        assert result == 0


# ─── AVG / MIN / MAX ─────────────────────────────────────────────────────────

class TestAggregats:
    def test_avg(self, ctx_activites):
        result = _eval_formula("AVG($activites.avancement)", ctx_activites)
        assert result == pytest.approx(50.0)  # (60+30+100+10)/4

    def test_min(self, ctx_activites):
        assert _eval_formula("MIN($activites.avancement)", ctx_activites) == 10

    def test_max(self, ctx_activites):
        assert _eval_formula("MAX($activites.avancement)", ctx_activites) == 100

    def test_avg_table_vide(self):
        ctx = {"$t": []}
        assert _eval_formula("AVG($t.v)", ctx) == 0.0


# ─── MEAN_WEIGHTED ───────────────────────────────────────────────────────────

class TestMeanWeighted:
    def test_poids_egaux(self):
        ctx = {"$t": [
            {"val": 50, "poids": 100},
            {"val": 80, "poids": 100},
        ]}
        result = _eval_formula("MEAN_WEIGHTED($t.val, $t.poids)", ctx)
        assert result == pytest.approx(65.0)

    def test_poids_differents(self):
        # 60% sur 100h + 30% sur 200h = (60*100 + 30*200) / 300 = 12000/300 = 40%
        ctx = {"$t": [
            {"val": 60, "poids": 100},
            {"val": 30, "poids": 200},
        ]}
        result = _eval_formula("MEAN_WEIGHTED($t.val, $t.poids)", ctx)
        assert result == pytest.approx(40.0)

    def test_poids_zero(self):
        ctx = {"$t": [{"val": 50, "poids": 0}]}
        assert _eval_formula("MEAN_WEIGHTED($t.val, $t.poids)", ctx) == 0.0

    def test_ignore_none(self):
        ctx = {"$t": [
            {"val": 60,   "poids": 100},
            {"val": None, "poids": 50},
            {"val": 80,   "poids": 100},
        ]}
        result = _eval_formula("MEAN_WEIGHTED($t.val, $t.poids)", ctx)
        assert result == pytest.approx(70.0)  # (60*100 + 80*100) / 200


# ─── FILTER ──────────────────────────────────────────────────────────────────

class TestFilter:
    def test_filter_egal(self, ctx_activites):
        result = _eval_formula('FILTER($activites, statut = "CLOTUREE")', ctx_activites)
        assert len(result) == 1
        assert result[0]["id"] == "ACT-003"

    def test_filter_different(self, ctx_activites):
        result = _eval_formula('FILTER($activites, statut != "CLOTUREE")', ctx_activites)
        assert len(result) == 3
        assert all(r["statut"] != "CLOTUREE" for r in result)

    def test_filter_numerique_superieur(self, ctx_activites):
        result = _eval_formula("FILTER($activites, avancement >= 60)", ctx_activites)
        assert len(result) == 2  # ACT-001 (60%) et ACT-003 (100%)

    def test_filter_and(self, ctx_activites):
        result = _eval_formula(
            'FILTER($activites, statut = "EN_COURS" AND avancement > 50)',
            ctx_activites
        )
        assert len(result) == 1
        assert result[0]["id"] == "ACT-001"

    def test_filter_or(self, ctx_activites):
        result = _eval_formula(
            'FILTER($activites, statut = "CLOTUREE" OR statut = "SUSPENDUE")',
            ctx_activites
        )
        assert len(result) == 2

    def test_filter_retourne_liste_vide(self, ctx_activites):
        result = _eval_formula('FILTER($activites, statut = "INCONNU")', ctx_activites)
        assert result == []

    def test_filter_puis_calcul(self, ctx_activites):
        """Enchaîner FILTER et SUM via le contexte."""
        ctx = dict(ctx_activites)
        ctx["$actives"] = _eval_formula(
            'FILTER($activites, statut != "CLOTUREE")', ctx
        )
        total = _eval_formula("SUM($actives.heures)", ctx)
        assert total == 380  # 100+200+80 (pas ACT-003)


# ─── DIV ─────────────────────────────────────────────────────────────────────

class TestDiv:
    def test_div_normal(self):
        ctx = {"$nb_cloturees": 2, "$nb_total": 4}
        result = _eval_formula("DIV($nb_cloturees, $nb_total)", ctx)
        assert result == pytest.approx(0.5)

    def test_div_par_zero(self):
        ctx = {"$nb_cloturees": 2, "$nb_total": 0}
        result = _eval_formula("DIV($nb_cloturees, $nb_total)", ctx)
        assert result == 0.0  # pas d'exception


# ─── TRAFFIC_LIGHT ───────────────────────────────────────────────────────────

class TestTrafficLight:
    def test_rouge(self):
        ctx = {"$avancement": 20.0}
        assert _eval_formula("TRAFFIC_LIGHT($avancement, warn=30, ok=70)", ctx) == "ROUGE"

    def test_orange(self):
        ctx = {"$avancement": 50.0}
        assert _eval_formula("TRAFFIC_LIGHT($avancement, warn=30, ok=70)", ctx) == "ORANGE"

    def test_vert(self):
        ctx = {"$avancement": 75.0}
        assert _eval_formula("TRAFFIC_LIGHT($avancement, warn=30, ok=70)", ctx) == "VERT"

    def test_seuil_exact_warn(self):
        ctx = {"$avancement": 30.0}
        # 30 < 30 est faux → ORANGE
        assert _eval_formula("TRAFFIC_LIGHT($avancement, warn=30, ok=70)", ctx) == "ORANGE"

    def test_seuil_exact_ok(self):
        ctx = {"$avancement": 70.0}
        assert _eval_formula("TRAFFIC_LIGHT($avancement, warn=30, ok=70)", ctx) == "VERT"

    def test_none_retourne_rouge(self):
        ctx = {"$avancement": None}
        assert _eval_formula("TRAFFIC_LIGHT($avancement, warn=30, ok=70)", ctx) == "ROUGE"


# ─── SWITCH_RANGE ────────────────────────────────────────────────────────────

class TestSwitchRange:
    FORMULA = '''SWITCH_RANGE($val,
        [0,    0]   : "NON_DÉMARRÉ",
        [1,   25]   : "DÉMARRAGE",
        [26,  75]   : "EN_COURS",
        [76,  99]   : "FINALISATION",
        [100, 100]  : "TERMINÉ"
    )'''

    def test_zero(self):
        assert _eval_formula(self.FORMULA, {"$val": 0}) == "NON_DÉMARRÉ"

    def test_debut(self):
        assert _eval_formula(self.FORMULA, {"$val": 15}) == "DÉMARRAGE"

    def test_en_cours(self):
        assert _eval_formula(self.FORMULA, {"$val": 50}) == "EN_COURS"

    def test_finalisation(self):
        assert _eval_formula(self.FORMULA, {"$val": 90}) == "FINALISATION"

    def test_termine(self):
        assert _eval_formula(self.FORMULA, {"$val": 100}) == "TERMINÉ"


# ─── IF ──────────────────────────────────────────────────────────────────────

class TestIf:
    def test_if_vrai(self):
        ctx = {"$avancement": 75}
        result = _eval_formula('IF($avancement >= 70, "OUI", "NON")', ctx)
        assert result == "OUI"

    def test_if_faux(self):
        ctx = {"$avancement": 50}
        result = _eval_formula('IF($avancement >= 70, "OUI", "NON")', ctx)
        assert result == "NON"

    def test_if_null(self):
        ctx = {"$val": None}
        result = _eval_formula("IF_NULL($val, 0)", ctx)
        assert result == 0

    def test_if_null_valeur_presente(self):
        ctx = {"$val": 42}
        result = _eval_formula("IF_NULL($val, 0)", ctx)
        assert result == 42


# ─── Scénario intégré ────────────────────────────────────────────────────────

class TestScenarioComplet:
    """Simule l'exécution d'un _Manifeste complet en pur Python."""

    def test_uo_complet(self):
        ctx = {
            "$activites": [
                {"id": "A1", "statut": "EN_COURS",  "avancement": 40, "heures": 100},
                {"id": "A2", "statut": "EN_COURS",  "avancement": 80, "heures": 200},
                {"id": "A3", "statut": "CLOTUREE",  "avancement": 100, "heures": 50},
                {"id": "A4", "statut": "SUSPENDUE", "avancement": 0,   "heures": 80},
            ]
        }

        # Filtrage
        ctx["$actives"]   = _eval_formula('FILTER($activites, statut != "CLOTUREE")', ctx)
        ctx["$cloturees"] = _eval_formula('FILTER($activites, statut = "CLOTUREE")', ctx)

        # Calculs
        ctx["$avancement_global"] = _eval_formula(
            "MEAN_WEIGHTED($actives.avancement, $actives.heures)", ctx
        )
        ctx["$total_heures"]  = _eval_formula("SUM($activites.heures)", ctx)
        ctx["$nb_total"]      = _eval_formula("COUNT($activites.id)", ctx)
        ctx["$nb_cloturees"]  = _eval_formula("COUNT($cloturees.id)", ctx)
        ctx["$taux_cloture"]  = _eval_formula("DIV($nb_cloturees, $nb_total)", ctx)
        ctx["$statut_global"] = _eval_formula(
            "TRAFFIC_LIGHT($avancement_global, warn=30, ok=70)", ctx
        )

        # Vérifications
        # Actives : A1(40h=100), A2(80h=200), A4(0h=80) → (40*100+80*200+0*80)/380
        expected_avancement = (40 * 100 + 80 * 200 + 0 * 80) / (100 + 200 + 80)
        assert ctx["$avancement_global"] == pytest.approx(expected_avancement)
        assert ctx["$total_heures"]  == 430
        assert ctx["$nb_total"]      == 4
        assert ctx["$nb_cloturees"]  == 1
        assert ctx["$taux_cloture"]  == pytest.approx(0.25)
        assert ctx["$statut_global"] in ("ROUGE", "ORANGE", "VERT")
