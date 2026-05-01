"""
Tests pour le système de rôles/ownership (feature owner-roles).

Couvre :
- load_file_types() : lecture de file_types.yaml avec owner_role
- EntreeRegistre.owner_id : nouveau champ
- UOInstance.owner_id : nouveau champ
- load_registre() : lecture du owner_id depuis registre.json
- load_uo_instances() : lecture du owner_id depuis uo_instances.json
- validate_owner_roles() : détection des violations rôle/type
"""
import json
from pathlib import Path

import pytest
import yaml


# ─── load_file_types ──────────────────────────────────────────────────────────

class TestLoadFileTypes:
    def test_tous_types_charges(self):
        from src.config_loader import load_file_types
        ft = load_file_types()
        assert set(ft.keys()) == {
            "uo_instance", "referentiel_uo", "referentiel_projet",
            "cockpit", "pilote", "consolidation", "client",
        }

    def test_owner_role_present_sur_chaque_type(self):
        from src.config_loader import load_file_types
        ft = load_file_types()
        for type_id, defn in ft.items():
            assert "owner_role" in defn, f"owner_role manquant sur {type_id}"

    def test_owner_roles_corrects(self):
        from src.config_loader import load_file_types
        ft = load_file_types()
        assert ft["uo_instance"]["owner_role"] == "ingenieur_sys"
        assert ft["cockpit"]["owner_role"] == "pilote_tech"
        assert ft["pilote"]["owner_role"] == "pilote_tech"
        assert ft["client"]["owner_role"] == "donneur_ordre"
        assert ft["consolidation"]["owner_role"] == "engagement_mgr"
        assert ft["referentiel_projet"]["owner_role"] == "engagement_mgr"
        assert ft["referentiel_uo"]["owner_role"] == "it_manager"


# ─── EntreeRegistre.owner_id ──────────────────────────────────────────────────

class TestEntreeRegistreOwner:
    def test_champ_owner_id_existe(self):
        from src.models import EntreeRegistre
        e = EntreeRegistre(
            id="TEST", type_fichier="uo_instance",
            chemin="test.xlsx", synchro_periodicite="quotidien",
        )
        assert e.owner_id is None   # valeur par défaut

    def test_champ_owner_id_assignable(self):
        from src.models import EntreeRegistre
        e = EntreeRegistre(
            id="UO-001", type_fichier="uo_instance",
            chemin="test.xlsx", synchro_periodicite="quotidien",
            owner_id="USR001",
        )
        assert e.owner_id == "USR001"


# ─── UOInstance.owner_id ──────────────────────────────────────────────────────

class TestUOInstanceOwner:
    def test_champ_owner_id_existe(self):
        from src.models import UOInstance, StatutUO
        from datetime import date
        uo = UOInstance(
            id="UO-001", uo_type_id="spec_fonctionnelle",
            system_id="clim", project_id="MI20",
            engineer_name="Alice Dubois",
            total_hours=32,
            start_date=date(2026, 4, 1),
            end_date=date(2026, 5, 30),
        )
        assert uo.owner_id is None

    def test_champ_owner_id_assignable(self):
        from src.models import UOInstance, StatutUO
        from datetime import date
        uo = UOInstance(
            id="UO-001", uo_type_id="spec_fonctionnelle",
            system_id="clim", project_id="MI20",
            engineer_name="Alice Dubois",
            total_hours=32,
            start_date=date(2026, 4, 1),
            end_date=date(2026, 5, 30),
            owner_id="USR001",
        )
        assert uo.owner_id == "USR001"


# ─── load_registre — owner_id lu ──────────────────────────────────────────────

class TestLoadRegistreOwner:
    def test_uo001_a_owner_id(self):
        from src.config_loader import load_registre
        entrees = {e.id: e for e in load_registre()}
        assert "UO-001" in entrees
        assert entrees["UO-001"].owner_id == "USR001"

    def test_uo003_a_owner_bruno(self):
        from src.config_loader import load_registre
        entrees = {e.id: e for e in load_registre()}
        assert entrees["UO-003"].owner_id == "USR002"

    def test_registre_contient_5_uo(self):
        """Le registre ne contient que les 5 UO actives (référentiels/cockpits supprimés)."""
        from src.config_loader import load_registre
        entrees = load_registre()
        assert len(entrees) == 5
        ids = {e.id for e in entrees}
        assert ids == {"UO-001", "UO-002", "UO-003", "UO-004", "UO-005"}

    def test_uo002_owner_alice(self):
        from src.config_loader import load_registre
        entrees = {e.id: e for e in load_registre()}
        assert entrees["UO-002"].owner_id == "USR001"

    def test_uo004_owner_bruno(self):
        from src.config_loader import load_registre
        entrees = {e.id: e for e in load_registre()}
        assert entrees["UO-004"].owner_id == "USR002"


# ─── load_uo_instances — owner_id lu ─────────────────────────────────────────

class TestLoadUOInstancesOwner:
    def test_uo001_owner_alice(self):
        from src.config_loader import load_uo_instances
        instances = {i.id: i for i in load_uo_instances()}
        assert instances["UO-001"].owner_id == "USR001"

    def test_uo003_owner_bruno(self):
        from src.config_loader import load_uo_instances
        instances = {i.id: i for i in load_uo_instances()}
        assert instances["UO-003"].owner_id == "USR002"

    def test_uo005_owner_camille(self):
        from src.config_loader import load_uo_instances
        instances = {i.id: i for i in load_uo_instances()}
        assert instances["UO-005"].owner_id == "USR003"

    def test_engineer_name_toujours_present(self):
        """Rétro-compat : engineer_name doit rester même si owner_id est défini."""
        from src.config_loader import load_uo_instances
        instances = {i.id: i for i in load_uo_instances()}
        assert instances["UO-001"].engineer_name == "Alice Dubois"


# ─── validate_owner_roles ─────────────────────────────────────────────────────

class TestValidateOwnerRoles:
    def test_registre_reel_sans_violation(self):
        """Le registre.json réel doit être cohérent (0 violation)."""
        from src.config_loader import validate_owner_roles
        violations = validate_owner_roles()
        assert violations == [], (
            f"Violations inattendues : {[str(v) for v in violations]}"
        )

    def test_violation_role_incorrect(self, tmp_path):
        """Un fichier uo_instance avec un owner pilote_tech doit déclencher une violation."""
        from src.config_loader import validate_owner_roles, load_acteurs, load_file_types
        from src.models import EntreeRegistre

        # Fabriquer un registre avec une violation volontaire
        entrees = [
            EntreeRegistre(
                id="UO-TEST", type_fichier="uo_instance",
                chemin="test.xlsx", synchro_periodicite="quotidien",
                owner_id="USR004",  # Jean-Luc, pilote_tech — attendu: ingenieur_sys
            )
        ]
        violations = validate_owner_roles(entrees)
        assert len(violations) == 1
        v = violations[0]
        assert v.file_id == "UO-TEST"
        assert v.expected_role == "ingenieur_sys"
        assert v.owner_role == "pilote_tech"

    def test_violation_owner_absent(self):
        """Un fichier avec un owner_role attendu mais sans owner_id doit signaler une violation."""
        from src.config_loader import validate_owner_roles
        from src.models import EntreeRegistre

        entrees = [
            EntreeRegistre(
                id="UO-SANS-OWNER", type_fichier="uo_instance",
                chemin="test.xlsx", synchro_periodicite="quotidien",
                owner_id=None,
            )
        ]
        violations = validate_owner_roles(entrees)
        assert len(violations) == 1
        assert violations[0].owner_id == ""
        assert violations[0].owner_nom == "<non assigné>"

    def test_violation_owner_inconnu(self):
        """Un owner_id qui n'existe pas dans acteurs.json doit déclencher une violation."""
        from src.config_loader import validate_owner_roles
        from src.models import EntreeRegistre

        entrees = [
            EntreeRegistre(
                id="UO-X", type_fichier="uo_instance",
                chemin="test.xlsx", synchro_periodicite="quotidien",
                owner_id="USR999",
            )
        ]
        violations = validate_owner_roles(entrees)
        assert len(violations) == 1
        assert violations[0].owner_nom == "<inconnu>"

    def test_type_sans_owner_role_pas_de_violation(self):
        """Un type de fichier sans owner_role dans file_types.yaml ne génère pas de violation."""
        from src.config_loader import validate_owner_roles
        from src.models import EntreeRegistre

        # "type_inconnu" n'existe pas dans file_types.yaml → pas de contrainte
        entrees = [
            EntreeRegistre(
                id="X", type_fichier="type_inexistant",
                chemin="test.xlsx", synchro_periodicite="quotidien",
                owner_id=None,
            )
        ]
        violations = validate_owner_roles(entrees)
        assert violations == []

    def test_violation_str(self):
        """__str__ de OwnerRoleViolation doit être lisible."""
        from src.config_loader import OwnerRoleViolation
        v = OwnerRoleViolation(
            file_id="UO-001", type_fichier="uo_instance",
            owner_id="USR004", owner_nom="Jean-Luc Bernard",
            owner_role="pilote_tech", expected_role="ingenieur_sys",
        )
        s = str(v)
        assert "UO-001" in s
        assert "pilote_tech" in s
        assert "ingenieur_sys" in s

    def test_registre_reel_uniquement_uo(self):
        """Le registre actif ne contient que des uo_instance — tous valides."""
        from src.config_loader import load_registre, validate_owner_roles
        entrees = load_registre()
        assert all(e.type_fichier == "uo_instance" for e in entrees)
        violations = validate_owner_roles(entrees)
        assert violations == []

    def test_toutes_uo_instances_valides(self):
        """Les 5 UO du registre ont des owners ingenieur_sys → pas de violation."""
        from src.config_loader import load_registre, validate_owner_roles
        uo_entrees = [e for e in load_registre() if e.type_fichier == "uo_instance"]
        assert len(uo_entrees) == 5
        violations = validate_owner_roles(uo_entrees)
        assert violations == []
