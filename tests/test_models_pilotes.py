import json
from pathlib import Path
from unittest.mock import patch

def test_uo_instance_has_pilotes_field():
    """UOInstance doit avoir un champ pilotes vide par défaut."""
    from src.models import UOInstance, StatutUO
    from datetime import date
    uo = UOInstance(
        id="X", uo_type_id="T", system_id="S", project_id="P",
        engineer_name="Alice", total_hours=10,
        start_date=date(2026,1,1), end_date=date(2026,6,1),
    )
    assert hasattr(uo, "pilotes")
    assert uo.pilotes == {}


def test_uo_instance_pilotes_populated():
    """Le champ pilotes accepte un dict rôle → id."""
    from src.models import UOInstance, StatutUO
    from datetime import date
    uo = UOInstance(
        id="X", uo_type_id="T", system_id="S", project_id="P",
        engineer_name="Alice", total_hours=10,
        start_date=date(2026,1,1), end_date=date(2026,6,1),
        pilotes={"metier_ts": "USR004", "metier_projet": "USR007"},
    )
    assert uo.pilotes["metier_ts"] == "USR004"
    assert uo.pilotes["metier_projet"] == "USR007"


def test_config_loader_reads_pilotes(tmp_path):
    """load_uo_instances() doit lire le champ pilotes depuis le JSON."""
    import json
    from unittest.mock import patch
    from src.config_loader import load_uo_instances

    uo_data = [{
        "id": "UO-TEST",
        "uo_type_id": "TS",
        "system_id": "SYS1",
        "project_id": "PRJ1",
        "engineer_name": "Alice Dubois",
        "total_hours": 32,
        "start_date": "2026-01-01",
        "end_date": "2026-06-30",
        "statut": "EN_COURS",
        "pilotes": {"metier_ts": "USR004"},
    }]

    with patch("src.config_loader._load_json") as mock_load:
        def side_effect(filename):
            if filename == "uo_instances.json":
                return uo_data
            return {} if filename.endswith(".json") else []
        mock_load.side_effect = side_effect

        with patch("src.config_loader.load_uo_types", return_value={}), \
             patch("src.config_loader.load_systems", return_value={}), \
             patch("src.config_loader.load_projects", return_value={}):
            instances = load_uo_instances()

    assert len(instances) == 1
    assert instances[0].pilotes == {"metier_ts": "USR004"}
