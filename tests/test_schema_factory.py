"""Tests for SchemaConfigService and router factory pattern."""
import pytest
from dataclasses import fields
from web.schema_config import SchemaConfigService


def test_schema_config_service_has_required_fields():
    """Verify SchemaConfigService has all required field names."""
    f_names = {f.name for f in fields(SchemaConfigService)}
    assert "load_file_types" in f_names
    assert "save_file_types" in f_names
    assert "load_tables" in f_names
    assert "save_tables" in f_names
    assert "load_relations" in f_names
    assert "save_relations" in f_names
    assert "load_namespaces" in f_names
    assert "save_namespaces" in f_names
    assert "load_functions" in f_names
    assert "save_functions" in f_names
    assert "load_templates" in f_names
    assert "save_templates" in f_names


import json, yaml
from pathlib import Path
from fastapi.testclient import TestClient
from fastapi import FastAPI
from web.schema_config import SchemaConfigService
from web.schema_app.api import classes as classes_mod


def _make_cfg(tmp_path: Path) -> SchemaConfigService:
    """Crée une SchemaConfigService pointant vers tmp_path."""
    cfg_dir = tmp_path / "config"
    cfg_dir.mkdir()
    ft_file = cfg_dir / "file_types.yaml"
    ft_file.write_text("file_types: {}\n", encoding="utf-8")
    tbl_file = cfg_dir / "tables.json"
    tbl_file.write_text('{"version":"1","tables":{}}', encoding="utf-8")

    def load_ft():
        import yaml
        return (yaml.safe_load(ft_file.read_text()) or {}).get("file_types", {})

    def save_ft(types):
        import yaml
        ft_file.write_text(
            yaml.dump({"file_types": types}, allow_unicode=True, default_flow_style=False),
            encoding="utf-8",
        )

    def load_tbl():
        import json
        return json.loads(tbl_file.read_text())

    def save_tbl(data):
        tbl_file.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    def noop_load(): return []
    def noop_save(x): pass

    return SchemaConfigService(
        load_file_types=load_ft, save_file_types=save_ft,
        load_tables=load_tbl, save_tables=save_tbl,
        load_relations=noop_load, save_relations=noop_save,
        load_namespaces=noop_load, save_namespaces=noop_save,
        load_functions=noop_load, save_functions=noop_save,
        load_templates=noop_load, save_templates=noop_save,
    )


def test_classes_make_router_creates_class(tmp_path):
    cfg = _make_cfg(tmp_path)
    app = FastAPI()
    app.include_router(classes_mod.make_router(cfg), prefix="/api/classes")
    client = TestClient(app)

    body = {
        "id": "test_class", "label": "Test", "description": "",
        "owner_function": "", "min_sheets": [], "optional_sheets": [],
        "allowed_namespaces": [], "push_prefix": "", "template": "",
        "min_fields": [], "std_tables": [],
    }
    r = client.post("/api/classes", json=body)
    assert r.status_code == 201
    data = r.json()
    assert data["id"] == "test_class"
    assert data["schema_version"] == 1


def test_classes_make_router_increments_schema_version(tmp_path):
    cfg = _make_cfg(tmp_path)
    app = FastAPI()
    app.include_router(classes_mod.make_router(cfg), prefix="/api/classes")
    client = TestClient(app)

    body = {
        "id": "cls1", "label": "Classe 1", "description": "",
        "owner_function": "", "min_sheets": [], "optional_sheets": [],
        "allowed_namespaces": [], "push_prefix": "", "template": "",
        "min_fields": [], "std_tables": [],
    }
    client.post("/api/classes", json=body)
    body["label"] = "Classe 1 modifiée"
    r = client.put("/api/classes/cls1", json=body)
    assert r.status_code == 200
    assert r.json()["schema_version"] == 2
