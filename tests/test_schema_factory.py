"""Tests for SchemaConfigService and router factory pattern."""
import json
import pytest
import yaml
from dataclasses import fields
from pathlib import Path
from fastapi import FastAPI
from fastapi.testclient import TestClient
from web.schema_config import SchemaConfigService
from web.schema_app.api import classes as classes_mod
from web.schema_app.api import relations as relations_mod
from web.registry_app.api import registry as registry_mod
from web.registry_app.services import config_service as reg_cfg


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


def _make_cfg_rel(tmp_path: Path) -> SchemaConfigService:
    """Crée une SchemaConfigService avec relations file-backed."""
    cfg_dir = tmp_path / "config"
    cfg_dir.mkdir()
    ft_file = cfg_dir / "file_types.yaml"
    ft_file.write_text("file_types: {}\n", encoding="utf-8")
    tbl_file = cfg_dir / "tables.json"
    tbl_file.write_text('{"version":"1","tables":{}}', encoding="utf-8")
    rel_file = cfg_dir / "relations.json"
    rel_file.write_text('[]', encoding="utf-8")

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

    def load_rel():
        import json
        return json.loads(rel_file.read_text())

    def save_rel(data):
        rel_file.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    def noop_load(): return []
    def noop_save(x): pass

    return SchemaConfigService(
        load_file_types=load_ft, save_file_types=save_ft,
        load_tables=load_tbl, save_tables=save_tbl,
        load_relations=load_rel, save_relations=save_rel,
        load_namespaces=noop_load, save_namespaces=noop_save,
        load_functions=noop_load, save_functions=noop_save,
        load_templates=noop_load, save_templates=noop_save,
    )


def _make_full_cfg(tmp_path: Path) -> SchemaConfigService:
    """Crée une SchemaConfigService avec tous les fichiers file-backed."""
    cfg_dir = tmp_path / "config"
    cfg_dir.mkdir()
    ft_file = cfg_dir / "file_types.yaml"
    ft_file.write_text("file_types: {}\n", encoding="utf-8")
    tbl_file = cfg_dir / "tables.json"
    tbl_file.write_text('{"version":"1","tables":{}}', encoding="utf-8")
    rel_file = cfg_dir / "relations.json"
    rel_file.write_text('[]', encoding="utf-8")
    ns_file = cfg_dir / "namespaces.json"
    ns_file.write_text('[]', encoding="utf-8")
    fn_file = cfg_dir / "functions.json"
    fn_file.write_text('[]', encoding="utf-8")
    tpl_file = cfg_dir / "templates.json"
    tpl_file.write_text('[]', encoding="utf-8")

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

    def load_rel():
        import json
        return json.loads(rel_file.read_text())

    def save_rel(data):
        rel_file.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    def load_ns():
        import json
        return json.loads(ns_file.read_text())

    def save_ns(data):
        ns_file.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    def load_fn():
        import json
        return json.loads(fn_file.read_text())

    def save_fn(data):
        fn_file.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    def load_tpl():
        import json
        return json.loads(tpl_file.read_text())

    def save_tpl(data):
        tpl_file.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

    return SchemaConfigService(
        load_file_types=load_ft, save_file_types=save_ft,
        load_tables=load_tbl, save_tables=save_tbl,
        load_relations=load_rel, save_relations=save_rel,
        load_namespaces=load_ns, save_namespaces=save_ns,
        load_functions=load_fn, save_functions=save_fn,
        load_templates=load_tpl, save_templates=save_tpl,
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


def test_relations_make_router_crud(tmp_path):
    cfg = _make_cfg_rel(tmp_path)
    app = FastAPI()
    app.include_router(relations_mod.make_router(cfg), prefix="/api/relations")
    client = TestClient(app)

    body = {"parent_class": "cls_a", "child_class": "cls_b",
            "qualifier": "TYPICAL", "cardinality": "1..N",
            "description": "", "flux": []}
    r = client.post("/api/relations", json=body)
    assert r.status_code == 201
    rel_id = r.json()["id"]

    r2 = client.get("/api/relations")
    assert len(r2.json()) == 1

    r3 = client.delete(f"/api/relations/{rel_id}")
    assert r3.status_code == 204


def test_namespaces_make_router(tmp_path):
    from web.schema_app.api import namespaces as ns_mod
    cfg = _make_full_cfg(tmp_path)
    app = FastAPI()
    app.include_router(ns_mod.make_router(cfg), prefix="/api/namespaces")
    client = TestClient(app)
    r = client.post("/api/namespaces", json={"id": "ns1", "label": "NS1", "prefix": "ns1.", "description": ""})
    assert r.status_code == 201
    r2 = client.get("/api/namespaces")
    assert len(r2.json()) == 1


def test_functions_make_router(tmp_path):
    from web.schema_app.api import functions as fn_mod
    cfg = _make_full_cfg(tmp_path)
    app = FastAPI()
    app.include_router(fn_mod.make_router(cfg), prefix="/api/functions")
    client = TestClient(app)
    r = client.post("/api/functions", json={"label": "Pilote", "description": "", "side": "interne"})
    assert r.status_code == 201
    fn_id = r.json()["id"]
    r2 = client.get(f"/api/functions/{fn_id}")
    assert r2.json()["label"] == "Pilote"


def test_templates_make_router(tmp_path):
    from web.schema_app.api import templates as tpl_mod
    cfg = _make_full_cfg(tmp_path)
    # Créer une Classe d'abord pour la validation template
    app = FastAPI()
    app.include_router(classes_mod.make_router(cfg), prefix="/api/classes")
    app.include_router(tpl_mod.make_router(cfg), prefix="/api/templates")
    client = TestClient(app)
    client.post("/api/classes", json={
        "id": "cls1", "label": "C1", "description": "", "owner_function": "",
        "min_sheets": [], "optional_sheets": [], "allowed_namespaces": [],
        "push_prefix": "", "template": "", "min_fields": [], "std_tables": [],
    })
    r = client.post("/api/templates", json={
        "label": "Tpl1", "class_id": "cls1", "description": "",
        "extra_sheets": [], "field_defaults": {}, "std_tables": [],
        "mxl_defaults": {}, "source_file": "",
    })
    assert r.status_code == 201


def test_file_instance_has_schema_fields():
    from web.schemas.models import FileInstance
    f = FileInstance(
        id="F-001", type_fichier="cls1", chemin="/path/f.xlsx",
        schema_version=3, schema_outdated=True,
    )
    assert f.schema_version == 3
    assert f.schema_outdated is True


def test_file_instance_schema_fields_optional():
    from web.schemas.models import FileInstance
    f = FileInstance(id="F-002", type_fichier=None, chemin="/path/f.xlsx")
    assert f.schema_version is None
    assert f.schema_outdated is None


def _make_registry_app(tmp_path):
    cfg_dir = tmp_path / "config"
    cfg_dir.mkdir()

    ft_file = cfg_dir / "file_types.yaml"
    ft_file.write_text(
        "file_types:\n  cls1:\n    label: C1\n    schema_version: 3\n    required_sheets: []\n    optional_sheets: []\n",
        encoding="utf-8",
    )

    reg_file = cfg_dir / "registre.json"
    reg_file.write_text(json.dumps({"version": "1", "fichiers": [
        {"id": "P-001", "type_fichier": "cls1", "chemin": "/p1.xlsx",
         "synchro_periodicite": "manuel", "owner_role": "", "genere_par_script": False,
         "schema_version": 3},
        {"id": "P-002", "type_fichier": "cls1", "chemin": "/p2.xlsx",
         "synchro_periodicite": "manuel", "owner_role": "", "genere_par_script": False,
         "schema_version": 1},
    ]}), encoding="utf-8")

    reg_cfg.set_active_config(cfg_dir)

    app = FastAPI()
    app.include_router(registry_mod.router, prefix="/api/registry")
    return TestClient(app)


def test_registry_schema_outdated_flag(tmp_path):
    client = _make_registry_app(tmp_path)
    r = client.get("/api/registry")
    assert r.status_code == 200
    posts = {p["id"]: p for p in r.json()}
    assert posts["P-001"]["schema_outdated"] is False
    assert posts["P-002"]["schema_outdated"] is True
