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
