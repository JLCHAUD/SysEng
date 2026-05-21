"""Dataclass de configuration partagée pour les routers schema N1/N2."""
from __future__ import annotations
from dataclasses import dataclass
from typing import Callable


@dataclass
class SchemaConfigService:
    """Dependency injection container for schema router configuration.

    Holds callable references to load/save functions so routers can work
    with different config directories without knowing about N1 or N2 config
    service specifics.
    """
    load_file_types:  Callable[[], dict]
    save_file_types:  Callable[[dict], None]
    load_tables:      Callable[[], dict]
    save_tables:      Callable[[dict], None]
    load_relations:   Callable[[], list]
    save_relations:   Callable[[list], None]
    load_namespaces:  Callable[[], list]
    save_namespaces:  Callable[[list], None]
    load_functions:   Callable[[], list]
    save_functions:   Callable[[list], None]
    load_templates:   Callable[[], list]
    save_templates:   Callable[[list], None]
