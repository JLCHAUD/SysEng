"""ExoSync Studio N1 — Schema Designer (port 8001)."""
from pathlib import Path
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from web.schema_config import SchemaConfigService
from web.schema_app.services import config_service
from web.schema_app.api import (
    ecosystem_manager,
    classes,
    relations,
    functions,
    templates,
    namespaces,
    blueprint,
    mxl,
    excel_n1,
    workspace,
)

app = FastAPI(title="ExoSync Studio N1 — Schema Designer", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Create SchemaConfigService with config_service functions
cfg = SchemaConfigService(
    load_file_types=config_service.load_file_types,
    save_file_types=config_service.save_file_types,
    load_tables=config_service.load_tables,
    save_tables=config_service.save_tables,
    load_relations=config_service.load_relations,
    save_relations=config_service.save_relations,
    load_namespaces=config_service.load_namespaces,
    save_namespaces=config_service.save_namespaces,
    load_functions=config_service.load_functions,
    save_functions=config_service.save_functions,
    load_templates=config_service.load_templates,
    save_templates=config_service.save_templates,
)

app.include_router(ecosystem_manager.router, prefix="/api/ecosystem", tags=["ecosystem"])
app.include_router(classes.make_router(cfg),           prefix="/api/classes",   tags=["classes"])
app.include_router(relations.make_router(cfg),         prefix="/api/relations", tags=["relations"])
app.include_router(functions.make_router(cfg),         prefix="/api/functions", tags=["functions"])
app.include_router(templates.make_router(cfg),         prefix="/api/templates",  tags=["templates"])
app.include_router(namespaces.make_router(cfg),        prefix="/api/namespaces", tags=["namespaces"])
app.include_router(blueprint.router,         prefix="/api/blueprint",  tags=["blueprint"])
app.include_router(mxl.router,               prefix="/api/mxl",       tags=["mxl"])
app.include_router(excel_n1.router,          prefix="/api/excel",     tags=["excel"])
app.include_router(workspace.router,         prefix="/api/workspace", tags=["workspace"])

static_dir = Path(__file__).parent / "static"
if static_dir.exists():
    app.mount("/", StaticFiles(directory=str(static_dir), html=True), name="static")
