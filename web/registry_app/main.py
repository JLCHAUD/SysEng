"""ExoSync Studio N2 — Registry Populator (port 8000)."""
from pathlib import Path
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from web.registry_app.api import (
    registry,
    actors,
    functions,
    hierarchy,
    tables,
    ecosystem,
    ecosystem_manager,
    excel_import,
    mxl,
    file_types_ro,
    gabarits,
    sync,
    directories,
    workspace,
    xlsx_generator,
)

from web.schema_config import SchemaConfigService
from web.registry_app.services.config_service import (
    load_file_types, save_file_types,
    load_tables,     save_tables,
    load_relations,  save_relations,
    load_functions,  save_functions,
    load_templates,  save_templates,
)
from web.schema_app.api import (
    classes    as schema_classes,
    relations  as schema_relations,
    namespaces as schema_namespaces,
    functions  as schema_functions,
    templates  as schema_templates,
)

app = FastAPI(title="ExoSync Studio N2 — Registry Populator", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

schema_cfg = SchemaConfigService(
    load_file_types=load_file_types, save_file_types=save_file_types,
    load_tables=load_tables,         save_tables=save_tables,
    load_relations=load_relations,   save_relations=save_relations,
    load_namespaces=lambda: [],      save_namespaces=lambda x: None,
    load_functions=load_functions,   save_functions=save_functions,
    load_templates=load_templates,   save_templates=save_templates,
)

app.include_router(schema_classes.make_router(schema_cfg),    prefix="/api/schema/classes",    tags=["schema"])
app.include_router(schema_relations.make_router(schema_cfg),  prefix="/api/schema/relations",  tags=["schema"])
app.include_router(schema_namespaces.make_router(schema_cfg), prefix="/api/schema/namespaces", tags=["schema"])
app.include_router(schema_functions.make_router(schema_cfg),  prefix="/api/schema/functions",  tags=["schema"])
app.include_router(schema_templates.make_router(schema_cfg),  prefix="/api/schema/templates",  tags=["schema"])

app.include_router(registry.router,          prefix="/api/registry",    tags=["registry"])
app.include_router(actors.router,            prefix="/api/actors",      tags=["actors"])
app.include_router(functions.router,         prefix="/api/functions",   tags=["functions"])
app.include_router(hierarchy.router,         prefix="/api/hierarchy",   tags=["hierarchy"])
app.include_router(tables.router,            prefix="/api/tables",      tags=["tables"])
app.include_router(ecosystem_manager.router, prefix="/api/ecosystem",   tags=["ecosystem"])
app.include_router(ecosystem.router,         prefix="/api/ecosystem",   tags=["ecosystem"])
app.include_router(excel_import.router,  prefix="/api/excel",       tags=["excel"])
app.include_router(mxl.router,           prefix="/api/mxl",         tags=["mxl"])
app.include_router(file_types_ro.router, prefix="/api/file-types",  tags=["file-types"])
app.include_router(gabarits.router,      prefix="/api/gabarits",    tags=["gabarits"])
app.include_router(sync.router,          prefix="/api/sync",         tags=["sync"])
app.include_router(directories.router,   prefix="/api/directories",  tags=["directories"])
app.include_router(workspace.router,       prefix="/api/workspace",    tags=["workspace"])
app.include_router(xlsx_generator.router, prefix="/api/xlsx",         tags=["xlsx"])

static_dir = Path(__file__).parent / "static"
if static_dir.exists():
    app.mount("/", StaticFiles(directory=str(static_dir), html=True), name="static")
