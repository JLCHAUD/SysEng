"""ExoSync Studio N1 — Schema Designer (port 8001)."""
from pathlib import Path
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from web.schema_app.api import (
    ecosystem_manager,
    classes,
    relations,
    functions,
    templates,
    blueprint,
    mxl,
    excel_n1,
)

app = FastAPI(title="ExoSync Studio N1 — Schema Designer", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(ecosystem_manager.router, prefix="/api/ecosystem", tags=["ecosystem"])
app.include_router(classes.router,           prefix="/api/classes",   tags=["classes"])
app.include_router(relations.router,         prefix="/api/relations", tags=["relations"])
app.include_router(functions.router,         prefix="/api/functions", tags=["functions"])
app.include_router(templates.router,         prefix="/api/templates", tags=["templates"])
app.include_router(blueprint.router,         prefix="/api/blueprint", tags=["blueprint"])
app.include_router(mxl.router,               prefix="/api/mxl",       tags=["mxl"])
app.include_router(excel_n1.router,          prefix="/api/excel",     tags=["excel"])

static_dir = Path(__file__).parent / "static"
if static_dir.exists():
    app.mount("/", StaticFiles(directory=str(static_dir), html=True), name="static")
