"""ExoSync Studio N2 — Registry Populator (port 8000)."""
from pathlib import Path
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from web.registry_app.api import (
    registry,
    actors,
    hierarchy,
    tables,
    ecosystem,
    excel_import,
    mxl,
    file_types_ro,
)

app = FastAPI(title="ExoSync Studio N2 — Registry Populator", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(registry.router,      prefix="/api/registry",    tags=["registry"])
app.include_router(actors.router,        prefix="/api/actors",      tags=["actors"])
app.include_router(hierarchy.router,     prefix="/api/hierarchy",   tags=["hierarchy"])
app.include_router(tables.router,        prefix="/api/tables",      tags=["tables"])
app.include_router(ecosystem.router,     prefix="/api/ecosystem",   tags=["ecosystem"])
app.include_router(excel_import.router,  prefix="/api/excel",       tags=["excel"])
app.include_router(mxl.router,           prefix="/api/mxl",         tags=["mxl"])
app.include_router(file_types_ro.router, prefix="/api/file-types",  tags=["file-types"])

static_dir = Path(__file__).parent / "static"
if static_dir.exists():
    app.mount("/", StaticFiles(directory=str(static_dir), html=True), name="static")
