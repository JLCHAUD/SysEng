from pathlib import Path
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse

from web.api import file_types, registry, actors, ecosystem, hierarchy, tables, mxl_generator, schema, xlsx_generator

BASE_DIR = Path(__file__).parent.parent

app = FastAPI(title="ExoSync Studio", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(file_types.router, prefix="/api/file-types", tags=["file-types"])
app.include_router(registry.router, prefix="/api/registry", tags=["registry"])
app.include_router(actors.router, prefix="/api/actors", tags=["actors"])
app.include_router(ecosystem.router,     prefix="/api/ecosystem",    tags=["ecosystem"])
app.include_router(hierarchy.router,    prefix="/api/hierarchy",    tags=["hierarchy"])
app.include_router(tables.router,       prefix="/api/tables",       tags=["tables"])
app.include_router(mxl_generator.router,  prefix="/api/mxl",          tags=["mxl"])
app.include_router(xlsx_generator.router, prefix="/api/xlsx",         tags=["xlsx"])
app.include_router(schema.router,         prefix="/api/schema",       tags=["schema"])

static_dir = Path(__file__).parent / "static"
if static_dir.exists():
    app.mount("/", StaticFiles(directory=str(static_dir), html=True), name="static")
