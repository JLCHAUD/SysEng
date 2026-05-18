from fastapi import APIRouter
from web.schemas.models import EcosystemGraph, EcosystemNode, EcosystemEdge
from web.services.config_service import load_ecosystem

router = APIRouter()


@router.get("/graph", response_model=EcosystemGraph)
def get_graph():
    eco = load_ecosystem()
    nodes = [
        EcosystemNode(
            id=fid,
            label=fid,
            file_type=fdata.get("file_type", "unknown"),
            status=fdata.get("status", "unknown"),
            path=fdata.get("path", ""),
        )
        for fid, fdata in eco.get("files", {}).items()
    ]
    edges = [
        EcosystemEdge(
            edge_type=e["edge_type"],
            from_node=e["from_node"],
            to_node=e["to_node"],
            mode=e.get("mode", ""),
        )
        for e in eco.get("edges", [])
    ]
    return EcosystemGraph(
        version=eco.get("version", "2.0"),
        last_scan=eco.get("last_scan"),
        nodes=nodes,
        edges=edges,
    )


@router.get("/status")
def get_status():
    eco = load_ecosystem()
    files = eco.get("files", {})
    total = len(files)
    ok = sum(1 for f in files.values() if f.get("status") == "ok")
    return {
        "total_files": total,
        "ok": ok,
        "errors": total - ok,
        "last_scan": eco.get("last_scan"),
        "edge_count": len(eco.get("edges", [])),
    }
