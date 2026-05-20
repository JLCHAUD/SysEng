"""Service workspace global — partagé entre N1 et N2.

Stocke gabarits_dir et workspace_dir dans .exosync_workspace.json
à la racine du projet Python. Indépendant des écosystèmes actifs.
"""
import json
from pathlib import Path

_PROJECT_ROOT   = Path(__file__).parent.parent
_WORKSPACE_FILE = _PROJECT_ROOT / ".exosync_workspace.json"


def load_workspace() -> dict:
    if not _WORKSPACE_FILE.exists():
        return {"gabarits_dir": "", "workspace_dir": ""}
    with open(_WORKSPACE_FILE, encoding="utf-8") as f:
        return json.load(f)


def save_workspace(data: dict) -> None:
    with open(_WORKSPACE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_gabarits_dir() -> Path | None:
    d = load_workspace()
    g = d.get("gabarits_dir", "")
    return Path(g) if g else None


def get_workspace_dir() -> Path | None:
    d = load_workspace()
    w = d.get("workspace_dir", "")
    return Path(w) if w else None
