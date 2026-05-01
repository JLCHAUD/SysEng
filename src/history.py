"""
ExoSync — Historique & Snapshots (M04)
=======================================
Sauvegarde l'historique des runs et des snapshots du store central.

Fonctions publiques :
    save_run_history(resultats, debut, fin)  → Path du fichier run
    save_store_snapshot(store)               → Path du snapshot
    load_run_history(run_path)               → dict
    list_runs()                              → [Path]
    list_snapshots()                         → [Path]
    compare_snapshots(path_a, path_b)        → dict des différences
    history_of_key(key)                      → [(timestamp, valeur)]
    purge_old_files(max_runs, max_snapshots) → int (fichiers supprimés)
"""
import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

ROOT = Path(__file__).parent.parent

HISTORY_DIR   = ROOT / "output" / "history"
SNAPSHOT_DIR  = ROOT / "output" / "snapshots"


# ─── Sauvegarde ───────────────────────────────────────────────────────────────

def save_run_history(
    resultats: List[dict],
    debut: datetime,
    fin: datetime,
) -> Path:
    """
    Sauvegarde un journal de run JSON dans output/history/.

    Args:
        resultats : liste de dicts résultat par fichier (id, statut, log, …)
        debut, fin : horodatages du run

    Returns:
        Chemin du fichier JSON créé.
    """
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    nom = f"run_{debut.strftime('%Y%m%d_%H%M%S')}.json"
    chemin = HISTORY_DIR / nom

    run = {
        "debut":    debut.isoformat(timespec="seconds"),
        "fin":      fin.isoformat(timespec="seconds"),
        "duree_s":  round((fin - debut).total_seconds(), 2),
        "nb_total": len(resultats),
        "nb_ok":    sum(1 for r in resultats if r.get("statut") == "ok"),
        "nb_skip":  sum(1 for r in resultats if r.get("statut") == "skip_verrouille"),
        "nb_erreur":sum(1 for r in resultats if r.get("statut") == "erreur"),
        "fichiers": resultats,
    }
    with open(chemin, "w", encoding="utf-8") as f:
        json.dump(run, f, ensure_ascii=False, indent=2, default=str)
    return chemin


def save_store_snapshot(store_path: Optional[Path] = None) -> Path:
    """
    Sauvegarde un snapshot complet du store JSON dans output/snapshots/.

    Args:
        store_path : chemin du store source (défaut : output/store.json)

    Returns:
        Chemin du snapshot créé.
    """
    from src.store import DEFAULT_STORE_PATH
    src_path = store_path or DEFAULT_STORE_PATH

    SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    dest = SNAPSHOT_DIR / f"store_{ts}.json"

    # Sur Windows le timer système peut avoir ~15ms de résolution :
    # deux appels rapides produiraient le même nom → on ajoute un compteur.
    counter = 0
    while dest.exists():
        counter += 1
        dest = SNAPSHOT_DIR / f"store_{ts}_{counter}.json"

    if src_path.exists():
        shutil.copy2(src_path, dest)
    else:
        # store vide — créer un snapshot vide
        with open(dest, "w", encoding="utf-8") as f:
            json.dump({"variables": {}, "snapshot_ts": ts}, f)
    return dest


# ─── Lecture ──────────────────────────────────────────────────────────────────

def load_run_history(run_path: Path) -> dict:
    """Charge un journal de run depuis son chemin."""
    with open(run_path, encoding="utf-8") as f:
        return json.load(f)


def list_runs(history_dir: Optional[Path] = None) -> List[Path]:
    """Retourne la liste des runs triés du plus récent au plus ancien."""
    d = history_dir or HISTORY_DIR
    if not d.exists():
        return []
    return sorted(d.glob("run_*.json"), reverse=True)


def list_snapshots(snapshot_dir: Optional[Path] = None) -> List[Path]:
    """Retourne la liste des snapshots triés du plus récent au plus ancien."""
    d = snapshot_dir or SNAPSHOT_DIR
    if not d.exists():
        return []
    return sorted(d.glob("store_*.json"), reverse=True)


# ─── Comparaison ──────────────────────────────────────────────────────────────

def compare_snapshots(path_a: Path, path_b: Path) -> Dict[str, Any]:
    """
    Compare deux snapshots du store.

    Returns:
        {
          "ajouts":       {key: val},   # clés dans B mais pas dans A
          "suppressions": {key: val},   # clés dans A mais pas dans B
          "modifications":{key: {"avant": v_a, "apres": v_b}},
          "inchanges":    int,
        }
    """
    def _load_vars(p: Path) -> Dict[str, Any]:
        with open(p, encoding="utf-8") as f:
            data = json.load(f)
        return data.get("variables", data)

    vars_a = _load_vars(path_a)
    vars_b = _load_vars(path_b)

    ajouts       = {k: v for k, v in vars_b.items() if k not in vars_a}
    suppressions = {k: v for k, v in vars_a.items() if k not in vars_b}
    modifications: Dict[str, Any] = {}
    inchanges = 0

    for k in vars_a:
        if k in vars_b:
            if vars_a[k] != vars_b[k]:
                modifications[k] = {"avant": vars_a[k], "apres": vars_b[k]}
            else:
                inchanges += 1

    return {
        "ajouts":        ajouts,
        "suppressions":  suppressions,
        "modifications": modifications,
        "inchanges":     inchanges,
    }


# ─── Historique d'une clé ─────────────────────────────────────────────────────

def history_of_key(key: str, snapshot_dir: Optional[Path] = None) -> List[Tuple[str, Any]]:
    """
    Retourne l'historique des valeurs d'une clé dans tous les snapshots.

    Returns:
        Liste de (timestamp_iso, valeur) du plus ancien au plus récent.
        La valeur est None si la clé était absente à ce moment.
    """
    snapshots = list_snapshots(snapshot_dir)
    result = []
    for snap_path in reversed(snapshots):   # du plus ancien au plus récent
        try:
            with open(snap_path, encoding="utf-8") as f:
                data = json.load(f)
            ts  = data.get("derniere_maj") or snap_path.stem.replace("store_", "")
            val = data.get("variables", {}).get(key)
            result.append((ts, val))
        except Exception:
            continue
    return result


# ─── Purge automatique ────────────────────────────────────────────────────────

def purge_old_files(
    max_runs: int = 30,
    max_snapshots: int = 30,
    history_dir: Optional[Path] = None,
    snapshot_dir: Optional[Path] = None,
) -> int:
    """
    Supprime les fichiers d'historique et de snapshots au-delà du seuil.
    Conserve les N plus récents.

    Returns:
        Nombre total de fichiers supprimés.
    """
    supprime = 0

    for chemin in list_runs(history_dir)[max_runs:]:
        chemin.unlink(missing_ok=True)
        supprime += 1

    for chemin in list_snapshots(snapshot_dir)[max_snapshots:]:
        chemin.unlink(missing_ok=True)
        supprime += 1

    return supprime
