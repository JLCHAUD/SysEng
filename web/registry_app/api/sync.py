"""API de déclenchement du moteur ExoSync depuis le Registry Populator N2.

Pont entre l'écosystème actif N2 (registre.json dynamique) et le moteur
src/sync.py qui utilise historiquement src/config_loader (CONFIG_DIR fixe).

Stratégie :
  1. Lire le registre N2 (écosystème actif)
  2. Convertir en EntreeRegistre (src.models)
  3. Appeler _sync_fichier() directement pour chaque entrée
  4. Mettre à jour le registre N2 avec statut + timestamp
  5. Retourner le rapport
"""
import asyncio
import json
from datetime import datetime
from pathlib import Path
from typing import Optional

from fastapi import APIRouter, HTTPException
from pydantic import BaseModel

# Moteur ExoSync
from src.sync import _sync_fichier, _generer_rapport, ORDRE_TYPES
from src.models import EntreeRegistre
from src import history as History

# Config N2
from web.registry_app.services.config_service import (
    get_active_config,
    load_registre as n2_load_registre,
    save_registre as n2_save_registre,
)
from web.registry_app.api.directories import get_posts_base_path

router = APIRouter()
ROOT = Path(__file__).parent.parent.parent.parent

# ── État du sync (simple verrou pour éviter les lancements parallèles) ────────
_sync_running: bool = False
_last_rapport: Optional[dict] = None


# ── Modèles ───────────────────────────────────────────────────────────────────

class SyncRequest(BaseModel):
    ids: list[str] = []      # vide = tout synchroniser
    force: bool = False      # ignorer vérification verrouillage Excel


class SyncStatusResponse(BaseModel):
    running: bool
    last_rapport: Optional[dict] = None
    ecosystem_path: str = ""


# ── Helpers ───────────────────────────────────────────────────────────────────

def _to_entree(f: dict, posts_base: Path | None = None) -> EntreeRegistre:
    """Convertit un dict registre N2 en EntreeRegistre (src.models).

    Si posts_base est défini et que chemin est relatif, résout le chemin
    complet = posts_base / chemin. Permet au moteur de trouver les fichiers
    quel que soit le répertoire de travail courant.
    """
    raw = f["chemin"]
    if posts_base and raw and not Path(raw).is_absolute():
        chemin = str(posts_base / raw)
    else:
        chemin = raw
    return EntreeRegistre(
        id=f["id"],
        type_fichier=f.get("type_fichier") or "",   # optionnel en N2, requis ici
        chemin=chemin,
        synchro_periodicite=f.get("synchro_periodicite", "quotidien"),
        derniere_synchro=f.get("derniere_synchro"),
        statut_dernier_synchro=f.get("statut_dernier_synchro"),
        genere_par_script=f.get("genere_par_script", False),
        owner_id=f.get("owner_role") or f.get("owner_id"),  # N2 utilise owner_role
    )


def _run_sync(ids: list[str], force: bool) -> dict:
    """Exécute la synchronisation et retourne le rapport."""
    global _sync_running, _last_rapport

    _sync_running = True
    debut = datetime.now()
    resultats = []

    try:
        fichiers_n2 = n2_load_registre()

        # Filtrer par IDs si spécifiés
        if ids:
            fichiers_n2 = [f for f in fichiers_n2 if f["id"] in ids]

        # Résoudre le répertoire Posts (affaire_dir / posts_dir) pour les chemins relatifs
        posts_base = get_posts_base_path()

        # Convertir en EntreeRegistre
        entrees = [_to_entree(f, posts_base) for f in fichiers_n2]

        # Trier par ordre de traitement (référentiels d'abord)
        def _ordre(e: EntreeRegistre) -> int:
            try:
                return ORDRE_TYPES.index(e.type_fichier)
            except ValueError:
                return 99

        entrees.sort(key=_ordre)

        # Synchroniser chaque fichier
        for entree in entrees:
            print(f"  [N2 Sync] {entree.id} ({entree.type_fichier or 'libre'})...")
            res = _sync_fichier(entree, force=force)
            resultats.append(res)
            for line in res["log"]:
                print(f"    {line}")

        # Mettre à jour le registre N2 avec statut + timestamp
        statuts = {r["id"]: r for r in resultats}
        for f in fichiers_n2:
            if f["id"] in statuts:
                r = statuts[f["id"]]
                f["derniere_synchro"] = r["timestamp"]
                f["statut_dernier_synchro"] = r["statut"]

        # Sauvegarder tous les fichiers du registre (pas seulement ceux syncés)
        all_fichiers = n2_load_registre()
        for f in all_fichiers:
            if f["id"] in statuts:
                r = statuts[f["id"]]
                f["derniere_synchro"] = r["timestamp"]
                f["statut_dernier_synchro"] = r["statut"]
        n2_save_registre(all_fichiers)

        fin = datetime.now()

        # Générer le rapport fichier
        rapport_path = _generer_rapport(resultats, debut, fin)

        # Historique
        try:
            History.save_run_history(resultats, debut, fin)
            History.save_store_snapshot()
            History.purge_old_files(max_runs=30, max_snapshots=30)
        except Exception as e:
            print(f"  [WARN] Historique non sauvegardé : {e}")

        nb_ok   = sum(1 for r in resultats if r["statut"] == "ok")
        nb_skip = sum(1 for r in resultats if r["statut"] == "skip_verrouille")
        nb_err  = sum(1 for r in resultats if r["statut"] == "erreur")

        rapport = {
            "debut": debut.isoformat(timespec="seconds"),
            "fin": fin.isoformat(timespec="seconds"),
            "nb_total": len(resultats),
            "nb_ok": nb_ok,
            "nb_skip": nb_skip,
            "nb_erreur": nb_err,
            "fichiers": resultats,
            "rapport_path": str(rapport_path),
            "ecosystem": str(get_active_config()),
        }

        _last_rapport = rapport
        print(f"\n[N2 Sync] Terminé : {nb_ok} OK / {nb_skip} skips / {nb_err} erreurs")
        return rapport

    finally:
        _sync_running = False


# ── Routes ────────────────────────────────────────────────────────────────────

@router.post("/run")
async def run_sync(body: SyncRequest):
    """Lance le moteur de synchronisation sur l'écosystème actif."""
    global _sync_running

    if _sync_running:
        raise HTTPException(409, "Une synchronisation est déjà en cours")

    # Vérifier qu'il y a des fichiers à synchroniser
    fichiers = n2_load_registre()
    if not fichiers:
        raise HTTPException(422, "Le registre est vide — aucun Post à synchroniser")

    # Exécuter dans un thread (ne bloque pas l'event loop FastAPI)
    try:
        rapport = await asyncio.to_thread(_run_sync, body.ids, body.force)
        return rapport
    except Exception as e:
        _sync_running = False
        raise HTTPException(500, f"Erreur lors de la synchronisation : {e}")


@router.get("/status", response_model=SyncStatusResponse)
def sync_status():
    """Retourne l'état courant du sync et le dernier rapport."""
    try:
        eco_path = str(get_active_config())
    except Exception:
        eco_path = ""
    return SyncStatusResponse(
        running=_sync_running,
        last_rapport=_last_rapport,
        ecosystem_path=eco_path,
    )


@router.get("/reports")
def list_reports():
    """Liste les rapports de synchronisation disponibles."""
    rapports_dir = ROOT / "output" / "rapports"
    if not rapports_dir.exists():
        return []
    files = sorted(rapports_dir.glob("Rapport_Synchro_*.json"), reverse=True)
    result = []
    for f in files[:20]:   # 20 derniers
        try:
            with open(f, encoding="utf-8") as fh:
                data = json.load(fh)
            result.append({
                "filename": f.name,
                "debut": data.get("debut"),
                "nb_ok": data.get("nb_ok", 0),
                "nb_erreur": data.get("nb_erreur", 0),
                "nb_total": data.get("nb_total", 0),
            })
        except Exception:
            result.append({"filename": f.name, "debut": None, "nb_ok": 0,
                           "nb_erreur": 0, "nb_total": 0})
    return result


@router.get("/reports/{filename}")
def get_report(filename: str):
    """Retourne le contenu d'un rapport de synchronisation."""
    if ".." in filename or "/" in filename or "\\" in filename:
        raise HTTPException(400, "Nom de fichier invalide")
    f = ROOT / "output" / "rapports" / filename
    if not f.exists():
        raise HTTPException(404, "Rapport introuvable")
    with open(f, encoding="utf-8") as fh:
        return json.load(fh)
