"""CLI — Système de Gestion des UO."""
import click

from pathlib import Path

from src.config_loader import load_uo_instances, load_registre
from src.generators.uo_generator import generate_uo_file
from src.generators.cockpit_generator import generate_cockpit
from src.generators.consolidation_generator import generate_consolidation
from src.generators.creator_generator import generate_creator
from src.sync import synchroniser, auditer_fichier


@click.group()
def cli():
    """Systeme de Gestion des UO — Secteur Ferroviaire."""


# ─── Génération ───────────────────────────────────────────────────────────────

@cli.command("generate-uo")
@click.option("--uo-id", required=True, help="ID de l'UO (ex: UO-001)")
def cmd_generate_uo(uo_id: str):
    """Genere le fichier Excel pour une UO specifique."""
    instances = load_uo_instances()
    target = next((uo for uo in instances if uo.id == uo_id), None)
    if not target:
        click.echo(f"[ERREUR] UO '{uo_id}' introuvable. IDs: {[uo.id for uo in instances]}")
        raise SystemExit(1)
    path = generate_uo_file(target)
    click.echo(f"[OK] {path}")


@cli.command("generate-all-uo")
def cmd_generate_all_uo():
    """Genere les fichiers Excel pour toutes les UO."""
    instances = load_uo_instances()
    for uo in instances:
        path = generate_uo_file(uo)
        click.echo(f"[OK] {uo.id} ({uo.statut.value}{'*' if uo.degrade else ''}) -> {path.name}")
    click.echo(f"\n{len(instances)} fichier(s) UO genere(s).")


@cli.command("generate-cockpit")
@click.option("--engineer", required=True, help="Nom de l'ingenieur")
def cmd_generate_cockpit(engineer: str):
    """Genere le cockpit pour un ingenieur specifique."""
    instances = load_uo_instances()
    engineers = set(uo.engineer_name for uo in instances)
    if engineer not in engineers:
        click.echo(f"[ERREUR] Ingenieur '{engineer}' introuvable. Disponibles: {sorted(engineers)}")
        raise SystemExit(1)
    path = generate_cockpit(engineer, instances)
    click.echo(f"[OK] {path}")


@cli.command("generate-all-cockpits")
def cmd_generate_all_cockpits():
    """Genere les cockpits pour tous les ingenieurs."""
    instances = load_uo_instances()
    engineers = sorted(set(uo.engineer_name for uo in instances))
    for eng in engineers:
        path = generate_cockpit(eng, instances)
        click.echo(f"[OK] {eng} -> {path.name}")
    click.echo(f"\n{len(engineers)} cockpit(s) genere(s).")


@cli.command("generate-consolidation")
def cmd_generate_consolidation():
    """Genere le fichier de consolidation centrale."""
    instances = load_uo_instances()
    path = generate_consolidation(instances)
    click.echo(f"[OK] {path}")


@cli.command("generate-all")
def cmd_generate_all():
    """Genere tous les fichiers (UO + cockpits + consolidation)."""
    click.echo(">> Generation de tous les fichiers...\n")
    instances = load_uo_instances()

    click.echo("Fichiers UO :")
    for uo in instances:
        path = generate_uo_file(uo)
        degrade = "*" if uo.degrade else ""
        click.echo(f"  [OK] {uo.id}{degrade} ({uo.statut.value}) -> {path.name}")

    click.echo("\nCockpits :")
    engineers = sorted(set(uo.engineer_name for uo in instances))
    for eng in engineers:
        path = generate_cockpit(eng, instances)
        click.echo(f"  [OK] {eng} -> {path.name}")

    click.echo("\nConsolidation :")
    path = generate_consolidation(instances)
    click.echo(f"  [OK] {path.name}")

    click.echo(f"\nTermine -- {len(instances)} UO, {len(engineers)} cockpit(s), 1 consolidation.")


# ─── Synchronisation ──────────────────────────────────────────────────────────

@cli.command("sync")
@click.option("--id", "ids", multiple=True, help="IDs specifiques a synchroniser (repetable)")
@click.option("--type", "types", multiple=True,
              help="Types de fichiers a synchroniser (repetable): uo_instance, referentiel_uo, cockpit...")
@click.option("--force", is_flag=True, default=False,
              help="Ignorer la verification de verrouillage")
def cmd_sync(ids, types, force):
    """Lance une synchronisation des passerelles."""
    click.echo(">> Synchronisation en cours...\n")
    rapport = synchroniser(
        ids=list(ids) or None,
        types=list(types) or None,
        force=force,
    )
    click.echo(f"\nRapport genere : {rapport}")


@cli.command("sync-uo")
@click.argument("uo_id")
def cmd_sync_uo(uo_id: str):
    """Synchronise une UO specifique (ex: python main.py sync-uo UO-001)."""
    click.echo(f">> Synchro {uo_id}...")
    rapport = synchroniser(ids=[uo_id])
    click.echo(f"Rapport : {rapport}")


# ─── Registre & Onboarding ────────────────────────────────────────────────────

@cli.command("create-creator")
def cmd_create_creator():
    """Genere Creator.xlsx — console d'administration de l'ecosysteme."""
    path = generate_creator()
    click.echo(f"[OK] Creator genere : {path}")


@cli.command("list-registre")
def cmd_list_registre():
    """Affiche le contenu du registre des fichiers."""
    entrees = load_registre()
    click.echo(f"\n{'ID':<20} {'TYPE':<22} {'PERIODICITE':<14} {'STATUT':<12} CHEMIN")
    click.echo("-" * 100)
    for e in entrees:
        statut = e.statut_dernier_synchro or "jamais"
        click.echo(f"{e.id:<20} {e.type_fichier:<22} {e.synchro_periodicite:<14} {statut:<12} {e.chemin}")


@cli.command("onboard")
@click.argument("chemin")
def cmd_onboard(chemin: str):
    """
    Audit d'onboarding d'un fichier Excel cree manuellement.
    Verifie la structure et genere un template _Passerelle si absent.
    """
    click.echo(f">> Audit onboarding : {chemin}\n")
    rapport = auditer_fichier(chemin)
    click.echo(f"Resultat : {'OK' if rapport['ok'] else 'PROBLEMES DETECTES'}")
    for alerte in rapport.get("alertes", []):
        click.echo(f"  [!] {alerte}")
    for action in rapport.get("actions", []):
        click.echo(f"  [>] {action}")
    feuilles = rapport.get("feuilles_trouvees", [])
    if feuilles:
        click.echo(f"  Feuilles trouvees : {', '.join(feuilles)}")


# ─── Écosystème & Parser ──────────────────────────────────────────────────────

@cli.command("parse-file")
@click.argument("chemin")
@click.option("--enrich", is_flag=True, default=False,
              help="Enrichit l'ecosystem schema depuis l'AST parsé")
def cmd_parse_file(chemin: str, enrich: bool):
    """
    Parse la feuille _Passerelle d'un fichier Excel et affiche l'AST.
    Ex: python main.py parse-file output/UOs/UO-001_spec_fonctionnelle_climatisation.xlsx
    """
    from src.parser import parse_file, ast_summary, enrich_ecosystem
    path = Path(chemin)
    if not path.exists():
        click.echo(f"[ERREUR] Fichier introuvable : {chemin}")
        raise SystemExit(1)

    ast = parse_file(path)
    if ast is None:
        click.echo(f"[ERREUR] Aucune feuille _Passerelle trouvee dans {path.name}")
        raise SystemExit(1)

    click.echo(f"\n>> Parser : {path.name}\n")
    click.echo(ast_summary(ast))

    if enrich:
        nb_tables, nb_vars = enrich_ecosystem(ast)
        click.echo(f"\n[Ecosystem] +{nb_tables} table(s), +{nb_vars} variable(s) enregistrees.")
    elif ast.pushes:
        click.echo(f"\n(Utilisez --enrich pour alimenter l'ecosystem depuis ce fichier.)")


@cli.command("enrich-ecosystem")
@click.option("--dir", "scan_dir", default=None,
              help="Repertoire a scanner (defaut: output/)")
def cmd_enrich_ecosystem(scan_dir: str):
    """Scanne tous les fichiers Excel du repertoire et enrichit l'ecosystem."""
    from src.parser import parse_file, enrich_ecosystem
    from src import ecosystem as eco

    base = Path(scan_dir) if scan_dir else Path("output")
    xlsx_files = list(base.rglob("*.xlsx"))
    if not xlsx_files:
        click.echo(f"[WARN] Aucun fichier .xlsx dans {base}")
        return

    click.echo(f">> Scan de {len(xlsx_files)} fichier(s) dans {base}\n")
    total_t, total_v = 0, 0
    for f in xlsx_files:
        ast = parse_file(f)
        if ast is None:
            continue
        nb_t, nb_v = enrich_ecosystem(ast)
        if nb_t or nb_v:
            click.echo(f"  {f.name} : +{nb_t} table(s), +{nb_v} variable(s)")
        total_t += nb_t
        total_v += nb_v

    s = eco.summary()
    click.echo(f"\nEcosystem : {s['nb_tables']} table(s), {s['nb_variables']} variable(s) au total.")
    click.echo(f"Session   : +{total_t} table(s), +{total_v} variable(s) nouvelles.")


if __name__ == "__main__":
    cli()
