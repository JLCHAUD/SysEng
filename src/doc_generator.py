"""
ExoSync — Générateur de documentation (M13)
============================================
Génère une documentation HTML statique de l'écosystème depuis l'Exomap.

Fonctions publiques :
    generate_html_doc(output_dir, ecosystem_path) → Path
"""
import json
from datetime import datetime
from pathlib import Path
from typing import Optional

ROOT = Path(__file__).parent.parent


# ─── Templates HTML ───────────────────────────────────────────────────────────

_HTML_TEMPLATE = """\
<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>ExoSync — Documentation de l ecosysteme</title>
  <style>
    body {{ font-family: system-ui, sans-serif; margin: 0; padding: 0;
            background: #f8fafc; color: #1e293b; }}
    header {{ background: #1e40af; color: white; padding: 1.5rem 2rem; }}
    header h1 {{ margin: 0; font-size: 1.5rem; }}
    header p  {{ margin: 0.25rem 0 0; opacity: 0.8; font-size: 0.9rem; }}
    main {{ max-width: 1100px; margin: 2rem auto; padding: 0 1.5rem; }}
    h2 {{ color: #1e40af; border-bottom: 2px solid #e2e8f0;
          padding-bottom: 0.5rem; margin-top: 2rem; }}
    h3 {{ color: #334155; margin-top: 1.5rem; }}
    .stats {{ display: grid; grid-template-columns: repeat(4, 1fr);
              gap: 1rem; margin: 1.5rem 0; }}
    .stat-card {{ background: white; border-radius: 8px; padding: 1rem 1.5rem;
                  box-shadow: 0 1px 3px rgba(0,0,0,0.1); text-align: center; }}
    .stat-card .num {{ font-size: 2rem; font-weight: bold; color: #1e40af; }}
    .stat-card .lbl {{ font-size: 0.85rem; color: #64748b; }}
    table {{ width: 100%; border-collapse: collapse; background: white;
             border-radius: 8px; overflow: hidden;
             box-shadow: 0 1px 3px rgba(0,0,0,0.1); margin: 1rem 0; }}
    th {{ background: #1e40af; color: white; padding: 0.6rem 1rem;
          text-align: left; font-size: 0.85rem; }}
    td {{ padding: 0.5rem 1rem; border-top: 1px solid #e2e8f0;
          font-size: 0.9rem; }}
    tr:hover td {{ background: #f1f5f9; }}
    .badge {{ display: inline-block; padding: 0.1rem 0.5rem;
              border-radius: 99px; font-size: 0.75rem; font-weight: bold; }}
    .badge-ok      {{ background: #dcfce7; color: #166534; }}
    .badge-error   {{ background: #fee2e2; color: #991b1b; }}
    .badge-unknown {{ background: #f1f5f9; color: #64748b; }}
    .badge-pull    {{ background: #dbeafe; color: #1e40af; }}
    .badge-push    {{ background: #fef3c7; color: #92400e; }}
    .warn-box {{ background: #fef3c7; border-left: 4px solid #f59e0b;
                 padding: 0.75rem 1rem; border-radius: 4px;
                 margin: 1rem 0; font-size: 0.9rem; }}
    .ok-box   {{ background: #dcfce7; border-left: 4px solid #16a34a;
                 padding: 0.75rem 1rem; border-radius: 4px;
                 margin: 1rem 0; font-size: 0.9rem; }}
    pre {{ background: #1e293b; color: #e2e8f0; padding: 1rem;
           border-radius: 6px; overflow-x: auto; font-size: 0.85rem; }}
    footer {{ text-align: center; color: #94a3b8; font-size: 0.8rem;
              padding: 2rem; margin-top: 3rem; }}
  </style>
</head>
<body>
<header>
  <h1>ExoSync — Documentation de l ecosysteme</h1>
  <p>Generee le {generated_at} &bull; {nb_files} fichiers &bull; {nb_edges} arcs</p>
</header>
<main>

  <h2>Vue d ensemble</h2>
  <div class="stats">
    <div class="stat-card">
      <div class="num">{nb_files}</div>
      <div class="lbl">Fichiers</div>
    </div>
    <div class="stat-card">
      <div class="num">{nb_edges}</div>
      <div class="lbl">Arcs (PULL/PUSH)</div>
    </div>
    <div class="stat-card">
      <div class="num">{nb_push}</div>
      <div class="lbl">PUSH</div>
    </div>
    <div class="stat-card">
      <div class="num">{nb_pull}</div>
      <div class="lbl">PULL</div>
    </div>
  </div>

  {warnings_section}

  <h2>Fichiers de l ecosysteme</h2>
  <table>
    <thead>
      <tr><th>ID</th><th>Type</th><th>Chemin</th><th>Derniere sync</th><th>Statut</th></tr>
    </thead>
    <tbody>
      {files_rows}
    </tbody>
  </table>

  <h2>Graphe de dependances</h2>
  <table>
    <thead>
      <tr><th>Type</th><th>Source</th><th>Destination</th><th>Mode</th></tr>
    </thead>
    <tbody>
      {edges_rows}
    </tbody>
  </table>

  <h2>Lineage (texte)</h2>
  <pre>{lineage_text}</pre>

  <h2>Donnees brutes (JSON)</h2>
  <pre>{raw_json}</pre>

</main>
<footer>
  Genere par ExoSync &mdash; {generated_at}
</footer>
</body>
</html>
"""


def _badge(val: str, mapping: dict) -> str:
    cls = mapping.get(val, "badge-unknown")
    return f'<span class="badge {cls}">{val}</span>'


def generate_html_doc(
    output_dir: Optional[Path] = None,
    ecosystem_path: Optional[Path] = None,
) -> Path:
    """
    Génère un fichier HTML documentant l'écosystème.

    Args:
        output_dir      : dossier de sortie (défaut : output/doc/)
        ecosystem_path  : chemin de l'Exomap (défaut : output/ecosystem.json)

    Returns:
        Chemin du fichier HTML généré.
    """
    from src.ecosystem import load as eco_load, lineage_text, ECOSYSTEM_PATH, check_consistency

    eco_path = ecosystem_path or ECOSYSTEM_PATH
    schema   = eco_load(eco_path)
    out_dir  = output_dir or (ROOT / "output" / "doc")
    out_dir.mkdir(parents=True, exist_ok=True)

    ts = datetime.now()
    generated_at = ts.strftime("%Y-%m-%d %H:%M")

    # Statistiques
    nb_files = len(schema.files)
    nb_edges = len(schema.edges)
    nb_push  = sum(1 for e in schema.edges if e.edge_type == "PUSH")
    nb_pull  = sum(1 for e in schema.edges if e.edge_type == "PULL")

    # Avertissements
    warnings = check_consistency(eco_path)
    if warnings:
        items = "".join(
            f'<div class="warn-box"><b>[{w.code}]</b> {w.message}'
            f'{(" — " + w.details) if w.details else ""}</div>'
            for w in warnings
        )
        warnings_section = f"<h2>Avertissements ({len(warnings)})</h2>{items}"
    else:
        warnings_section = '<div class="ok-box">Aucune incoherence detectee.</div>'

    # Fichiers
    status_map = {"ok": "badge-ok", "error": "badge-error"}
    files_rows = ""
    for fid, fr in schema.files.items():
        badge = _badge(fr.status, status_map)
        sync  = fr.last_sync or "jamais"
        files_rows += (
            f"<tr><td><b>{fid}</b></td><td>{fr.file_type}</td>"
            f"<td>{fr.path}</td><td>{sync}</td><td>{badge}</td></tr>\n"
        )

    # Arcs
    type_map = {"PULL": "badge-pull", "PUSH": "badge-push"}
    edges_rows = ""
    for e in schema.edges:
        badge = _badge(e.edge_type, type_map)
        edges_rows += (
            f"<tr><td>{badge}</td><td>{e.from_node}</td>"
            f"<td>{e.to_node}</td><td>{e.mode or ''}</td></tr>\n"
        )

    # Lineage texte
    lt = lineage_text(path=eco_path)

    # JSON brut (tronqué si trop grand)
    raw = {
        "version":   schema.version,
        "last_scan": schema.last_scan,
        "files":     {fid: vars(f) for fid, f in schema.files.items()},
        "edges":     [vars(e) for e in schema.edges],
    }
    raw_json_str = json.dumps(raw, ensure_ascii=False, indent=2, default=str)
    if len(raw_json_str) > 20000:
        raw_json_str = raw_json_str[:20000] + "\n... (tronque)"

    html = _HTML_TEMPLATE.format(
        generated_at     = generated_at,
        nb_files         = nb_files,
        nb_edges         = nb_edges,
        nb_push          = nb_push,
        nb_pull          = nb_pull,
        warnings_section = warnings_section,
        files_rows       = files_rows or "<tr><td colspan='5'>Aucun fichier</td></tr>",
        edges_rows       = edges_rows or "<tr><td colspan='4'>Aucun arc</td></tr>",
        lineage_text     = lt.replace("<", "&lt;").replace(">", "&gt;"),
        raw_json         = raw_json_str.replace("<", "&lt;").replace(">", "&gt;"),
    )

    out_file = out_dir / f"ecosystem_{ts.strftime('%Y%m%d_%H%M%S')}.html"
    with open(out_file, "w", encoding="utf-8") as f:
        f.write(html)

    return out_file
