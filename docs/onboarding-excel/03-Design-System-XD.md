# Design System Excel — `src/xl_design.py`

Tout ce qui touche à la mise en forme d'un fichier Excel généré par ExoSync
passe par la classe `XD` de `src/xl_design.py`. **Règle non négociable : zéro
style Excel écrit en inline dans un générateur.** Si tu as besoin d'un style
qui n'existe pas encore dans `XD`, ajoute-le dans `xl_design.py`, ne le code
pas en dur dans ton générateur.

## Pourquoi cette règle

Elle garantit que tout fichier généré (UO, cockpit, dashboard, et toute vue
future comme le dashboard client) partage la même palette, la même
typographie, les mêmes composants visuels — sans effort de synchronisation
manuelle entre générateurs.

## Import et usage

```python
from src.xl_design import XD

XD.banner(ws, "activites", "Mes Activités", subtitle="UO L09U1-CFL2400-CLIM", n_cols=10)
XD.table_header(ws, row=5, headers=["ID", "Libellé", "Heures"], key="activites", col_start=2)
XD.named_table(ws, display_name="tbl_activites", ref="B5:D20", key="activites")
```

## Les 11 familles d'onglets

Chaque onglet appartient à une famille (`key`), qui détermine sa couleur
(3 tons dérivés : bannière foncée, en-tête moyen, lignes alternées claires) et
son glyphe. Familles disponibles : `general`, `dashboard`, `description`,
`planning`, `donnees_entree`, `activites`, `livrables`, `oil`, `kpi`, `orga`,
`manifeste`. Voir `XD.SHEETS` dans `src/xl_design.py` pour la palette exacte
de chaque famille — ne pas inventer de nouvelle couleur ailleurs, choisir la
famille la plus proche du sens de l'onglet.

## Composants principaux

| Méthode | Rôle |
|---|---|
| `XD.banner(ws, key, title, subtitle, se, n_cols)` | Bannière 1 ligne en tête d'onglet, pose aussi `tabColor` |
| `XD.table_header(ws, row, headers, key, col_start)` | En-tête de tableau coloré au ton moyen de la famille |
| `XD.data_row(ws, row, i, col_start, col_end, key)` | Ligne de données avec alternance de fond |
| `XD.named_table(ws, display_name, ref, key)` | Table Excel nommée (nécessaire pour `GET_TABLE`/`COLLECT` côté MXL) avec en-tête coloré |
| `XD.statut_cf(ws, rng)` | Mise en forme conditionnelle badges de statut (TERMINEE, EN_COURS, A_FAIRE...) |
| `XD.criticite_cf(ws, rng)` | Mise en forme conditionnelle badges de criticité OIL |
| `XD.traffic_light(ws, row, col, value, thresholds)` | Cellule feu rouge/ambre/vert selon un seuil |
| `XD.section_box(ws, title, r1, c1, r2, c2, key)` | Bande de titre + cadre, pour les onglets clé-valeur |
| `XD.health_spine(ws, key, header_row, row_start, row_end, status_col, spine_col)` | Colonne santé (voir ci-dessous) |

## La colonne "spine santé"

Composant contextuel pour les onglets à lignes "vivantes" (Activités du côté
UO, Mes UOs du côté cockpit, Synthèse du côté dashboard) : une colonne A fine
(largeur 2.5), colorée en direct par mise en forme conditionnelle selon une
colonne statut, qui donne un repère visuel immédiat sans avoir à lire le
détail de chaque ligne.

**Règle absolue : la colonne spine ne doit jamais faire partie du `ref` d'une
table Excel nommée.** Si elle l'est, `GET_TABLE`/`COLLECT` côté MXL lit une
colonne en trop et casse l'agrégation. Conséquence pratique : quand une table
a une spine, le `ref` de la table nommée commence en colonne B, pas en
colonne A.

```python
# Bon : la table commence en B, la spine (col A) est hors du ref
XD.named_table(ws, "tbl_activites", "B4:J20", "activites")
XD.health_spine(ws, "activites", header_row=4, row_start=5, row_end=20, status_col=7)
```

## Règles universelles à respecter sur tout nouvel onglet

1. Bannière en tête (`XD.banner`)
2. Police Segoe UI partout (gérée automatiquement par `XD.fnt`)
3. En-tête de tableau au ton de la famille (`XD.table_header`)
4. Lignes alternées (`XD.data_row` ou géré par le style de table nommée)
5. Jaune de saisie `XD.INPUT` (`FFF2CC`) là où l'utilisateur doit saisir
6. Couleur d'onglet (`tabColor`) posée automatiquement par `XD.banner`
