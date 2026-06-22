# Design System Excel — `src/xl_design.py`

- **Date** : 2026-06-22
- **Conversation** : CONV-13 — Design System Excel
- **Branche cible** : `feature/design-system`
- **Statut** : spec validée, prête pour plan d'implémentation
- **Référence** : `projet_TrainSystem/design_b.py` (prototype conservé)

---

## 1. Objectif & périmètre

Centraliser toute la charte graphique Excel d'ExoSync dans un module unique
`src/xl_design.py` (classe `XD`). Tous les générateurs l'importent ; **aucun
style n'est défini inline**. Résultat : tout fichier généré par ExoSync partage
la même palette, la même typographie, les mêmes composants — sans effort.

Périmètre : le module + la migration des 4 générateurs existants. `design_b.py`
et `Design_B_Cockpit.xlsx` restent comme prototype/référence (non supprimés).

---

## 1bis. Charte générique — invariants vs contextuels

`xl_design` n'est **pas** « le style des UO » : c'est une **charte générique**
que **tout** fichier ExoSync applique — UO instance, cockpit ingénieur, dashboard
métier, et toute **copie / vue future** (dashboard client, vue projet…). Les
grandes règles se retrouvent **partout** ; seuls les composants contextuels
varient selon ce qui est applicable à chaque onglet.

### Règles universelles (sur TOUT fichier, tout onglet où ça a un sens)

1. **Bannière en tête** (1 ligne, glyphe + couleur d'onglet).
2. **Police Segoe UI** partout.
3. **Palette par famille + dérivation 3 tons** (§2-3).
4. **En-tête de tableau au ton de l'onglet**.
5. **Lignes alternées** (accent / blanc).
6. **Jaune de saisie `FFF2CC`** là où l'utilisateur saisit.
7. **Couleur d'onglet** (`tabColor`) par feuille.
8. **Onglet `_Manifeste`** gris technique.

### Composants contextuels (convoqués quand c'est applicable)

- **Colonne spine santé** — onglets à lignes « vivantes » (Activités, Mes UOs,
  Synthèse…), pas sur les onglets clé-valeur ou texte.
- **AutoFilter** — tables nommées.
- **Cartes KPI** — Dashboard, onglet KPI.
- **Feu tricolore** — KPI, Alertes.
- **Badges statut / criticité** — onglets à colonne statut.
- **Section box** — onglets clé-valeur (General, KPI).

### Mécanique

Chaque générateur (présent **et futur**) `from src.xl_design import XD`, déclare
la **famille** de chaque onglet (`key`), et les invariants s'appliquent seuls.
Les composants contextuels sont des appels explicites à la carte. Conséquence :
créer un nouveau type de fichier (ex. dashboard client) **hérite gratuitement**
du look ExoSync — il n'a qu'à composer les composants pertinents.

---

## 2. Principe directeur — identité couleur par onglet

**Chaque type d'onglet possède une famille de couleur**, déclinée en 3 tons.
Quand l'utilisateur est dans un onglet, **bannière + en-tête de tableau + fonds
partagent la même teinte** → il *sait* visuellement où il est.

Les couleurs sont ordonnées en **parcours chromatique continu** suivant l'ordre
logique des onglets : marine (pilotage) → cyan/teal/vert (cadrage + flux de
travail) → rouge (risque) → violet (mesure) → gris (technique). Le seul saut
volontaire est **Livrables (vert) → OIL (rouge)** : la rupture fait que l'alerte
saute aux yeux.

La **fonction porte la couleur, pas le nom de l'onglet** → réutilisable entre
types de fichiers (un dashboard réutilise le marine de *General*, ses alertes le
rouge de *OIL*, etc.).

### Règle de dérivation (3 tons par famille)

| Élément | Ton | Rôle |
|---|---|---|
| **Bannière** (ligne 1) | foncé | identité forte, texte blanc + glyphe |
| **Onglet (tabColor) + en-tête de tableau** | moyen | « tu es ici » |
| **Lignes alternées / cartes / sections** | très clair | respiration |

**Exception transversale** : le jaune de saisie `FFF2CC` (zones que l'ingénieur
remplit) **ne dérive jamais** — il doit rester reconnaissable à l'identique sur
tous les onglets et tous les fichiers.

---

## 3. Palette complète — 11 familles

Ordre = ordre des onglets (parcours continu). Contraste renforcé : deux onglets
adjacents ne partagent jamais à la fois teinte **et** luminosité.

| # | Onglet | Bloc | Bannière (foncé) | Onglet + en-tête (moyen) | Accent (clair) | Glyphe |
|---|---|---|---|---|---|---|
| 1 | General | Pilotage | `08335E` | `0C447C` | `E6F1FB` | ⬢ |
| 2 | Dashboard | Pilotage | `0E4474` | `1763A8` | `E3EFFA` | ◉ |
| 3 | Description / Besoin | Cadrage | `1C5E92` | `2E86C8` | `E7F2FB` | ✎ |
| 4 | Planning | Cadrage | `074E60` | `0A6E88` | `DEEFF3` | ◷ |
| 5 | Données d'entrée | Flux de travail | `0A6149` | `0F8A66` | `E1F5EE` | ⤓ |
| 6 | Activités | Flux de travail | `084434` | `0C5E49` | `E1F5EE` | ✔ |
| 7 | Livrables | Flux de travail | `386114` | `4F8A1E` | `EBF3DE` | ▣ |
| 8 | OIL — points ouverts | Risque | `791F1F` | `A32D2D` | `FCEBEB` | ⚑ |
| 9 | KPI | Mesure | `3C3489` | `534AB7` | `EEEDFE` | ▲ |
| 10 | Orga | Technique | `4D4C47` | `6B6A64` | `F1EFE8` | ❖ |
| 11 | _Manifeste | Technique | `1C1C1A` | `2C2C2A` | `F1EFE8` | ⚙ |

Le texte sur bannière et en-tête est **blanc** (`FFFFFF`) — tous les tons
« foncé » et « moyen » offrent un contraste suffisant.

### Couleurs de statut (sémantiques, transversales)

Reprises de `design_b.py`, identiques sur tous les onglets/fichiers :

| Famille | Fond | Texte |
|---|---|---|
| Vert (TERMINEE / VALIDE / BASSE / CLOS / OK) | `EAF3DE` | `27500A` |
| Bleu (EN_COURS / LIVRE) | `E6F1FB` | `0C447C` |
| Ambre (STAND_BY / MOYENNE / échéance proche) | `FAEEDA` | `854F0B` |
| Rouge (HAUTE / OUVERT / dérive) | `FCEBEB` | `791F1F` |
| Gris (A_FAIRE / EN_ATTENTE) | `F1EFE8` | `5F5E5A` |

**Feu tricolore** (`traffic_light`) : `< 50%` → fond rouge `FCEBEB` · `50–80%` →
ambre `FAEEDA` · `> 80%` → vert `EAF3DE`. Seuils par défaut `(0.5, 0.8)`.

### Jaune de saisie

`INPUT = "FFF2CC"` — zones remplies par l'ingénieur. Universel.

---

## 4. Typographie

- **Police unique** : `Segoe UI` partout (constante `FONT_FAMILY`).
- Texte courant : 10 pt, couleur `2C2C2A`.
- En-têtes : 10 pt gras, blanc.
- Titre bannière : 14 pt gras, blanc. Sous-titre : 12 pt. Mention SE : 9–10 pt.

---

## 5. Bannière — anatomie (sans navigation)

Une **seule ligne** (~30 px), 4 zones de gauche à droite :

```
[glyphe]  Titre (UO L09U1 — Activités)   projet · système          Ingénieur Système / nom
```

- Le **glyphe monochrome** (même couleur que le texte blanc) + la couleur de
  bannière donnent l'identité immédiate de l'onglet (« mémo technique » :
  où suis-je ?).
- Les **boutons de navigation par hyperlien sont retirés** (présents dans
  l'ancien `banner_B`). On pourra réintroduire une nav plus tard si besoin.
- `tabColor` de la feuille = ton « moyen » de la famille (cohérence onglet ↔
  bannière).

---

## 5bis. Colonne « spine » santé (gutter de gauche)

Une **colonne fine en tout début de tableau** (colonne A, largeur ~2,5) qui donne
un **scan visuel immédiat** de l'état de chaque ligne — « ce qui est fait / à
faire / en alerte » d'un coup d'œil.

- **Colonne dédiée**, *distincte de l'ID* (l'ID a une largeur variable, il garde
  sa propre colonne).
- **En-tête de la spine** = ton bannière de l'onglet (ferme proprement le tableau
  à gauche).
- **Hors de la table nommée** : la spine n'est **pas** incluse dans la plage de la
  table Excel nommée (`tbl_activites`…), pour que les colonnes lues par
  `GET_TABLE` / `COLLECT` restent inchangées. La table nommée démarre donc en
  **colonne B**, la spine occupe la **colonne A**.
- **Pilotage = mise en forme conditionnelle** (pas de script) : la couleur dérive
  du **statut + avancement** de la ligne et se recolore **live** quand
  l'ingénieur saisit. Le moteur de scan pourra surcharger plus tard (valeur
  0-100, `W`, alerte…).

### Mapping santé → couleur de la spine

| Condition (ligne) | Couleur | Sens |
|---|---|---|
| statut = TERMINEE / VALIDE (ou avancement = 100) | `3B6D11` vert | terminé |
| en cours, dans les clous (avancement > 0, pas d'alerte) | `0F8A66` teal | en bonne voie |
| échéance proche / STAND_BY / à surveiller | `EF9F27` ambre | à surveiller |
| dérive heures / criticité HAUTE / OUVERT critique | `A32D2D` rouge | alerte |
| statut = A_FAIRE / EN_ATTENTE (avancement = 0) | `6B6A64` gris | à faire |

(Plus tard : petits glyphes monochromes dans la spine — balle chez qui, santé.)

---

## 6. API du module `XD`

```python
# src/xl_design.py
from dataclasses import dataclass

@dataclass(frozen=True)
class SheetStyle:
    banner: str      # ton foncé
    header: str      # ton moyen (= tabColor)
    accent: str      # ton clair
    glyph: str       # glyphe monochrome
    # tab = header

class XD:
    FONT_FAMILY = "Segoe UI"

    # ── Palette (constantes) ───────────────────────────────
    NAVY_D = "0C447C"; NAVY = "185FA5"; BLUE = "378ADD"; BLUE_L = "E6F1FB"
    GREEN = "639922"; GREEN_L = "EAF3DE"; GREEN_D = "27500A"
    AMBER = "EF9F27"; AMBER_L = "FAEEDA"; AMBER_D = "854F0B"
    RED = "E24B4A"; RED_L = "FCEBEB"; RED_D = "791F1F"
    GREY_D = "5F5E5A"; GREY_L = "F5F4F0"; GREY_B = "D3D1C7"
    WHITE = "FFFFFF"; INPUT = "FFF2CC"

    # ── Registre central des onglets ───────────────────────
    SHEETS: dict[str, SheetStyle]    # "general", "dashboard", "description",
                                     # "planning", "donnees_entree", "activites",
                                     # "livrables", "oil", "kpi", "orga", "manifeste"

    # ── Primitives ─────────────────────────────────────────
    @staticmethod
    def fnt(size=10, bold=False, color="2C2C2A", italic=False) -> Font
    @staticmethod
    def fill(hex_color: str) -> PatternFill
    @staticmethod
    def center() -> Alignment      # centré + wrap
    @staticmethod
    def left() -> Alignment        # gauche + wrap
    HAIR: Border                   # bordure fine GREY_B 4 côtés

    # ── Accès au style d'un onglet ─────────────────────────
    @classmethod
    def sheet(cls, key: str) -> SheetStyle

    # ── Composants ─────────────────────────────────────────
    @classmethod
    def banner(cls, ws, key: str, title: str, subtitle: str = "",
               se: str = "", n_cols: int = 10, height: int = 30)
    # pose tabColor=header, remplit la ligne 1 en banner, écrit glyphe+titre
    # (blanc), sous-titre projet·système, mention SE à droite.

    @classmethod
    def table_header(cls, ws, row: int, headers: list[str], key: str)
    # en-tête coloré au ton header de l'onglet, texte blanc, bordure HAIR.

    @classmethod
    def data_row(cls, ws, row: int, i: int, col_start: int, col_end: int, key: str)
    # i pair → blanc, i impair → accent de l'onglet. Bordure HAIR.

    @staticmethod
    def card_border(ws, r1, c1, r2, c2, accent_left: str | None = None)

    @classmethod
    def named_table(cls, ws, display_name: str, ref: str, key: str)
    # Table Excel nommée avec STYLE CLAIR (Light) + en-tête coloré manuellement
    # au ton de l'onglet. AutoFilter natif actif (filtrage par colonne). Voir §7.

    @classmethod
    def health_spine(cls, ws, key: str, header_row: int,
                     row_start: int, row_end: int,
                     status_col: int, pct_col: int, spine_col: int = 1)
    # Colonne A fine + en-tête au ton bannière ; pose les règles de mise en forme
    # conditionnelle santé (§5bis) lues sur status_col / pct_col.

    @classmethod
    def traffic_light(cls, ws, row: int, col: int, value: float,
                      thresholds=(0.5, 0.8))

    @staticmethod
    def statut_cf(ws, rng: str)      # badges statut (TERMINEE/EN_COURS/...)
    @staticmethod
    def criticite_cf(ws, rng: str)   # badges criticité (HAUTE/MOYENNE/BASSE)

    # ── Conventions onglets ────────────────────────────────
    @classmethod
    def tab_colors(cls) -> dict[str, str]   # dérivé de SHEETS (ton header)
```

---

## 7. Tables nommées Excel — style clair + en-tête manuel

**Décision (corrige le CONV-13 qui disait `TableStyleMedium2`)** :

`TableStyleMedium2` impose un en-tête **bleu Excel** à toutes les tables, ce qui
**écrase la couleur d'onglet** et casse l'identité visuelle. Pour avoir des
en-têtes **aux couleurs de chaque onglet**, on adopte la technique déjà utilisée
par `design_b.py` :

1. **Style de table clair** sans en-tête imposé (`TableStyleLight15`,
   `showRowStripes=True`).
2. **Coloration manuelle de l'en-tête** au ton « moyen » de l'onglet (texte
   blanc), via `table_header` ou directement dans `named_table`.

Les tables restent de **vraies tables Excel nommées** (indispensable pour
`GET_TABLE` / `COLLECT` du moteur MXL) — seul le style change.

**Frontière `named_table` vs `data_row`** (pour éviter le double striping) :

- **`named_table`** (Activités, Livrables, OIL, Données, `tbl_mes_uos`…) : on
  laisse les **rayures du style Excel clair** gérer l'alternance des lignes
  (neutres). On ne colore **que l'en-tête** manuellement. On n'appelle **pas**
  `data_row` sur ces plages.
- **`data_row`** : réservé aux **plages NON-table** construites à la main
  (listes de cockpit/dashboard, sections) où l'on veut l'alternance
  blanc / accent de la famille.

### Filtrage rapide — AutoFilter natif (pas de slicers)

Le filtrage rapide s'appuie sur l'**AutoFilter natif** des tables Excel nommées
(déroulant par colonne, gratuit). **Décision : pas de slicers/segments en v1** —
`openpyxl` ne sait pas les générer (injection XML brut, fragile). La **colonne
spine** (§5bis) couvre l'essentiel du besoin de scan visuel rapide. Un éventuel
slicer fera l'objet d'un **spike séparé** post-v1 si le besoin persiste.

---

## 8. Réutilisation cross-fichiers

Les onglets des cockpits et dashboards réutilisent les familles par **fonction** :

| Fichier | Onglet | Famille réutilisée |
|---|---|---|
| Cockpit ingénieur | Mes UOs | `general` (marine) |
| Cockpit ingénieur | Agenda | `planning` (cyan) |
| Cockpit ingénieur | _Manifeste | `manifeste` (gris) |
| Dashboard métier | Synthèse | `dashboard` (marine moyen) |
| Dashboard métier | Vue Synthèse | `dashboard` |
| Dashboard métier | Par Ingénieur | `activites` (teal) |
| Dashboard métier | Alertes | `oil` (rouge) |
| Dashboard métier | _Manifeste | `manifeste` |

---

## 9. Migration des générateurs

`xl_design.py` d'abord, puis générateur par générateur, avec **vérification
visuelle** à chaque étape (générer un fichier, ouvrir dans Excel).

| Fichier | Effort | Ce qui change |
|---|---|---|
| `src/generators/cockpit_ingenieur_generator.py` | moyen | remplace `src/styles.py` → `XD` ; jaune `FFFF99` → `FFF2CC` ; bannière + en-têtes aux tons d'onglet |
| `src/generators/dashboard_metier_generator.py` | moyen | idem ; Alertes au rouge OIL ; tables Light + en-tête manuel |
| `projet_TrainSystem/creer_uo.py` | moyen | migre de `design_b` → `XD` ; bannière sans nav ; glyphes ; palette unifiée |
| `projet_TrainSystem/creer_cockpit_se.py` | faible | styles inline → `XD` ; `Calibri` du _Manifeste → `Segoe UI` |

`src/styles.py` : conservé tant que d'autres modules en dépendent, puis retiré
quand plus aucun import (objectif : suppression à terme).

---

## 10. Décisions tranchées (log)

1. **Boutons de navigation** dans la bannière → **retirés** (réintroductibles).
2. **KPI → violet** (et non corail) pour l'éloigner du rouge OIL.
3. **Données d'entrée → famille teal/flux** (et non bleu) : c'est un onglet de
   suivi, proche des Activités.
4. **Planning → remonté en position 4** (après Description) et **recoloré en
   cyan/pétrole** pour préserver l'ordre logique ET la continuité chromatique.
5. **Run froid 3–6** : contraste renforcé (zigzag de luminosité).
6. **Glyphes monochromes** (et non emoji couleur) : sobres, corporate, stables
   tous OS.
7. **Tables nommées → style clair + en-tête manuel** (corrige Medium2).
8. **Spec écrite dans les deux** : ce fichier (versionné) + note dans le vault.
9. **Colonne spine santé en v1**, pilotée par **mise en forme conditionnelle**
   (statut + avancement), distincte de l'ID, hors de la table nommée.
10. **Filtrage = AutoFilter natif seul** ; pas de slicers en v1 (spike séparé si
    besoin).

---

## 11. Critères de succès

- [ ] `src/xl_design.py` existe, importable depuis n'importe quel générateur.
- [ ] `cockpit_ingenieur_generator.py` : 0 `PatternFill(` / `Font(` défini localement.
- [ ] `dashboard_metier_generator.py` : idem.
- [ ] `creer_uo.py` : idem (passe de `design_b` à `XD`).
- [ ] `creer_cockpit_se.py` : idem (plus de `Calibri`, plus de helpers `_fill/_fnt`).
- [ ] Générer un cockpit + un dashboard + une UO → mêmes couleurs, même police, même look.
- [ ] Chaque onglet a sa bannière + son en-tête au ton de sa famille + son glyphe.
- [ ] Jaune de saisie `FFF2CC` identique partout.
- [ ] Colonne spine présente sur les onglets à tableau, recolorée live par mise
      en forme conditionnelle (statut + avancement), hors de la table nommée.
- [ ] AutoFilter actif sur les tables nommées (filtrage rapide par colonne).
- [ ] `python -m pytest tests/ -q` → 382+ passed, 0 failed.
- [ ] Un non-technicien ouvre les fichiers et les trouve professionnels et cohérents.
