# Spec — Extraction du module `design_b.py` (Phase 1)

**Date** : 2026-06-15  
**Scope** : Phase 1 uniquement — extraction et centralisation des primitives de style Design B.  
**Fichiers touchés** : `projet_TrainSystem/design_b.py` (nouveau), `projet_TrainSystem/design_demo.py` (refactoré).  
**Fichiers NON touchés** : `creer_uo.py`, `src/`, `tests/`.

---

## Objectif

La branche `feature/design-excel` a produit un prototype visuel validé (Design B "Cockpit bandeau")
dans `design_demo.py`. Ce fichier mêle données de démo et primitives de style réutilisables.
La phase 1 extrait ces primitives dans un module autonome `design_b.py` afin que la phase 2
(intégration dans `creer_uo.py`) puisse les importer sans dupliquer de code.

---

## Nouveau fichier : `projet_TrainSystem/design_b.py`

### 1. Palette & constantes

Toutes les constantes de couleur, police et bordure :

```python
NAVY_D = "0C447C"   # marine foncé — titres, bandeaux
NAVY   = "185FA5"   # marine — accents
BLUE   = "378ADD"   # bleu — data bars
BLUE_L = "E6F1FB"   # bleu pâle — fonds doux
GREY_D = "5F5E5A"   # gris foncé — texte secondaire
GREY_L = "F5F4F0"   # gris chaud pâle — fonds de carte
GREY_B = "D3D1C7"   # gris — bordures
GREEN  = "639922";  GREEN_L  = "EAF3DE";  GREEN_D  = "27500A"
AMBER  = "EF9F27";  AMBER_L  = "FAEEDA";  AMBER_D  = "854F0B"
RED    = "E24B4A";  RED_L    = "FCEBEB";  RED_D    = "791F1F"
WHITE  = "FFFFFF"
TEAL   = "1D7068";  TEAL_L   = "D6EDEB"   # bandeau Activités
AMB_B  = "9A6200"                          # bandeau OIL

F      = "Segoe UI"
THIN_G = Side(style="thin", color=GREY_B)
HAIR   = Border(left=THIN_G, right=THIN_G, top=THIN_G, bottom=THIN_G)
```

### 2. Primitives de style (6 fonctions)

| Fonction | Signature | Rôle |
|----------|-----------|------|
| `fnt()` | `(size, bold, color, italic)` | Crée un objet `Font` Segoe UI |
| `fill()` | `(color)` | Crée un `PatternFill` solid |
| `card_border()` | `(ws, r1, c1, r2, c2, side, color)` | Encadre une zone avec bordure fine |
| `add_table()` | `(ws, name, ref)` | Crée un tableau Excel nommé avec style Light15 |
| `statut_cf()` | `(ws, rng)` | Mise en forme conditionnelle badges statut (TERMINEE/EN_COURS/A_FAIRE/STAND_BY) |
| `criticite_cf()` | `(ws, rng)` | Mise en forme conditionnelle badges criticité (HAUTE/MOYENNE/BASSE) |

### 3. Constructeurs de bandeau (3 fonctions)

| Fonction | Couleur onglet | Utilisée pour |
|----------|---------------|---------------|
| `banner_B()` | `NAVY_D` | Dashboard, General, KPI |
| `banner_teal()` | `TEAL` | Activités, Livrables, Données d'entrée |
| `banner_amber()` | `AMB_B` | OIL |

Signature commune : `(ws, subtitle, ncols, ...)` — génère un bandeau 3 lignes en haut de la feuille.

### 4. Composants de mise en page (3 fonctions)

| Fonction | Rôle |
|----------|------|
| `section_box()` | Zone délimitée avec titre en bandeau gris pâle |
| `kpi_card_B()` | Carte KPI 4 colonnes : label + valeur + sous-titre + bordure colorée |
| `make_donut()` | Graphique anneau (jauge %) — lit depuis une feuille `_chart_data` cachée |

---

## Fichier refactoré : `projet_TrainSystem/design_demo.py`

- Supprime toutes les définitions extraites vers `design_b.py`
- Ajoute en tête : `from design_b import *` (ou imports explicites)
- Conserve : `ACTIVITES`, `OIL` (données de démo), `activites_sheet()`, `oil_sheet()`,
  `build_design_A()`, `build_design_B()`

### Note sur `activites_sheet()` et `oil_sheet()`

Ces fonctions utilisent les constantes `ACTIVITES` et `OIL` codées en dur dans `design_demo.py`.
Elles **ne sont pas extraites** en phase 1. La phase 2 écrira des variantes dynamiques
(acceptant des données catalogue) directement dans `creer_uo.py`.

---

## Critère de validation

```bash
python projet_TrainSystem/design_demo.py
```

Doit produire `Design_A_Studio.xlsx` et `Design_B_Cockpit.xlsx` **identiques** aux fichiers
actuels (mêmes styles, mêmes données de démo). Aucun test unitaire nouveau requis en phase 1.

---

## Ce que cette phase ne fait PAS

- Ne modifie pas `creer_uo.py`
- Ne modifie pas `src/`
- Ne change pas le comportement visible des fichiers Excel générés
- N'introduit pas de données dynamiques dans les sheet builders

---

## Phase suivante

Phase 2 : intégration de `design_b.py` dans `creer_uo.py` — appliquer le Design B
aux 11 feuilles de l'UO avec données dynamiques depuis le catalogue.
