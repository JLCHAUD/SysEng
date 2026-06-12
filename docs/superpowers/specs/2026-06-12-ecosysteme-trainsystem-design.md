# Design — Écosystème Train System complet

**Date :** 2026-06-12  
**Statut :** approuvé  
**Branche :** feature/design-excel (puis merge master)

## Contexte

Jusqu'ici : moteur ExoSync validé, UO type conçue, assembleur `creer_uo.py` opérationnel, Design B (Cockpit) validé visuellement.

Objectif de cette itération : intégrer le Design B dans toutes les UOs générées, et construire les couches supérieures de la pyramide — cockpit SE (vue par ingénieur) et dashboard métier (vue pilote de pôle).

---

## 1. Évolutions du fichier UO

### 1.1 Nouvelles colonnes dans `tbl_activites`

Deux colonnes ajoutées après `commentaire` :

| Colonne | Type | Saisie | Rôle |
|---|---|---|---|
| `date_demarrage` | date | ingénieur | Quand l'activité commence réellement |
| `heures_par_semaine` | nombre | ingénieur | Rythme hebdo alloué à cette activité |
| `date_fin_estimee` | date | **formule** | `=date_demarrage + (reste_a_faire / heures_par_semaine) * 7` |

`date_fin_estimee` est une formule Excel locale — elle n'est pas lue par le moteur (contrainte data_only). Elle est lue par le moteur via la valeur calculée à l'ouverture Excel, ou exportée via PUSH explicite depuis une colonne de valeurs.

> **Décision** : `date_fin_estimee` n'est pas PUSHée en v1 — trop complexe (colonne formule). On PUSHe à la place `date_demarrage` et `heures_par_semaine` bruts ; le cockpit recalcule lui-même la date fin.

### 1.2 Nouveau champ dans `General`

- `date_livraison` — plage nommée, date contractuelle de livraison de l'UO. Saisie à la création via `--livraison YYYY-MM-DD`.
- PUSHée : `uo.<code>.date_livraison`

### 1.3 Design B intégré dans `creer_uo.py`

Toutes les fonctions de design (`banner_B`, `section_box`, `make_donut`, `activites_sheet`, `oil_sheet`, palette complète) sont extraites de `design_demo.py` dans un module partagé `projet_TrainSystem/uo_design.py`. `creer_uo.py` importe ce module pour styliser chaque feuille générée.

Règle de couleur par onglet :
- Dashboard → marine `0C447C`
- Activités → vert `0F6E56`
- OIL → ambre `BA7517`
- General, Description, Livrables, Donnees_Entree, KPI, Planning, Orga → marine `185FA5`

### 1.4 Nouvelles clés PUSH dans le Manifeste

```
PUSH $date_livraison -> uo.<code>.date_livraison
PUSH $act_table      -> uo.<code>.activites        # déjà existant, enrichi
```

La table activités PUSHée contient les colonnes `date_demarrage` et `heures_par_semaine` — le cockpit peut recalculer la charge à partir de là.

---

## 2. Les 5 UOs de démonstration

Générées via `creer_uo.py`, puis enrichies avec des données de démo réalistes (avancements variés, PO, dates, heures/semaine) par un script `demo_ecosystem.py`.

| Code UO | SE | Heures vendues | Date livraison |
|---|---|---|---|
| L09U1-CFL2400-CLIM | Alice | 200h | 2026-09-30 |
| L09U2-CFL2400-CLIM | Alice | 150h | 2026-10-31 |
| L10U1-RERNG-ELEC | Alice | 180h | 2026-08-31 |
| L11U1-RERNG-FREIN | Bob | 240h | 2026-11-30 |
| L11U2-RERNG-FREIN | Bob | 120h | 2027-02-28 |

`demo_ecosystem.py` :
1. Crée les 5 UOs via `creer_uo.py`
2. Injecte des données de démo réalistes dans chaque fichier (openpyxl)
3. Lance `valider_un.py` sur chacun pour peupler le store
4. Lance `creer_cockpit_se.py` pour Alice et Bob
5. Lance `creer_dashboard_metier.py`

---

## 3. Cockpit SE

### Fichiers générés
- `Cockpit_Alice.xlsx` (UOs : L09U1, L09U2, L10U1)
- `Cockpit_Bob.xlsx` (UOs : L11U1, L11U2)

### Commande
```bash
python projet_TrainSystem/creer_cockpit_se.py \
    --se "Alice" --uo L09U1-CFL2400-CLIM L09U2-CFL2400-CLIM L10U1-RERNG-ELEC \
    --capacite 35
```

### Layout (Design B, bandeau vert `0F6E56`)

**Feuille unique : `Cockpit`**

1. **Bandeau** (lignes 1-4) : "Cockpit Alice — Lot 9+10 · 3 UOs · 530h vendues" · navigation interne si plusieurs onglets à terme
2. **Section Cartes UO** (lignes 6-12) : une carte par UO (3 colonnes de 5 col chacune)
   - Titre UO + lien HYPERLINK vers le fichier (chemin relatif)
   - Avancement % (grand chiffre)
   - Badge santé (VERT/ORANGE/ROUGE)
   - PO ouverts / dont critiques
   - Heures consommées / vendues / RAF
3. **Section Workload** (lignes 14-30) : graphique barres empilées (openpyxl DoughnutChart remplacé par BarChart)
   - X : 12 semaines glissantes à partir d'aujourd'hui
   - Y : heures/semaine par UO (empilées, couleur par UO)
   - Ligne de capacité = `--capacite` (défaut 35h/sem)
   - Feuille masquée `_workload_data` pour les données du graphique
4. **Section Alertes** (lignes 32-45) : tableau des PO ouverts balle chez SE (toutes UOs), triés par criticité puis date_besoin

### Mécanisme de données
- Lit le store ExoSync (`python -m src status --prefix uo.` → JSON)
- Filtre les clés correspondant aux UOs listées
- Recalcule la charge hebdo depuis la table activités (date_demarrage + heures_par_semaine + RAF)

---

## 4. Dashboard Métier

### Fichier généré
- `Dashboard_Metier_TrainSystem.xlsx`

### Commande
```bash
python projet_TrainSystem/creer_dashboard_metier.py
```
(lit toutes les clés `uo.*` du store)

### Layout (Design B, bandeau ambre `BA7517`)

**Feuille unique : `Dashboard`**

1. **Bandeau** (lignes 1-4) : "Dashboard Métier — Train System · 5 UOs · 2 SEs"
2. **Bandeau synthèse** (ligne 6) : compteurs — UOs actives · en risque · critiques · RAF total
3. **Section Risques délais** (lignes 8-20) : tableau trié par risque décroissant
   - Colonnes : UO · SE · Avancement · Date contractuelle · Date fin estimée · Écart (jours) · Risque
   - Badge couleur : CRITIQUE (rouge) / RISQUE (ambre) / OK (vert)
   - Mise en forme conditionnelle sur l'écart
4. **Section Chemin critique** (lignes 22-32) : activités bloquantes des UOs à risque
   - UO · Activité · Date fin estimée · Balle chez · RAF · h/sem
   - Bordure gauche colorée selon criticité
5. **Section Workload globale** (lignes 34-48) : barres groupées Alice vs Bob par semaine

### Calcul du risque
```
ecart_jours = date_livraison - date_fin_estimee_max
  CRITIQUE  : ecart < -14j
  RISQUE    : -14j ≤ ecart < 0j
  OK        : ecart ≥ 0j

date_fin_estimee_max = max(
    date_demarrage + (RAF_activite / heures_par_semaine) * 7
    pour chaque activité applicable et non TERMINEE
)
```

---

## 5. Séquence d'implémentation

```
Étape 1 : uo_design.py        — extraire Design B dans un module partagé
Étape 2 : creer_uo.py         — importer uo_design, ajouter colonnes + date_livraison
Étape 3 : demo_ecosystem.py   — générer et syncer les 5 UOs
Étape 4 : creer_cockpit_se.py — cockpit par SE (store → Excel)
Étape 5 : creer_dashboard_metier.py — dashboard global (store → Excel)
Étape 6 : commit + push feature/design-excel
```

---

## 6. Ce qui n'est PAS dans ce spec (backlog)

- Planning onglet Gantt dans les UOs individuelles (phase 2)
- Mise à jour automatique du cockpit/dashboard à chaque sync (scheduler)
- Vue expert (LOT par LOT, filtre par domaine OIL)
- Authentification / droits par SE
