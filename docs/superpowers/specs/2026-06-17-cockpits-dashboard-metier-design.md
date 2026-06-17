# Design — Cockpits Ingénieur & Dashboard Métier

**Date :** 2026-06-17  
**Projet :** ExoSync / SysEng  
**Statut :** Approuvé

---

## Contexte

ExoSync gère des UOs (Unités d'Œuvre) pour le pôle SysEng d'Alstom. Trois ingénieurs système (Alice Dubois, Bruno Lecomte, Camille Vidal) ont chacun plusieurs UOs. Un pilote métier (Jean-Luc Bernard) supervise les trois.

Les cockpits ingénieurs existent déjà mais sont orientés statut. Ce design introduit :
- Un cockpit ingénieur orienté **agenda** (quoi faire, quand)
- Un dashboard métier avec **vue consolidée** sur toute l'équipe
- Un mécanisme de **push/pull via store.json** (architecture ExoSync native)

---

## Architecture Globale

```
[ Cockpit_Alice.xlsx ]   [ Cockpit_Bruno.xlsx ]   [ Cockpit_Camille.xlsx ]
        ↓ push _Manifeste         ↓ push _Manifeste          ↓ push _Manifeste
                              [ store.json ]
                                    ↓ pull _Manifeste
                       [ Dashboard_JL_Bernard.xlsx ]
```

**Règle de filtrage :** USR004 (Jean-Luc Bernard) a `filtre_type: ingenieur`, `filtre_valeur: "Alice Dubois,Bruno Lecomte,Camille Vidal"`. Le dashboard ne contient que les UOs de ces trois ingénieurs.

---

## Cockpit Ingénieur Système

**Fichiers :** `output/cockpits/Cockpit_Alice_Dubois.xlsx`, `Cockpit_Bruno_Lecomte.xlsx`, `Cockpit_Camille_Vidal.xlsx`  
**Générateur :** `src/generators/cockpit_ingenieur_generator.py` (à créer)

### Onglet `Agenda`

Vue orientée "quoi faire cette semaine / ce mois" :

| Colonne | Contenu |
|---------|---------|
| UO ID | ID de l'UO concernée |
| Activité | Nom de l'activité |
| Priorité | Haute / Normale / Faible |
| Date échéance | Date limite de l'activité |
| Statut | EN_COURS / EN_RETARD / À_FAIRE |
| Action | Zone de saisie libre |

- Section **"Semaine en cours"** : activités dont l'échéance tombe dans les 7 prochains jours
- Section **"Prochaines échéances"** : horizon 30 jours
- Section **"Points ouverts"** : actions non résolues
- Mise en forme conditionnelle : rouge = EN_RETARD, orange = échéance < 3j, vert = OK

### Onglet `Mes UOs`

Vue synthèse de toutes les UOs de l'ingénieur :

| Colonne | Contenu | Saisie |
|---------|---------|--------|
| UO ID | ID (hyperlink vers fichier UO) | — |
| Type UO | Nom du type | — |
| Système | Nom du système | — |
| Projet | Nom du projet | — |
| Charge allouée (h) | Lu depuis store | — |
| % Avancement | **Zone de saisie ingénieur** | ✅ |
| H réalisées | **Zone de saisie ingénieur** | ✅ |
| Date fin | Lu depuis store | — |
| Alerte | Formule dérive heures / échéance | — |

**Zone de saisie** : colonnes F et G (% Avancement, H réalisées) sont les seules cellules modifiables par l'ingénieur. Fond coloré pour les distinguer.

### Onglet `_Manifeste`

Règles ExoSync de push/pull. La feuille contient une **3ème colonne `COMMENTAIRE`** (après `DIRECTION` et `NOM_GLOBAL`) qui décrit en langage naturel ce que fait la ligne MXL — obligatoire pour chaque règle générée.

| Direction | Clé store | Commentaire (col 3) | Feuille | Cellule/Tableau | Type |
|-----------|-----------|---------------------|---------|-----------------|------|
| PUSH | `uo.<id>.avancement` | Remonte le % d'avancement saisi par l'ingénieur vers le store central | Mes UOs | col % Avancement | CELL_PCT |
| PUSH | `uo.<id>.heures_realisees` | Remonte les heures réalisées saisies par l'ingénieur | Mes UOs | col H réalisées | CELL_NUM |
| PUSH | `uo.<id>.points_ouverts` | Remonte le nombre de points ouverts non résolus | Agenda | col Points ouverts | CELL_NUM |
| PULL | `uo.<id>.charge_allouee` | Injecte la charge allouée depuis le store (lecture seule) | Mes UOs | col Charge allouée | CELL_NUM |
| PULL | `uo.<id>.date_fin` | Injecte la date de fin planifiée depuis le store (lecture seule) | Mes UOs | col Date fin | CELL_DATE |

**Règle de génération :** le générateur Python doit remplir la colonne `COMMENTAIRE` pour chaque règle. Le commentaire doit être explicite en français, sans jargon technique MXL.

---

## Dashboard Métier

**Fichier :** `output/cockpits/Dashboard_JL_Bernard.xlsx`  
**Générateur :** `src/generators/dashboard_metier_generator.py` (à créer)  
**Acteur :** USR004 Jean-Luc Bernard (`pilote_metier`)

### Onglet `Synthèse`

**Bandeau KPIs (ligne 2-3) :**
- Nb total UOs équipe
- Charge totale équipe (h)
- % avancement moyen
- Nb alertes actives

**Tableau consolidé :**
Toutes les UOs des 3 ingénieurs dans un seul tableau, avec colonne `Ingénieur` et colonne `Alerte`. Triable par ingénieur, projet, système.

### Onglet `Par Ingénieur`

Une section par ingénieur (Alice / Bruno / Camille) :
- En-tête ingénieur avec total charge et % avancement moyen
- Tableau de ses UOs : UO ID, Type, Système, Projet, Charge, % Avancement, H réalisées, Date fin
- Lien vers son cockpit

### Onglet `Alertes`

Tableau trié par criticité décroissante :

| Colonne | Contenu |
|---------|---------|
| Ingénieur | Qui est concerné |
| UO ID | L'UO en alerte |
| Type alerte | Dépassement H / Échéance proche / Livrable retard |
| Détail | Description de l'alerte |
| Criticité | 🔴 Critique / 🟠 Élevée / 🟡 Normale |

Seuils :
- **Dépassement H** : `heures_realisees > charge_allouee`
- **Échéance critique** : `date_fin < aujourd'hui + 7j`
- **Livrable en retard** : `date_fin < aujourd'hui AND statut != LIVRE`

### Onglet `_Manifeste`

Règles de pull depuis store.json pour chaque UO de l'équipe. Même convention : **3ème colonne `COMMENTAIRE`** obligatoire.

| Direction | Clé store | Commentaire (col 3) | Feuille | Type |
|-----------|-----------|---------------------|---------|------|
| PULL | `uo.<id>.avancement` | Récupère l'avancement poussé par le cockpit ingénieur | Synthèse + Par Ingénieur | CELL_PCT |
| PULL | `uo.<id>.heures_realisees` | Récupère les heures réalisées poussées par le cockpit ingénieur | Synthèse + Par Ingénieur | CELL_NUM |
| PULL | `uo.<id>.points_ouverts` | Récupère le nb de points ouverts pour alimentation des alertes | Alertes | CELL_NUM |

---

## Données échangées via store.json

```json
{
  "uo.L03U12-P042.avancement": 0.65,
  "uo.L03U12-P042.heures_realisees": 20,
  "uo.L03U12-P042.points_ouverts": 2,
  "uo.L03U12-P042.charge_allouee": 32,
  "uo.L03U12-P042.date_fin": "2026-05-30"
}
```

Ces clés sont définies dans le `_Manifeste` de chaque cockpit ingénieur et lues par le `_Manifeste` du dashboard métier.

---

## Nouveaux fichiers à créer

| Fichier | Rôle |
|---------|------|
| `src/generators/cockpit_ingenieur_generator.py` | Remplace `cockpit_generator.py` — génère Agenda + Mes UOs + _Manifeste |
| `src/generators/dashboard_metier_generator.py` | Génère le dashboard avec filtrage par acteur |
| `tests/test_cockpit_ingenieur.py` | Tests génération cockpit + vérification onglets |
| `tests/test_dashboard_metier.py` | Tests génération dashboard + filtrage + alertes + push/pull |

---

## Tests à écrire

### `test_cockpit_ingenieur.py`
- Génération cockpit pour Alice → fichier créé
- Onglets `Agenda`, `Mes UOs`, `_Manifeste` présents
- Zone de saisie (% Avancement, H réalisées) correctement positionnée
- Formule alerte dérive présente sur toutes les lignes UO
- Contenu filtré : seules les UOs de Alice dans `Mes UOs`
- Section "Semaine en cours" contient uniquement activités échéance ≤ 7j

### `test_dashboard_metier.py`
- Génération dashboard pour Jean-Luc Bernard → fichier créé
- Onglets `Synthèse`, `Par Ingénieur`, `Alertes`, `_Manifeste` présents
- **Filtrage respecté** : pas d'UOs de Denis Renard dans le dashboard
- KPI charge totale = somme des charges des 3 ingénieurs
- Onglet Alertes contient les UOs dont `heures_realisees > charge_allouee`
- Cycle push/pull : mock store avec avancement=0.8 → dashboard affiche 0.8

---

## Contraintes techniques

- Générateurs dans `src/generators/` — pattern identique aux existants
- Filtrage via `ProfilActeur.filtre_valeur` (liste CSV des noms ingénieurs)
- `_Manifeste` suit la spec ExoSync existante (`passerelle.py`)
- Pas de régression sur les 334 tests existants
- Fichiers de sortie dans `output/cockpits/`
