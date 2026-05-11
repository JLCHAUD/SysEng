# Fichiers exemples — SysEng (SCHEMA-01)

Exemples de données conformes au schéma SCHEMA-01-SysEng.
Chaque fichier illustre le contenu attendu d'un Post Excel pour une classe donnée.

## Fichiers

| Fichier | Classe | Post exemple |
|---------|--------|--------------|
| `syseng_referentiel_uo.json` | `referentiel_uo` | REF-UO-01 — catalogue 5 types d'UO |
| `syseng_referentiel_acteurs.json` | `referentiel_acteurs` | REF-ACT-01 — 7 acteurs externes |
| `syseng_uo_instance_L03U12-P042.json` | `uo_instance` | L03U12-P042 — Climatisation MI20 RATP |
| `syseng_cockpit_ingenieur_USR001.json` | `cockpit_ingenieur` | COCKPIT-ING-01 — Alice Dubois |
| `syseng_cockpit_metier_MET01.json` | `cockpit_metier` | COCKPIT-MET-01 — Jean-Luc Bernard |
| `syseng_cockpit_pole.json` | `cockpit_pole` | COCKPIT-POLE-01 — Pierre Legrand |

## Structure d'un fichier exemple

Chaque fichier contient :
- `_comment`, `_classe`, `_post`, `_owner` : métadonnées de contexte
- `min_fields` : champs scalaires DEF publiés au store
- Tables standard (`TabActivites`, `SyntheseMetier`, etc.)

## Correspondance avec le registre

Les IDs de Posts correspondent à `config/registre.json`.
Les IDs d'acteurs (`USR00x`) correspondent à `config/acteurs.json`.
Les IDs d'acteurs externes (`ACT-00x`) correspondent au `TabActeurs` du referentiel_acteurs.
