# Vue d'ensemble du projet ExoSync

## Le problème que ça résout

Un ingénieur travaille sur un fichier Excel (une "UO", Unité d'Œuvre). Son
avancement doit remonter à son cockpit personnel, qui agrège toutes ses UO.
Ce cockpit doit à son tour remonter vers le dashboard de son responsable
métier, qui voit la synthèse de toute son équipe. Sans outil, ça veut dire
resaisir la même donnée à plusieurs endroits, ou construire des liaisons
Excel fragiles.

ExoSync automatise cette remontée : chaque fichier Excel porte une feuille
`_Manifeste` qui décrit ce qu'il donne (`PUSH`) et ce qu'il reçoit (`PULL`,
`LIST`/`COLLECT`) des autres fichiers de l'écosystème. Un script Python
(`python -m src sync`) lit tous les `_Manifeste`, exécute les instructions, et
met à jour les fichiers en conséquence.

## Le concept d'exostructure

Il n'existe **pas** de base centrale qui décrit "voici tous les fichiers du
projet et comment ils sont reliés". Cette structure émerge dynamiquement à
partir des `_Manifeste` : c'est le fichier lui-même qui porte sa propre
identité (`FILE_TYPE`, `FILE_ID`) et ses propres règles de synchronisation.
C'est pour ça qu'on parle d'**exostructure** — la structure vit à l'extérieur
d'une base centrale, distribuée dans les fichiers.

## La pyramide des vues

```
DASHBOARD CLIENT (Alstom, filtré par le leader métier)      ← pas encore fait
      ▲ re-PUSH namespace client.*
DASHBOARD MÉTIER (leader Train System)                       ← fait
      ▲ LIST mes_cockpits TYPE=cockpit_ingenieur WHERE pilote_id=...
      ▲ COLLECT tbl_mes_uos FROM mes_cockpits INTO tbl_vue_synthese
COCKPIT INGÉNIEUR (×20+, un par ingénieur) — hub du matin     ← fait
      ▲ PUSH cockpit.<nom>.mes_uos
UO INSTANCES (50-150 fichiers vivants)                        ← fait
      L{NN}U{NN}-{PPPP}-{SSSS}  (ex: L09U1-CFL2400-CLIM)
```

- Une **UO instance** est le fichier de travail quotidien d'un ingénieur :
  activités, points ouverts, livrables.
- Un **cockpit ingénieur** agrège toutes les UO d'un même ingénieur — c'est
  son hub du matin : avancement, alertes, liens vers ses fichiers.
- Un **dashboard métier** agrège tous les cockpits des ingénieurs d'un même
  responsable (filtré par `pilote_id`) — synthèse d'équipe.
- Le **dashboard client** (pas encore construit) re-publiera une vue filtrée
  vers le client final (Alstom), avec curation de ce qui est montré.

## Projet réel : catalogue Train System

Le projet concret est un catalogue d'Unités d'Œuvre pour le client Alstom,
domaine Train System (ingénierie système). Structure : 5 lots (phases projet)
× 2 UO (système / sous-système), ~20+ ingénieurs, 50-150 fichiers UO vivants
à terme. Convention de nommage : `L{NN}U{NN}-{PPPP}-{SSSS}` où `{PPPP}` est
le code projet et `{SSSS}` le code système (ex : `L09U1-CFL2400-CLIM` = lot 9,
UO1, projet CFL2400, système climatisation).

## État d'avancement (2026-07-02)

| Brique | Contenu | État |
|---|---|---|
| Noyau moteur | Parser + executor MXL | ✅ fait, 415 tests passants |
| Modèle UO réel | UO type Train System, catalogue 5 lots × 2 UO | ✅ fait |
| Pyramide interne | Cockpit ingénieur + dashboard métier fonctionnels | ✅ fait |
| Design system Excel | Charte graphique centralisée (`src/xl_design.py`) | ✅ fait |
| Vues externes | Dashboard client, vue par projet | ⬜ à faire |
| Chaîne d'instanciation | Assembleur : commande client → UO prête en < 5 min | 🔄 partiel (`creer_uo.py` existe, orchestration manuelle) |
| Contrat hybride | Contrat minimal UO figé (imposé vs libre) | ⬜ après les vues externes |

C'est précisément la partie "Excel" de ce tableau — modèle UO, pyramide
interne, design system, vues externes, chaîne d'instanciation — qui constitue
ton périmètre de travail.
