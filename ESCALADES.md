# Escalades — besoins moteur MXL

Ce fichier sert à faire remonter un besoin qui dépasse le périmètre "Excel"
(générateurs, `_Manifeste`, design system) et qui nécessiterait une évolution
du **cœur du moteur** (`src/parser.py`, `src/executor.py`, `src/models.py`).

**Qui écrit ici** : la personne qui développe la partie Excel du projet.
**Qui relit** : le binôme qui développe le cœur (relecture périodique
manuelle, pas de notification automatique).

## Ce qui relève d'une escalade

- Une instruction MXL qui n'existe pas dans `_Manifeste` et dont tu aurais
  besoin (nouveau mot-clé, nouvelle fonction `COMPUTE`, nouvelle règle
  `VALIDATE`)
- Un comportement du moteur qui semble incohérent ou bloquant lors d'un
  `python -m src sync`
- Une question de conception sur `src/models.py` (structure de données
  partagée entre le cœur et les générateurs)

## Ce qui NE relève PAS d'une escalade (tu peux le faire toi-même)

- Mise en forme Excel (tout passe par `src/xl_design.py` — voir
  `docs/onboarding-excel/03-Design-System-XD.md`)
- Structure d'une feuille, ajout d'une colonne dans un générateur existant
- Utilisation d'une instruction MXL qui existe déjà dans le moteur mais n'est
  pas encore utilisée par les générateurs (voir la liste "disponible mais pas
  exploité" dans `docs/onboarding-excel/02-Langage-MXL.md`) — dans ce cas,
  utilise-la directement, pas besoin d'escalade

## Format d'une entrée

```markdown
## AAAA-MM-JJ — Titre court du besoin

**Contexte** : quel fichier/générateur, quel objectif métier.
**Ce qui manque** : ce que le moteur ne permet pas de faire aujourd'hui.
**Ce que je voudrais écrire** : la syntaxe MXL que tu voudrais pouvoir
utiliser (même approximative — l'idée compte plus que la syntaxe exacte).
**Contournement actuel** : comment tu fais en attendant (souvent : logique
codée en dur côté générateur Python, pas idéal, pas synchronisé si le store
change).
```

## Exemple

## 2026-07-15 — Besoin d'une fonction COMPUTE=MEDIANE

**Contexte** : dashboard métier, calcul du délai médian de clôture des UO.
**Ce qui manque** : `COMPUTE` supporte `MEAN_WEIGHTED`, `SUM`, `AVG`, `MIN`,
`MAX` mais pas de médiane.
**Ce que je voudrais écrire** :
`DEF $delai_median = COMPUTE(MEDIANE($uos_closes.delai_jours))`
**Contournement actuel** : calcul en dur côté générateur Python (pas idéal,
pas synchronisé si le store change).

---

## Entrées

<!-- Ajoute tes entrées ci-dessous, la plus récente en premier -->
