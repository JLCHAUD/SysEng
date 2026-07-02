# Manuel d'instruction collègue Excel (humain, pas LLM)

- **Date** : 2026-07-02
- **Conversation** : CONV-15 — Manuel Collègue Excel
- **Branche cible** : `feature/manuel-collegue-excel`
- **Statut** : spec validée, prête pour plan d'implémentation
- **Dépôt** : `SysEng` (github.com/JLCHAUD/SysEng)

---

## 1. Objectif & périmètre

CONV-14 a livré un package de contexte **pour le LLM** de la collègue
(`docs/onboarding-excel/` + `ESCALADES.md`). Ce package suppose qu'elle sache
déjà quel fichier utiliser dans quelle situation — ce n'est pas son rôle,
c'est celui d'un manuel **pour elle, humaine**.

Ce projet livre un document Word (`manuel_collegue_excel.docx`), écrit dans
un langage très pédagogique, qui explique concrètement : dans telle
situation, ouvre tel fichier, colle tel prompt dans ton LLM, mets la réponse
ici, vérifie comme ça, sauvegarde comme ça. Zéro ambiguïté possible sur le
"quoi faire, avec quoi, dans quel ordre".

**Hors périmètre** : ce document ne remplace pas `docs/onboarding-excel/`
(qui reste le contexte donné au LLM) — il l'orchestre. Aucune modification
du package existant n'est prévue dans ce projet.

---

## 2. Contexte — pour qui, avec quel outil

- **Public** : une collègue avec des notions de base (sait ouvrir un
  terminal, mais git et Python restent flous — chaque commande doit être
  expliquée en langage clair, aucun raccourci de vocabulaire non défini)
- **Outil LLM** : un chat simple (type ChatGPT/Gemini web), **sans accès
  direct aux fichiers du dépôt**. Conséquence structurante : chaque
  fiche-recette doit dire explicitement "ouvre le fichier X, sélectionne
  tout, copie, colle dans le chat" — jamais "donne le fichier X à ton LLM"
  de façon abstraite.
- **Format final** : Word (`.docx`), enregistré dans le dépôt Git SysEng à
  côté de `docs/onboarding-excel/`, pour rester avec le projet même si le
  format de consultation quotidien est Word plutôt que Markdown.

---

## 3. Structure du document

### 3.1 Premiers pas
Section courte : vérifier que le dossier du projet est accessible, lancer
`pytest` une fois pour confirmer que tout fonctionne avant de commencer.
Renvoie vers la section "Setup technique" de `00-DEMARRAGE.md` pour
l'installation complète (Python, dépendances) — pas dupliqué ici.

### 3.2 "Qu'est-ce que tu veux faire aujourd'hui ?"
Un tableau de décision : situation → renvoi direct vers la fiche-recette
correspondante (numéro de page / section). Elle n'a jamais besoin de tout
relire pour retrouver son cas.

### 3.3 Fiches-recettes (6 situations)
Chaque fiche suit une structure fixe identique, pour que la lecture devienne
un réflexe :

1. **Ce que tu veux faire** — reformulation en langage courant de la
   situation
2. **Fichier(s) à ouvrir et copier** — chemin exact, rappel du geste
   ("ouvre le fichier dans un éditeur de texte, Ctrl+A, Ctrl+C")
3. **Prompt à coller** — texte prêt à copier-coller tel quel dans le chat,
   après le contenu du fichier collé juste avant
4. **Où mettre la réponse** — quel fichier ouvrir, où coller le résultat du
   LLM, comment sauvegarder (Ctrl+S)
5. **Comment vérifier** — commande(s) exacte(s) à taper (`pytest`,
   `python -m src sync --dir projet_TrainSystem`), et à quoi ressemble un
   résultat correct vs un résultat qui indique un problème
6. **Comment sauvegarder dans Git** — commandes `git add`/`commit`/`push`
   explicitées en langage clair (ce que chaque commande fait, pourquoi)

Les 6 situations couvertes, avec le fichier de contexte LLM correspondant
(déjà livré par CONV-14) :

| # | Situation | Fichier(s) de contexte à copier |
|---|---|---|
| 1 | Comprendre le projet dans son ensemble | `01-Vue-Ensemble-Projet.md` |
| 2 | Écrire ou modifier un `_Manifeste` | `02-Langage-MXL.md` |
| 3 | Changer la mise en forme d'un Excel généré | `03-Design-System-XD.md` |
| 4 | Modifier un générateur Python | `02-Langage-MXL.md` + `03-Design-System-XD.md` + `04-Territoire-Et-Conventions.md` |
| 5 | Sauvegarder son travail (commit/push/PR) | `04-Territoire-Et-Conventions.md` |
| 6 | Signaler un besoin qui dépasse son périmètre | `05-Escalade.md` + `ESCALADES.md` (racine) |

### 3.4 Dépannage / FAQ
Réponses aux blocages typiques, formulées en langage rassurant et concret :
- "`pytest` affiche des erreurs, qu'est-ce que je fais ?"
- "Je ne sais pas quoi copier-coller, je suis perdue"
- "J'ai cassé quelque chose, comment revenir en arrière ?" (`git checkout`
  sur un fichier précis — commande sûre à utiliser, expliquée pas à pas,
  sans risque de perdre du travail sur d'autres fichiers)
- "Comment je vérifie que mon fichier Excel est bien généré correctement ?"
  (ouvrir le `.xlsx` généré, où le trouver, à quoi s'attendre visuellement)

---

## 4. Ton et niveau de langage

- Phrases courtes, une idée par phrase
- Chaque terme technique (commit, branche, sync, prompt...) défini en une
  phrase la première fois qu'il apparaît dans le document
- Instructions numérotées, jamais de paragraphe dense qui mélange plusieurs
  étapes
- Ton rassurant : anticiper l'inquiétude ("si tu vois ce message, c'est
  normal, voici pourquoi"), jamais de sous-entendu qu'une erreur est de sa
  faute
- Pas de jargon Git avancé (rebase, cherry-pick, etc.) — seulement les
  commandes strictement nécessaires au workflow décrit dans
  `04-Territoire-Et-Conventions.md`

---

## 5. Livrable et emplacement

- Fichier : `manuel_collegue_excel.docx`
- Emplacement : racine du dépôt SysEng (à côté de `README.md`, visible
  immédiatement), pas dans `docs/onboarding-excel/` pour ne pas le mélanger
  avec le package destiné au LLM
- Commité sur une branche dédiée `feature/manuel-collegue-excel`, mergée
  après revue comme CONV-14

---

## 6. Critères de succès

- Le fichier `manuel_collegue_excel.docx` existe à la racine du dépôt et est
  commité
- Les 6 fiches-recettes sont présentes, chacune avec les 6 sous-sections
  fixes (quoi faire / fichier à copier / prompt / où coller / vérifier /
  sauvegarder Git)
- Chaque prompt de fiche-recette est un texte concret, prêt à copier-coller
  (pas une description abstraite du type "demande à ton LLM de...")
- Aucun terme technique n'apparaît sans une explication en langage clair à
  sa première occurrence
- Le document reste cohérent avec `docs/onboarding-excel/` existant : les
  fichiers de contexte cités existent réellement et sont à jour avec ce qui
  a été livré en CONV-14
