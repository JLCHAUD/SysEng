# Manuel pour toi — travailler sur les fichiers Excel du projet ExoSync

Ce document est écrit pour toi, pas pour ton LLM (ton assistant IA — l'outil
de conversation avec lequel tu vas travailler). Il t'explique, situation
par situation, quel fichier ouvrir, quel texte coller dans ton assistant IA,
où mettre sa réponse, comment vérifier que ça marche, et comment sauvegarder
ton travail. Chaque terme technique est expliqué la première fois qu'il
apparaît — tu n'as besoin de rien connaître à l'avance.

## Avant de commencer

1. Assure-toi d'avoir le dossier du projet ouvert sur ton ordinateur (le
   dossier qui contient un fichier `README.md` et un dossier `src`).
2. Ouvre un terminal dans ce dossier (une fenêtre noire où tu peux taper des
   commandes — sur Windows, clic droit dans le dossier puis "Ouvrir dans le
   terminal", ou équivalent selon ton système).
3. Tape `pytest` puis appuie sur Entrée. Cette commande lance tous les tests
   automatiques du projet — c'est une vérification de routine, pas quelque
   chose que tu dois comprendre en détail. Si tu vois à la fin une ligne avec
   le mot "passed" en vert et aucun "failed", tout va bien, tu peux
   continuer.
4. Si cette commande ne fonctionne pas du tout (message "commande introuvable"
   par exemple), c'est que ton environnement n'est pas encore installé —
   demande de l'aide avant d'aller plus loin, ce manuel suppose que
   l'installation de base (Python — le langage de programmation utilisé par
   ce projet — et les outils du projet) est déjà faite.

## Qu'est-ce que tu veux faire aujourd'hui ?

Trouve la ligne qui ressemble le plus à ta situation, puis va directement à
la fiche indiquée. Pas besoin de lire tout le document dans l'ordre.

| Ta situation | Va à |
|---|---|
| Je découvre le projet, je veux comprendre de quoi il s'agit | Fiche 1 |
| Je dois écrire ou modifier une ligne dans la feuille `_Manifeste` d'un fichier Excel | Fiche 2 |
| Je dois changer une couleur, une police, une mise en forme dans un Excel généré automatiquement | Fiche 3 |
| Je dois modifier le code Python qui génère les fichiers Excel (nouvelle colonne, nouveau calcul, nouvelle feuille) | Fiche 4 |
| J'ai fini une modification et je veux l'enregistrer officiellement | Fiche 5 |
| Je suis bloquée, ce que je veux faire n'est pas possible avec les outils actuels | Fiche 6 |
| Quelque chose ne marche pas, un test échoue, je ne sais pas quoi faire | Section Dépannage, à la fin de ce document |

---

## Fiche 1 — Comprendre le projet

**Ce que tu veux faire** : tu commences à travailler sur ce projet et tu
veux comprendre de quoi il s'agit avant de te lancer.

**Fichier à ouvrir et copier** : `docs/onboarding-excel/01-Vue-Ensemble-Projet.md`

Ouvre ce fichier dans un éditeur de texte (Bloc-notes, VS Code, ou n'importe
quel logiciel qui affiche du texte), sélectionne tout son contenu (touches
Ctrl et A en même temps), copie (Ctrl et C).

**Prompt à coller** (colle d'abord le contenu du fichier copié, puis en
dessous, colle ce texte) :

> Voici la description d'un projet appelé ExoSync. Explique-moi en langage
> simple : à quoi sert ce projet, comment les différents fichiers Excel (UO,
> cockpit, dashboard) sont reliés entre eux, et où se situe ma zone de
> travail à moi dans tout ça.

**Où mettre la réponse** : nulle part — cette fiche sert juste à comprendre,
pas à produire un fichier. Lis la réponse, pose des questions de suivi si
quelque chose n'est pas clair.

**Comment vérifier** : pas de vérification technique ici — le vrai test,
c'est que tu puisses répondre toi-même à la question "à quoi sert ce
projet ?" en une phrase simple.

**Sauvegarder dans Git** : rien à sauvegarder pour cette fiche.

---

## Fiche 2 — Écrire ou modifier un `_Manifeste`

**Ce que tu veux faire** : un fichier Excel (UO, cockpit ou dashboard) a
besoin d'une nouvelle ligne dans sa feuille `_Manifeste` (la feuille qui
dicte comment ce fichier échange des données avec les autres), ou tu dois
corriger une ligne existante.

**Fichier à ouvrir et copier** : `docs/onboarding-excel/02-Langage-MXL.md`
(ouvre, Ctrl+A, Ctrl+C, comme dans la Fiche 1).

**Prompt à coller** (après le contenu du fichier) :

> Voici la référence du langage MXL utilisé dans les fichiers `_Manifeste`
> du projet ExoSync. Je travaille sur le fichier [NOM DU FICHIER, par
> exemple Cockpit_Alice_Dubois.xlsx], feuille `_Manifeste`. Je veux [DÉCRIS
> PRÉCISÉMENT CE QUE TU VEUX FAIRE, par exemple "ajouter une ligne qui
> exporte le nombre total d'UO vers le store central"]. Donne-moi la ou les
> lignes MXL exactes à ajouter, avec une explication de chaque partie.

Remplace tout ce qui est entre crochets `[ ]` par ta situation réelle avant
de coller ce texte — n'oublie pas d'enlever les crochets.

**Où mettre la réponse** : ouvre le fichier Excel concerné dans Excel, va
sur la feuille `_Manifeste`, trouve la première ligne vide en colonne A
(juste sous la dernière instruction déjà écrite), colle la ou les lignes que
ton LLM t'a données — une instruction par ligne, dans la colonne A. Si le
LLM te propose aussi un commentaire explicatif, mets-le en colonne C, sur la
même ligne que l'instruction.

**Comment vérifier** : ouvre un terminal dans le dossier du projet, tape
`python -m src sync --dir projet_TrainSystem` puis Entrée. Si un message
d'erreur mentionne ton fichier, quelque chose ne va pas dans ce que tu as
collé — retourne voir ton LLM avec le message d'erreur exact copié-collé et
demande-lui de corriger. Si la commande se termine sans erreur affichée pour
ton fichier, c'est bon signe.

**Sauvegarder dans Git** : cette fiche ne modifie que le fichier Excel
lui-même, qui n'est jamais sauvegardé dans Git (les fichiers `.xlsx` ne sont
pas suivis par le projet). Si tu as aussi modifié du code Python en même
temps, va à la Fiche 5.

---

## Fiche 3 — Changer la mise en forme d'un Excel généré

**Ce que tu veux faire** : tu veux changer une couleur, une taille de
police, ajouter une bordure... sur un onglet d'un fichier Excel généré
automatiquement par le projet. Important : il faut modifier le code Python
qui génère ce fichier, pas modifier le fichier Excel directement à la main —
sinon ta modification disparaîtra la prochaine fois que le fichier sera
régénéré.

**Fichier à ouvrir et copier** : `docs/onboarding-excel/03-Design-System-XD.md`

**Prompt à coller** :

> Voici la documentation du système de mise en forme (design system) utilisé
> par le projet ExoSync pour styliser les fichiers Excel générés
> automatiquement. Je veux [DÉCRIS CE QUE TU VEUX CHANGER, par exemple "que
> la couleur des badges de statut EN_COURS soit bleu plutôt que vert"].
> Dis-moi précisément quel fichier modifier, quelle ligne, et donne-moi le
> code exact à écrire.

**Où mettre la réponse** : ouvre le fichier Python que le LLM t'indique
(probablement `src/xl_design.py` ou un fichier dans `src/generators/`) dans
un éditeur de texte ou VS Code, trouve la ligne indiquée, remplace-la par le
code donné par le LLM, sauvegarde (Ctrl+S).

**Comment vérifier** : dans un terminal, dans le dossier du projet, tape
`pytest` puis Entrée. Si tu vois "passed" en vert à la fin sans mention de
"failed", c'est bon signe. Si tu vois des erreurs, copie le message
d'erreur complet et retourne voir ton LLM avec.

**Sauvegarder dans Git** : va à la Fiche 5.
