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
   l'installation de base (Python, les outils du projet) est déjà faite.

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
