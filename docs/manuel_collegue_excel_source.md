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
   chose que tu dois comprendre en détail. Si tu vois à la fin une ligne qui
   se termine par un nombre suivi de "passed", sans le mot "failed" nulle
   part, tout va bien, tu peux continuer.
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
sur la feuille `_Manifeste`, trouve la dernière ligne qui commence par une
instruction MXL (`DEF`, `COL`, `PUSH`, `PULL`, `BIND`, `VALIDATE`, `LIST`
ou `COLLECT` — pas une ligne d'en-tête comme `FILE_TYPE` ou `ingenieur`), et
colle la ou les lignes que ton LLM t'a données juste en dessous — une
instruction par ligne, dans la colonne A. Si le LLM te propose aussi un
commentaire explicatif, mets-le en colonne C, sur la même ligne que
l'instruction.

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
`pytest` puis Entrée. Si tu vois à la fin une ligne qui se termine par un
nombre suivi de "passed", sans le mot "failed" nulle part, c'est bon signe.
Si tu vois des erreurs, copie le message d'erreur complet et retourne voir
ton LLM avec.

**Sauvegarder dans Git** : va à la Fiche 5.

---

## Fiche 4 — Modifier un générateur Python

**Ce que tu veux faire** : tu dois ajouter une fonctionnalité plus
importante à un générateur — par exemple une nouvelle colonne dans un
tableau, une nouvelle feuille, un nouveau calcul.

**Fichiers à ouvrir et copier** : les 3 fichiers suivants, un par un, dans
cet ordre — copie le contenu du premier, colle-le dans le chat, puis
directement en dessous colle le contenu du deuxième, puis du troisième
(un retour à la ligne entre chaque, rien d'autre) :

1. `docs/onboarding-excel/02-Langage-MXL.md`
2. `docs/onboarding-excel/03-Design-System-XD.md`
3. `docs/onboarding-excel/04-Territoire-Et-Conventions.md`

**Prompt à coller** (après les 3 fichiers collés à la suite) :

> Voici trois documents de référence du projet ExoSync : le langage MXL, le
> système de mise en forme, et les conventions de développement. Je veux
> modifier le générateur [NOM DU FICHIER, par exemple
> `cockpit_ingenieur_generator.py`] pour [DÉCRIS PRÉCISÉMENT LE CHANGEMENT].
> Le projet suit le principe TDD (écrire le test avant le code) : propose-moi
> d'abord le test qui vérifie ce comportement, puis le code qui le fait
> passer, en respectant les conventions du troisième document.

**Où mettre la réponse** : ton LLM va te proposer du code de test ET du code
d'implémentation, dans deux blocs séparés. Ouvre les deux fichiers concernés
(un dans le dossier `tests/`, un dans `src/generators/` ou
`projet_TrainSystem/`), colle chaque morceau de code au bon endroit
(le LLM te dira dans quel fichier va chaque morceau), sauvegarde les deux
fichiers.

**Comment vérifier** : dans un terminal, tape `pytest` puis Entrée. Regarde
d'abord que ton nouveau test apparaît bien dans la liste affichée (cherche le
nom que ton LLM lui a donné). S'il y a écrit "failed" à côté de ce nom,
retourne voir ton LLM avec le message d'erreur exact copié-collé.

**Sauvegarder dans Git** : va à la Fiche 5 — c'est le cas typique où tu as
plusieurs fichiers à sauvegarder en même temps (le fichier de test et le
fichier de code).

---

## Fiche 5 — Sauvegarder son travail (Git)

**Ce que tu veux faire** : tu as terminé une modification (code ou
`_Manifeste`) et les vérifications de la fiche précédente sont bonnes. Tu
veux l'enregistrer officiellement dans le projet, pour que le reste de
l'équipe en profite aussi.

Cette fiche ne demande pas de fichier de contexte ni de prompt pour ton
LLM — ce sont des commandes que tu tapes toi-même dans le terminal.

**Les commandes, une par une** (tape chacune puis Entrée, dans l'ordre) :

0. Avant de committer, vérifie sur quelle branche tu es : tape `git branch`
   puis Entrée. La ligne qui commence par une étoile `*` est ta branche
   actuelle. Si c'est écrit `* master`, ne committe pas directement dessus —
   crée d'abord une branche pour ton travail : `git checkout -b
   feature/nom-court-de-ton-travail` (remplace par un nom qui décrit ce que
   tu fais, en gardant le préfixe `feature/` utilisé par le projet — par
   exemple `git checkout -b feature/ajoute-colonne-priorite`, sans espace ni
   accent). Une fois sur ta propre branche, continue avec les étapes
   ci-dessous.

1. `git status`
   Ça affiche la liste des fichiers que tu as modifiés depuis ta dernière
   sauvegarde. Vérifie que tu reconnais bien tous les fichiers listés — si
   un fichier inconnu apparaît, arrête-toi et demande de l'aide.

2. `git add nom_du_fichier`
   À répéter pour chaque fichier modifié (remplace `nom_du_fichier` par le
   chemin exact affiché par `git status`, par exemple
   `src/generators/cockpit_ingenieur_generator.py`). Cette commande prépare
   le fichier pour la sauvegarde, sans encore l'enregistrer définitivement.

3. `git commit -m "description courte de ce que tu as fait"`
   Ça enregistre officiellement tes modifications, avec un message qui
   explique ce que tu as changé. Le projet utilise un format précis :
   `type(scope): description`, où `type` est souvent `feat` (nouvelle
   fonctionnalité), `fix` (correction) ou `docs` (documentation), et `scope`
   est la partie du projet concernée. Écris ce message toi-même en suivant
   ce format (par exemple : `git commit -m "feat(cockpit): ajoute colonne
   priorite"`).

4. `git push`
   Ça envoie ta sauvegarde vers le serveur partagé (GitHub), pour que le
   reste de l'équipe puisse la voir.

Une fois ta branche poussée, préviens la personne qui gère le dépôt
principal (ou crée toi-même une Pull Request sur le site GitHub du projet
si tu sais le faire) : c'est une demande officielle pour que ton travail
soit relu puis intégré à `master`. Tant que ce n'est pas fait, ton travail
reste sur ta branche, séparé du reste — c'est normal et voulu, ça permet
une relecture avant que ça touche le travail de toute l'équipe.

**Comment vérifier** : après `git push`, tape `git status` une dernière
fois. Si tu vois "nothing to commit, working tree clean" et "Your branch is
up to date", tout est bien sauvegardé.

---

## Fiche 6 — Signaler un besoin qui dépasse ton périmètre (escalade)

**Ce que tu veux faire** : ton LLM te dit qu'il ne peut pas faire ce que tu
demandes parce que ça touche au "cœur du moteur" (les fichiers
`src/parser.py`, `src/executor.py` ou `src/models.py`), ou tu as
l'impression que ce que tu veux faire n'est tout simplement pas possible
avec les outils actuels du projet.

**Fichier à ouvrir et copier** : `docs/onboarding-excel/05-Escalade.md`
d'abord — pas pour le coller à un LLM, mais pour toi : lis-le et regarde le
tableau d'exemples dedans, pour confirmer que ta situation est bien une
"vraie" escalade et pas quelque chose que tu peux résoudre toi-même.

Cette fiche ne demande pas de prompt pour ton LLM — c'est toi qui écris
directement l'entrée.

**Où écrire** : ouvre le fichier `ESCALADES.md`, qui se trouve à la racine
du projet (pas dans le dossier `docs/`). Ajoute une nouvelle entrée tout en
haut de la section "Entrées", en suivant exactement le modèle déjà présent
dans ce même fichier (les parties "Contexte", "Ce qui manque", "Ce que je
voudrais écrire", "Contournement actuel").

**Comment vérifier** : relis ton entrée une fois écrite — est-ce qu'une
personne qui ne connaît pas ta situation comprendrait ton besoin juste en la
lisant ? Si oui, c'est bon.

**Sauvegarder dans Git** :

1. `git add ESCALADES.md`
2. `git commit -m "docs(escalade): resume court de ton besoin"`
3. `git push`

Pas besoin d'attendre une réponse pour continuer à travailler sur autre
chose — ton entrée sera relue périodiquement.

---

## Dépannage — les blocages typiques

### "Les tests ne passent pas, `pytest` affiche des erreurs"

Copie le message d'erreur complet (tout ce qui s'affiche après le mot
"FAILED" ou "Error"), colle-le dans le chat de ton LLM avec la question
"voici l'erreur que j'obtiens en lançant pytest, comment je la corrige ?".
Ne panique pas : dans ce projet, un test qui échoue au milieu du travail est
normal, ça fait partie du processus habituel — ce n'est pas le signe que tu
as fait quelque chose de grave.

### "`pytest` (ou une autre commande) n'est pas reconnu"

Ça arrive souvent sous Windows quand tu n'as pas les droits administrateur :
l'installation se fait quand même, mais dans un dossier que Windows ne
trouve pas automatiquement. Solution simple : au lieu de taper `pytest`
tout seul, tape `python -m pytest` (ça marche aussi pour les autres
commandes, par exemple `python -m pip install ...`). Si ça affiche bien un
résultat, tu peux utiliser `python -m` devant chaque commande de ce genre
dans tout le reste de ce manuel.

### "Je ne sais pas quoi copier-coller, je suis perdue"

Reviens au tableau "Qu'est-ce que tu veux faire aujourd'hui ?" au début de
ce document, trouve la ligne qui ressemble le plus à ta situation, et suis
la fiche indiquée pas à pas, dans l'ordre — sans sauter d'étape.

### "J'ai cassé quelque chose, comment revenir en arrière ?"

Si tu n'as encore rien sauvegardé avec `git commit` (voir Fiche 5), tape
dans le terminal :

```
git checkout -- nom_du_fichier
```

en remplaçant `nom_du_fichier` par le nom exact du fichier à annuler. Cette
commande remet ce fichier précis exactement comme il était à ta dernière
sauvegarde, sans toucher aux autres fichiers. Si tu as un doute sur ce qu'il
faut taper, arrête-toi et demande de l'aide plutôt que d'essayer une
commande que tu ne comprends pas complètement.

### "Comment je vérifie que mon fichier Excel est bien généré correctement ?"

Après avoir lancé `python -m src sync --dir projet_TrainSystem`, ouvre le
fichier `.xlsx` concerné directement dans Excel — il se trouve dans le
dossier `projet_TrainSystem/` — et regarde à l'œil si le changement que tu
voulais est bien visible. C'est la vérification la plus simple et la plus
fiable : si tu vois ce que tu attendais, c'est bon.
