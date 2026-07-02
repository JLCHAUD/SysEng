# Manuel Collègue Excel — Plan d'implémentation

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Produire `manuel_collegue_excel.docx` à la racine du dépôt SysEng — un guide humain, très pédagogique, qui dit à la collègue exactement quel fichier ouvrir, quel prompt coller dans son LLM, où mettre la réponse, comment vérifier, comment sauvegarder, pour chacune des 6 situations qu'elle rencontrera.

**Architecture:** Rédaction d'un fichier markdown source complet (`docs/manuel_collegue_excel_source.md`) contenant tout le texte final, puis conversion en `.docx` via le skill `docx`, avec mise en forme Word (titres, tableaux, encadrés pour les prompts à copier-coller).

**Tech Stack:** Markdown pour la rédaction, skill `docx` (python-docx) pour la génération finale.

**Spec de référence :** `docs/superpowers/specs/2026-07-02-manuel-collegue-excel-design.md`

---

## Contexte technique à lire avant de commencer

- Le public cible a des notions de base (terminal oui, git/Python flous) et un
  LLM en **chat simple sans accès aux fichiers** — chaque fiche doit dire
  explicitement "ouvre le fichier X, Ctrl+A, Ctrl+C, colle dans le chat".
- Ce manuel **orchestre** le package déjà livré en CONV-14
  (`docs/onboarding-excel/*.md` + `ESCALADES.md` à la racine) — il ne le
  remplace pas et ne duplique pas son contenu technique (la référence MXL
  complète reste dans `02-Langage-MXL.md`, pas recopiée ici).
- Fichiers de référence CONV-14 déjà existants dans le dépôt, à ne pas modifier :
  `docs/onboarding-excel/00-DEMARRAGE.md`, `01-Vue-Ensemble-Projet.md`,
  `02-Langage-MXL.md`, `03-Design-System-XD.md`, `04-Territoire-Et-Conventions.md`,
  `05-Escalade.md`, et `ESCALADES.md` à la racine.
- Commande de sync du projet : `python -m src sync --dir projet_TrainSystem`
- Commande de test : `pytest`

---

## Fichiers concernés

| Action | Fichier |
|--------|---------|
| Créer | `docs/manuel_collegue_excel_source.md` (source markdown, texte complet) |
| Créer | `manuel_collegue_excel.docx` (racine du dépôt, livrable final) |

---

### Task 0 : Créer la branche de travail

**Files:** aucun

- [ ] **Step 1 : Créer et basculer sur la branche**

```bash
cd "C:\Users\fabie\Documents\JLC\Python\SysEng"
git checkout master
git pull
git checkout -b feature/manuel-collegue-excel
```

---

### Task 1 : Introduction + sommaire de décision (`docs/manuel_collegue_excel_source.md`)

**Files:**
- Create: `docs/manuel_collegue_excel_source.md`

- [ ] **Step 1 : Écrire le début du fichier**

```markdown
# Manuel pour toi — travailler sur les fichiers Excel du projet ExoSync

Ce document est écrit pour toi, pas pour ton LLM. Il t'explique, situation
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
```

- [ ] **Step 2 : Vérifier**

Relire : aucun terme non défini avant sa première explication (terminal,
pytest sont introduits avec leur rôle). Le tableau de décision couvre bien
les 6 fiches + la section dépannage prévues dans la spec.

- [ ] **Step 3 : Commit**

```bash
git add docs/manuel_collegue_excel_source.md
git commit -m "docs(manuel): introduction + sommaire de decision"
```

---

### Task 2 : Fiches 1 à 3 (append à `docs/manuel_collegue_excel_source.md`)

**Files:**
- Modify: `docs/manuel_collegue_excel_source.md` (ajouter à la fin)

- [ ] **Step 1 : Ajouter les 3 fiches**

```markdown
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
```

- [ ] **Step 2 : Vérifier**

Confirmer que les 3 fiches respectent chacune la structure fixe à 6 points
(quoi faire / fichier à copier / prompt / où mettre la réponse / vérifier /
sauvegarder), que les prompts sont du texte concret prêt à copier-coller
(pas une description abstraite), et que les fichiers référencés
(`01-Vue-Ensemble-Projet.md`, `02-Langage-MXL.md`, `03-Design-System-XD.md`)
existent bien dans `docs/onboarding-excel/`.

- [ ] **Step 3 : Commit**

```bash
git add docs/manuel_collegue_excel_source.md
git commit -m "docs(manuel): fiches 1-3 (comprendre, manifeste, mise en forme)"
```

---

### Task 3 : Fiches 4 à 6 (append à `docs/manuel_collegue_excel_source.md`)

**Files:**
- Modify: `docs/manuel_collegue_excel_source.md` (ajouter à la fin)

- [ ] **Step 1 : Ajouter les 3 fiches**

```markdown
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
   explique ce que tu as changé. Écris ce message toi-même, en une phrase
   simple (par exemple : `git commit -m "ajoute colonne priorite dans le
   cockpit"`).

4. `git push`
   Ça envoie ta sauvegarde vers le serveur partagé (GitHub), pour que le
   reste de l'équipe puisse la voir.

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
2. `git commit -m "besoin : resume court de ton besoin"`
3. `git push`

Pas besoin d'attendre une réponse pour continuer à travailler sur autre
chose — ton entrée sera relue périodiquement.
```

- [ ] **Step 2 : Vérifier**

Confirmer que la Fiche 4 référence bien les 3 bons fichiers dans le bon
ordre, que la Fiche 5 explique chaque commande Git en langage clair (pas
juste la commande brute), que la Fiche 6 pointe vers `ESCALADES.md` à la
racine (pas vers `docs/onboarding-excel/05-Escalade.md` pour l'écriture de
l'entrée elle-même — ce dernier fichier sert seulement à décider si c'est
une escalade).

- [ ] **Step 3 : Commit**

```bash
git add docs/manuel_collegue_excel_source.md
git commit -m "docs(manuel): fiches 4-6 (generateur, git, escalade)"
```

---

### Task 4 : Section Dépannage / FAQ (append à `docs/manuel_collegue_excel_source.md`)

**Files:**
- Modify: `docs/manuel_collegue_excel_source.md` (ajouter à la fin)

- [ ] **Step 1 : Ajouter la section**

```markdown
---

## Dépannage — les blocages typiques

### "Les tests ne passent pas, `pytest` affiche des erreurs"

Copie le message d'erreur complet (tout ce qui s'affiche après le mot
"FAILED" ou "Error"), colle-le dans le chat de ton LLM avec la question
"voici l'erreur que j'obtiens en lançant pytest, comment je la corrige ?".
Ne panique pas : dans ce projet, un test qui échoue au milieu du travail est
normal, ça fait partie du processus habituel — ce n'est pas le signe que tu
as fait quelque chose de grave.

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
```

- [ ] **Step 2 : Vérifier**

Confirmer que les 4 items de dépannage prévus par la spec sont bien tous
présents (tests qui échouent / je suis perdue / revenir en arrière /
vérifier un Excel généré), et que le ton reste rassurant partout (aucune
formulation qui sous-entend que l'utilisatrice a fait une erreur grave).

- [ ] **Step 3 : Commit**

```bash
git add docs/manuel_collegue_excel_source.md
git commit -m "docs(manuel): section depannage / FAQ"
```

---

### Task 5 : Générer `manuel_collegue_excel.docx` depuis la source markdown

**Files:**
- Create: `manuel_collegue_excel.docx` (racine du dépôt)
- Read: `docs/manuel_collegue_excel_source.md`

- [ ] **Step 1 : Invoquer le skill `docx`**

Utilise le skill `docx` (disponible dans cet environnement) pour convertir
`docs/manuel_collegue_excel_source.md` en un document Word bien mis en
forme. Suis les instructions du skill pour la méthode de génération
(probablement un script Python utilisant `python-docx`).

Applique cette mise en forme :
- Le titre principal (`# Manuel pour toi...`) → style Titre 1
- Les titres de fiches (`## Fiche N — ...`) et les autres `##` → style
  Titre 2, pour qu'un sommaire automatique Word (table des matières) puisse
  s'appuyer dessus
- Les sous-titres en gras dans chaque fiche (`**Ce que tu veux faire**`,
  `**Fichier à ouvrir et copier**`, etc.) → style Titre 3 ou gras avec
  espacement, pour rester visuellement scannable
- Le tableau "Qu'est-ce que tu veux faire aujourd'hui ?" → un vrai tableau
  Word (pas du texte aligné avec des espaces)
- Les blocs "Prompt à coller" (introduits par `>`) → mis en forme dans un
  encadré ou avec un fond grisé distinctif, pour qu'ils sautent aux yeux
  comme "ceci est à copier tel quel"
- Les commandes à taper (blocs de code avec ` ``` `) → police à chasse fixe
  (type Consolas ou Courier New), fond légèrement grisé
- Ajoute une table des matières automatique Word en tout début de document
  (après le titre), générée à partir des styles Titre 1/Titre 2

- [ ] **Step 2 : Générer le fichier**

Produis `manuel_collegue_excel.docx` à la racine du dépôt
(`C:\Users\fabie\Documents\JLC\Python\SysEng\manuel_collegue_excel.docx`).

- [ ] **Step 3 : Vérifier le fichier généré**

Ouvre ou inspecte le `.docx` généré (via le skill `docx` ou une lecture
programmatique) et confirme :
- Le document s'ouvre sans erreur
- Les 6 fiches sont présentes avec leurs titres
- Le tableau de décision est un vrai tableau (pas du texte brut)
- Les prompts à copier-coller sont visuellement distincts du reste du texte
- La table des matières liste bien les 6 fiches + la section Dépannage

- [ ] **Step 4 : Commit**

```bash
git add manuel_collegue_excel.docx
git commit -m "docs(manuel): genere manuel_collegue_excel.docx depuis la source"
```

---

### Task 6 : Vérification finale

**Files:** aucun (vérification uniquement)

- [ ] **Step 1 : Vérifier les critères de succès de la spec**

Relire `docs/superpowers/specs/2026-07-02-manuel-collegue-excel-design.md`
section 6 et cocher chaque critère :
- [ ] `manuel_collegue_excel.docx` existe à la racine et est commité
- [ ] Les 6 fiches-recettes sont présentes, chacune avec ses 6 sous-sections
      fixes
- [ ] Chaque prompt de fiche-recette est un texte concret prêt à
      copier-coller (relire chacun des 6 prompts, confirmer qu'aucun n'est
      une description abstraite du type "demande à ton LLM de...")
- [ ] Aucun terme technique n'apparaît sans explication en langage clair à
      sa première occurrence (relecture ciblée : terminal, commit, branche,
      sync, prompt, `_Manifeste`)
- [ ] Les fichiers de contexte cités (`01-Vue-Ensemble-Projet.md` à
      `05-Escalade.md`, `ESCALADES.md`) existent réellement dans le dépôt

- [ ] **Step 2 : Vérifier qu'aucun code n'a été modifié**

```bash
cd "C:\Users\fabie\Documents\JLC\Python\SysEng"
git diff master --stat
```

Attendu : uniquement `docs/manuel_collegue_excel_source.md` et
`manuel_collegue_excel.docx` — aucun fichier `.py`.

- [ ] **Step 3 : Vérifier qu'aucun test n'est cassé**

```bash
pytest -q --tb=short
```

Attendu : même nombre de tests passants qu'avant ce projet (aucune
régression possible, ce projet ne touche aucun fichier `.py`).

- [ ] **Step 4 : Rapport final**

Résumer : liste des fichiers créés, résultat des critères de succès,
résultat de `pytest`, proposer la suite (merge vers `master`, ou attente de
relecture par l'utilisateur).
