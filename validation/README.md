# Validation du Moteur ExoSync — pour le testeur

Bienvenue. Ce dossier contient un **parcours de validation pas-à-pas** du noyau d'ExoSync
(moteur + métalangage MXL sur de vrais fichiers Excel). Tu peux le suivre **sans connaître
le projet**.

## Mise en route (une seule fois)

```bash
# Depuis la racine du dépôt SysEng
pip install -r requirements.txt
```

## Comment ça marche

1. Ouvre **`Parcours-Validation-Moteur.pdf`** (ou le `.md`) : c'est le mode d'emploi détaillé,
   marche par marche (0 à 5).
2. Ouvre **`Validation-Moteur-ExoSync.xlsx`** : c'est ta feuille de suivi. Pour chaque étape,
   note le résultat obtenu et coche le statut (OK / KO).
3. Suis les marches **dans l'ordre**. Ne passe à la suivante que lorsque la précédente est OK.

## Les outils fournis

| Fichier | Rôle |
|---------|------|
| `Parcours-Validation-Moteur.pdf` / `.md` | Le mode d'emploi (à suivre) |
| `Validation-Moteur-ExoSync.xlsx` | La feuille de suivi (à remplir) |
| `../scripts/valider_un.py` | Lanceur : synchronise UN fichier Excel par chemin |
| `CORRIGE_marches_1-4.py` | **Corrigé** : régénère les fichiers de test si tu bloques |

## Commandes utiles

```bash
python -m src store clear                              # repartir d'un store vide
python scripts/valider_un.py validation/mon_fichier.xlsx   # tester un fichier fait main
python -m src status                                   # voir le contenu du store
```

> Les fichiers Excel que tu crées pendant les tests ne sont pas versionnés (`.gitignore`).
> En cas de blocage sur les marches 1 à 4 : `python validation/CORRIGE_marches_1-4.py`
> régénère les fichiers attendus dans `validation/`.
