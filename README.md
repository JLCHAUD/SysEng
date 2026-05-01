# ExoSync — SysEng

Système de synchronisation de données à travers un écosystème de fichiers Excel.
Chaque fichier porte son propre contrat de données via une feuille `_Manifeste` (métalangage MXL).

---

## Prérequis

- **Python 3.11 ou supérieur** — [python.org/downloads](https://www.python.org/downloads/)
- **Git** — [git-scm.com](https://git-scm.com/)

Vérifier les versions :
```bash
python --version   # doit afficher 3.11+
git --version
```

---

## Installation

### 1. Cloner le dépôt

```bash
git clone https://github.com/JLCHAUD/SysEng.git
cd SysEng
```

### 2. Créer un environnement virtuel

**Windows (PowerShell) :**
```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

**Windows (CMD) :**
```cmd
python -m venv .venv
.venv\Scripts\activate.bat
```

**Linux / macOS :**
```bash
python -m venv .venv
source .venv/bin/activate
```

### 3. Installer les dépendances

```bash
pip install -r requirements.txt
```

### 4. Vérifier l'installation

```bash
python -m pytest -q
```

✅ Résultat attendu : `315 passed` (aucun fichier Excel requis pour les tests).

---

## Utilisation

Toutes les commandes se lancent depuis la racine du dépôt avec l'environnement virtuel activé.

### Commandes principales

```bash
# Aide générale
python -m src --help

# Synchroniser tous les fichiers du registre
python -m src sync

# Synchroniser un ou plusieurs fichiers spécifiques
python -m src sync --id UO-001 UO-002

# Générer les fichiers Excel des UO depuis les templates
python -m src generate

# Diagnostic de l'écosystème (santé, ownership, cohérence)
python -m src doctor

# Afficher le graphe de dépendances + owners
python -m src lineage

# Afficher le contenu du store central
python -m src status
```

### Commandes avancées

```bash
# Historique des synchronisations
python -m src history
python -m src history --last 5
python -m src history --compare

# Store central (lecture/écriture)
python -m src store get uo.UO-001.avancement
python -m src store set uo.UO-001.avancement 75.0

# Générer la documentation HTML
python -m src doc
```

---

## Structure du projet

```
SysEng/
├── src/                     # Code source
│   ├── __main__.py          # Point d'entrée (python -m src)
│   ├── cli.py               # Toutes les commandes CLI
│   ├── parser.py            # Parser AST du métalangage MXL
│   ├── executor.py          # Moteur d'exécution (PULL/COMPUTE/PUSH/BIND/NOTIFY)
│   ├── store.py             # Store central JSON
│   ├── sync.py              # Orchestrateur de synchronisation
│   ├── ecosystem.py         # Exomap — graphe de dépendances
│   ├── history.py           # Historique des runs et snapshots
│   ├── security.py          # Validation clés, hashes SHA256
│   ├── config_loader.py     # Chargement des configs + validate_owner_roles()
│   ├── models.py            # Dataclasses métier
│   └── generators/          # Générateurs de fichiers Excel
├── config/
│   ├── registre.json        # Liste des fichiers à synchroniser (avec owner_id)
│   ├── acteurs.json         # Profils des utilisateurs (rôles, emails)
│   ├── file_types.yaml      # Types de fichiers avec owner_role attendu
│   ├── uo_instances.json    # Instances UO (avec owner_id)
│   └── templates/           # Templates MXL pour la génération
├── tests/                   # 315 tests automatisés
├── output/                  # Fichiers générés (ignoré par git)
├── referentiels/            # Fichiers Excel de référence (ignoré par git)
└── requirements.txt
```

---

## Fichiers Excel

Les fichiers Excel (`.xlsx`) ne sont **pas inclus dans le dépôt** (exclus par `.gitignore`).

**Pour démarrer sur un nouveau PC :**

1. Générer les fichiers UO depuis les templates :
   ```bash
   python -m src generate
   ```
   → Crée les fichiers dans `output/UOs/` et `output/cockpits/`

2. Pour les fichiers de référence (`referentiels/`), les copier manuellement depuis l'autre machine ou les recréer.

---

## Variables d'environnement (optionnel)

Pour les notifications email (instruction `NOTIFY email`) :

```bash
# Créer un fichier .env à la racine (non versionné)
SMTP_HOST=smtp.example.com
SMTP_PORT=587
SMTP_USER=user@example.com
SMTP_PASSWORD=motdepasse
SMTP_FROM=exosync@example.com
SMTP_TO=destinataire@example.com
```

---

## Dépannage

**`ModuleNotFoundError: No module named 'src'`**
→ Lancer les commandes depuis la racine du dépôt (`SysEng/`), pas depuis un sous-dossier.

**`python -m pytest` : permission refusée sur `.venv\Scripts\Activate.ps1` (Windows)**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**Encodage UTF-8 sur Windows**
→ Le `__main__.py` force UTF-8 automatiquement. Si des caractères s'affichent mal dans le terminal :
```powershell
chcp 65001
```

---

## Tests

```bash
# Tous les tests
python -m pytest -q

# Avec couverture de code
python -m pytest --cov=src --cov-report=term

# Un fichier de tests spécifique
python -m pytest tests/test_owner_roles.py -v
```

---

## CI/CD

GitHub Actions tourne automatiquement à chaque push sur Python 3.11, 3.12 et 3.13.
Voir `.github/workflows/ci.yml`.
