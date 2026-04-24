"""
ExoSync — Sécurité & Intégrité (M10)
======================================
Validation des clés store, intégrité des Manifestes, namespaces isolés.

Fonctions publiques :
    validate_store_key(key)          → None ou lève StoreKeyError
    is_valid_store_key(key)          → bool
    compute_manifest_hash(filepath)  → str (sha256 hex)
    check_manifest_integrity(filepath, expected_hash) → bool
    validate_namespace(key, allowed_prefixes) → bool
"""
import hashlib
import re
from pathlib import Path
from typing import List, Optional


# ─── Exceptions ──────────────────────────────────────────────────────────────

class StoreKeyError(ValueError):
    """Clé store invalide (injection, format, namespace interdit)."""


class ManifestIntegrityError(RuntimeError):
    """Hash du Manifeste ne correspond pas au hash attendu."""


# ─── Validation des clés store ────────────────────────────────────────────────

# Clé valide : lettres, chiffres, tirets, underscores et points
# Pas de séquences dangereuses : ../ \x00 espaces etc.
_KEY_PATTERN = re.compile(r'^[a-zA-Z0-9_\-]([a-zA-Z0-9_\-\.]*[a-zA-Z0-9_\-])?$')
_MAX_KEY_LEN = 200
_FORBIDDEN_SEGMENTS = {"", "..", "."}


def is_valid_store_key(key: str) -> bool:
    """
    Retourne True si la clé store est valide.
    Règles :
    - Longueur 1–200 caractères
    - Caractères autorisés : [a-zA-Z0-9_\\-.] uniquement
    - Pas de segments vides, '..' ou '.' entre les points
    - Pas d'espaces, slashes, anti-slashes, ou caractères de contrôle
    """
    if not key or len(key) > _MAX_KEY_LEN:
        return False
    if not _KEY_PATTERN.match(key):
        return False
    segments = key.split(".")
    if any(seg in _FORBIDDEN_SEGMENTS for seg in segments):
        return False
    return True


def validate_store_key(key: str) -> None:
    """
    Valide une clé store. Lève StoreKeyError si invalide.
    Utiliser avant tout get/set sur le store.
    """
    if not is_valid_store_key(key):
        raise StoreKeyError(
            f"Cle store invalide : {key!r}. "
            "Seuls les caracteres [a-zA-Z0-9_\\-.] sont autorises, "
            "sans segments vides ni '..'."
        )


# ─── Namespaces isolés ────────────────────────────────────────────────────────

def validate_namespace(key: str, allowed_prefixes: List[str]) -> bool:
    """
    Vérifie qu'une clé appartient à l'un des namespaces autorisés.

    Args:
        key              : clé store à vérifier
        allowed_prefixes : ex. ["uo.", "projet.", "ref."]

    Returns:
        True si la clé commence par l'un des préfixes autorisés.
    """
    return any(key.startswith(p) for p in allowed_prefixes)


# ─── Intégrité des Manifestes ─────────────────────────────────────────────────

def compute_manifest_hash(filepath: Path) -> str:
    """
    Calcule le hash SHA-256 d'un fichier Excel (contenu binaire brut).

    Returns:
        Chaîne hexadécimale du hash SHA-256.
    """
    sha = hashlib.sha256()
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            sha.update(chunk)
    return sha.hexdigest()


def check_manifest_integrity(filepath: Path, expected_hash: str) -> bool:
    """
    Vérifie que le hash du fichier correspond au hash attendu.

    Returns:
        True si le fichier n'a pas été modifié.
    Raises:
        FileNotFoundError si le fichier est absent.
    """
    if not filepath.exists():
        raise FileNotFoundError(f"Fichier introuvable : {filepath}")
    actual = compute_manifest_hash(filepath)
    return actual == expected_hash


def compute_manifest_hash_from_sheet(ws) -> str:
    """
    Calcule un hash des instructions du Manifeste depuis une feuille openpyxl.
    Utile pour détecter des changements sans relire le binaire xlsx.

    Returns:
        Hash SHA-256 du contenu textuel de la feuille.
    """
    sha = hashlib.sha256()
    for row in ws.iter_rows(values_only=True):
        line = "|".join(str(c) if c is not None else "" for c in row)
        sha.update(line.encode("utf-8"))
    return sha.hexdigest()
