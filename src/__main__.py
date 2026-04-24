"""Point d'entrée pour `python -m src`."""
import sys

# Forcer UTF-8 sur Windows (cp1252 par défaut)
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from src.cli import main

sys.exit(main())
