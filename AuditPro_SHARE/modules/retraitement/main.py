"""Compatibilite retroactive pour le module retraitement.

Ce fichier permet deux usages sans erreur d'import relatif:
1. Import package normal: from modules.retraitement.main import ...
2. Execution directe: python modules/retraitement/main.py

La logique metier reste dans processor.py.
"""

from pathlib import Path
import sys


def _bootstrap_for_direct_execution() -> None:
    """Ajoute AuditPro_SHARE au path si ce fichier est execute directement."""
    app_root = Path(__file__).resolve().parents[2]
    app_root_str = str(app_root)
    if app_root_str not in sys.path:
        sys.path.insert(0, app_root_str)


if __package__ in (None, ""):
    _bootstrap_for_direct_execution()
    from modules.retraitement.processor import (  # type: ignore
        IntelligentRetraitement,
        process_file,
        process_files,
    )
else:
    from .processor import IntelligentRetraitement, process_file, process_files


__all__ = ["IntelligentRetraitement", "process_file", "process_files"]


if __name__ == "__main__":
    print("[OK] modules.retraitement.main charge correctement.")
    print("[INFO] Utilisez plutot: from modules.retraitement import process_file")
