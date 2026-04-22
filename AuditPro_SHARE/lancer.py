#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script d'installation et démarrage AuditPro
Vérifie les dépendances et lance l'application
"""
import subprocess
import sys
import os
import importlib.util
from pathlib import Path

# Ajouter le répertoire courant au path pour les imports
sys.path.insert(0, str(Path(__file__).parent))


def check_python():
    """Vérifier la version Python."""
    version = sys.version_info
    if version.major < 3 or (version.major == 3 and version.minor < 9):
        print("[x] Python 3.9+ requis")
        return False
    print(f"[OK] Python {version.major}.{version.minor}.{version.micro}")
    return True


def check_requirements():
    """Vérifier les dépendances essentielles sans lancer de téléchargement pip."""
    required_packages = ["PyQt6", "pandas", "openpyxl", "numpy"]
    missing = [pkg for pkg in required_packages if importlib.util.find_spec(pkg) is None]

    if not missing:
        print("[OK] Dépendances essentielles présentes")
        return True

    print("[x] Dépendances manquantes : " + ", ".join(missing))
    print("    Installez-les une seule fois avec :")
    print("    python -m pip install -r requirements.txt")
    return False


def verify_app():
    """Vérifier que l'app peut démarrer."""
    try:
        print("Vérification de l'application...")
        from core.config import APP_NAME, APP_VERSION
        from core.module_registry import ModuleRegistry

        print(f"[OK] {APP_NAME} v{APP_VERSION}")

        registry = ModuleRegistry()
        modules = registry.get_all()
        print(f"[OK] {len(modules)} modules détectés")
        return True
    except Exception as e:
        print(f"[x] Erreur: {e}")
        return False


def main():
    """Flux principal."""
    print("=" * 42)
    print("   AUDITPRO v1.0.0 - DÉMARRAGE")
    print("=" * 42 + "\n")

    # 1. Vérifier Python
    if not check_python():
        sys.exit(1)

    # 2. Vérifier dépendances (sans installation automatique)
    if not check_requirements():
        sys.exit(1)

    # 3. Vérifier l'app
    if not verify_app():
        sys.exit(1)

    # 4. Lancer l'app
    print("\nDémarrage de l'application...\n")
    try:
        from ui.main_window import MainWindow
        from PyQt6.QtWidgets import QApplication

        app = QApplication(sys.argv)
        window = MainWindow()
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        print(f"[x] Erreur au démarrage: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
