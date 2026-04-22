"""
╔══════════════════════════════════════════════════════════════╗
║                    AUDITPRO v1.0                             ║
║           Assistant d'audit intelligent                      ║
╠══════════════════════════════════════════════════════════════╣
║  Usage : python main.py                                      ║
║                                                              ║
║  Structure :                                                 ║
║    modules/ → Ajouter vos scripts dans chaque sous-dossier   ║
║    data/    → Profils clients et historique (auto-créé)       ║
╚══════════════════════════════════════════════════════════════╝
"""
import sys
import os
from pathlib import Path

# ── Résolution robuste des chemins (Windows / temp dirs) ──────
APP_ROOT = Path(__file__).resolve().parent
os.chdir(APP_ROOT)
sys.path.insert(0, str(APP_ROOT))

from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QIcon

from ui.main_window import MainWindow
from core.config import APP_NAME


def main():
    # High DPI
    os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"

    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    icon_path = APP_ROOT / "resources" / "AuditPro.ico"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))

    # Police par défaut
    font = QFont("Segoe UI", 10)
    app.setFont(font)

    window = MainWindow()
    if icon_path.exists():
        window.setWindowIcon(QIcon(str(icon_path)))
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
