"""Launch AuditPro from the editable _internal runtime tree.

This bypasses the frozen executable code archive so local .py edits are
immediately reflected when debugging UI/modules.
"""

from __future__ import annotations

import sys
from pathlib import Path

from PyQt6.QtWidgets import QApplication


def main() -> int:
    root = Path(__file__).resolve().parent
    internal = root / "dist" / "AuditPro" / "_internal"
    if not internal.exists():
        print(f"[ERREUR] Runtime introuvable: {internal}")
        return 1

    sys.path.insert(0, str(internal))

    from ui.main_window import MainWindow  # imported after sys.path setup

    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
