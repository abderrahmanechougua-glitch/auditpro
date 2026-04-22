"""
Panneau latéral gauche — liste des modules disponibles.
"""
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, QButtonGroup,
    QListWidget, QFrame
)
from PyQt6.QtCore import pyqtSignal, Qt


class ModulePanel(QWidget):
    """Sidebar avec la liste des modules et l'identité de l'app."""

    module_selected = pyqtSignal(str)  # Émet le nom du module cliqué

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("SideBar")
        self._buttons: dict[str, QPushButton] = {}
        self._btn_group = QButtonGroup(self)
        self._btn_group.setExclusive(True)
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # ── Logo / Titre ──────────────────────────────
        title = QLabel("AuditPro")
        title.setObjectName("AppTitle")
        layout.addWidget(title)

        subtitle = QLabel("Assistant d'audit intelligent")
        subtitle.setObjectName("AppSubtitle")
        layout.addWidget(subtitle)

        # ── Séparateur ────────────────────────────────
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color: #3A1F54; margin: 0 16px;")
        layout.addWidget(sep)

        # ── Section label ─────────────────────────────
        section = QLabel("MODULES")
        section.setObjectName("SectionLabel")
        layout.addWidget(section)

        # ── Container pour les boutons modules ────────
        self.buttons_layout = QVBoxLayout()
        self.buttons_layout.setContentsMargins(8, 0, 8, 0)
        self.buttons_layout.setSpacing(2)
        layout.addLayout(self.buttons_layout)

        # ── Spacer ────────────────────────────────────
        layout.addStretch(1)

        # ── Version ───────────────────────────────────
        version_label = QLabel("v1.0.0")
        version_label.setObjectName("AppSubtitle")
        version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(version_label)

    def set_modules(self, modules: dict):
        """Peuple la sidebar avec les modules découverts."""
        # Nettoyer les boutons existants
        for btn in self._buttons.values():
            self.buttons_layout.removeWidget(btn)
            btn.deleteLater()
        self._buttons.clear()

        # Catégoriser
        categories = {}
        for mod in modules.values():
            cat = getattr(mod, 'category', 'Général')
            categories.setdefault(cat, []).append(mod)

        for cat_name, mods in sorted(categories.items()):
            if len(categories) > 1:
                lbl = QLabel(cat_name.upper())
                lbl.setObjectName("SectionLabel")
                self.buttons_layout.addWidget(lbl)

            for mod in mods:
                btn = QPushButton(f"  {mod.name}")
                btn.setObjectName("ModuleButton")
                btn.setCheckable(True)
                btn.setToolTip(mod.description)
                btn.clicked.connect(lambda checked, n=mod.name: self.module_selected.emit(n))
                self._btn_group.addButton(btn)
                self.buttons_layout.addWidget(btn)
                self._buttons[mod.name] = btn

    def select_module(self, name: str):
        """Sélectionne visuellement un module."""
        if name in self._buttons:
            self._buttons[name].setChecked(True)
