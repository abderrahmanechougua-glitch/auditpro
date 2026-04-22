"""
Dialogue de gestion des profils clients.
"""
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QListWidget, QListWidgetItem, QMessageBox
)
from PyQt6.QtCore import pyqtSignal
from core.profiles import ProfileManager


class ProfileDialog(QDialog):
    """Fenêtre de sélection/création de profils clients."""

    profile_selected = pyqtSignal(str)

    def __init__(self, profile_mgr: ProfileManager, parent=None):
        super().__init__(parent)
        self.profile_mgr = profile_mgr
        self.setWindowTitle("Profils clients")
        self.setMinimumSize(400, 350)
        self._build_ui()
        self._refresh()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        layout.addWidget(QLabel("Sélectionnez ou créez un profil client :"))

        self.profile_list = QListWidget()
        self.profile_list.itemDoubleClicked.connect(self._on_select)
        layout.addWidget(self.profile_list)

        # Création
        row = QHBoxLayout()
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Nom du nouveau profil...")
        row.addWidget(self.name_input, 1)

        btn_add = QPushButton("Créer")
        btn_add.setObjectName("PrimaryButton")
        btn_add.clicked.connect(self._on_create)
        row.addWidget(btn_add)
        layout.addLayout(row)

        # Actions
        row2 = QHBoxLayout()
        btn_select = QPushButton("Sélectionner")
        btn_select.setObjectName("PrimaryButton")
        btn_select.clicked.connect(self._on_select)
        row2.addWidget(btn_select)

        btn_delete = QPushButton("Supprimer")
        btn_delete.setObjectName("DangerButton")
        btn_delete.clicked.connect(self._on_delete)
        row2.addWidget(btn_delete)

        row2.addStretch()

        btn_close = QPushButton("Fermer")
        btn_close.setObjectName("SecondaryButton")
        btn_close.clicked.connect(self.close)
        row2.addWidget(btn_close)

        layout.addLayout(row2)

    def _refresh(self):
        self.profile_list.clear()
        for name in self.profile_mgr.list_profiles():
            self.profile_list.addItem(name)

    def _on_create(self):
        name = self.name_input.text().strip()
        if not name:
            return
        if name in self.profile_mgr.list_profiles():
            QMessageBox.warning(self, "Doublon", f"Le profil '{name}' existe déjà.")
            return
        self.profile_mgr.save(name, {})
        self.name_input.clear()
        self._refresh()

    def _on_select(self):
        item = self.profile_list.currentItem()
        if item:
            self.profile_selected.emit(item.text())
            self.close()

    def _on_delete(self):
        item = self.profile_list.currentItem()
        if not item:
            return
        reply = QMessageBox.question(
            self, "Supprimer",
            f"Supprimer le profil '{item.text()}' ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.profile_mgr.delete(item.text())
            self._refresh()
