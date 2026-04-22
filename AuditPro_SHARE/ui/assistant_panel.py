"""
Panneau assistant contextuel (côté droit).
Affiche l'aide, les détections, et l'historique.
"""
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QListWidget, QListWidgetItem,
    QScrollArea, QFrame
)
from PyQt6.QtCore import Qt
from core.history import HistoryManager


class AssistantPanel(QWidget):
    """Panneau latéral d'assistance contextuelle."""

    def __init__(self, history_manager: HistoryManager, parent=None):
        super().__init__(parent)
        self.setObjectName("AssistantPanel")
        self.history_mgr = history_manager
        self._build_ui()
        self._refresh_history()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # ── Titre ─────────────────────────────────────
        title = QLabel("Assistant")
        title.setObjectName("AssistantTitle")
        layout.addWidget(title)

        # ── Zone de message contextuel ────────────────
        self.message = QLabel(
            "Sélectionnez un module dans le menu\n"
            "ou glissez-déposez un fichier Excel."
        )
        self.message.setObjectName("AssistantMessage")
        self.message.setWordWrap(True)
        layout.addWidget(self.message)

        # ── Détection ─────────────────────────────────
        self.detection_label = QLabel("")
        self.detection_label.setObjectName("AssistantMessage")
        self.detection_label.setWordWrap(True)
        self.detection_label.setVisible(False)
        layout.addWidget(self.detection_label)

        # ── Séparateur ────────────────────────────────
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color: #E5E7EB; margin: 8px 16px;")
        layout.addWidget(sep)

        # ── Aide module ───────────────────────────────
        help_section = QLabel("AIDE MODULE")
        help_section.setObjectName("AssistantSection")
        layout.addWidget(help_section)

        self.help_text = QLabel("Aucun module sélectionné.")
        self.help_text.setObjectName("AssistantMessage")
        self.help_text.setWordWrap(True)
        layout.addWidget(self.help_text)

        # ── Spacer ────────────────────────────────────
        layout.addStretch(1)

        # ── Historique ────────────────────────────────
        hist_label = QLabel("HISTORIQUE RÉCENT")
        hist_label.setObjectName("AssistantSection")
        layout.addWidget(hist_label)

        self.history_list = QListWidget()
        self.history_list.setObjectName("HistoryList")
        self.history_list.setMaximumHeight(200)
        layout.addWidget(self.history_list)

    def update_for_module(self, module):
        """Met à jour le panneau quand un module est sélectionné."""
        self.message.setText(
            f"Module : {module.name}\n\n"
            f"{module.description}"
        )
        help_txt = getattr(module, 'help_text', '') or "Pas d'aide détaillée disponible."
        self.help_text.setText(help_txt)
        self.detection_label.setVisible(False)

    def show_detection(self, results: list):
        """Affiche les résultats de détection automatique."""
        if not results:
            self.detection_label.setText("Aucun module détecté pour ce fichier.")
            self.detection_label.setVisible(True)
            return

        lines = ["Détection automatique :\n"]
        for r in results[:3]:
            mod = r.get("module")
            if mod is None or not hasattr(mod, "name"):
                continue
            score = r.get("score", 0)
            pct = int(score * 100)
            icon = "🟢" if pct >= 70 else "🟡" if pct >= 40 else "🔴"
            lines.append(f"{icon} {mod.name} — {pct}%")
            if r.get("matched_keywords"):
                kws = ", ".join(r["matched_keywords"][:5])
                lines.append(f"   Colonnes trouvées : {kws}")

        if len(lines) == 1:
            self.detection_label.setText("Aucun module détecté pour ce fichier.")
        else:
            self.detection_label.setText("\n".join(lines))
        self.detection_label.setVisible(True)

    def show_message(self, text: str):
        self.message.setText(text)

    def _refresh_history(self):
        self.history_list.clear()
        for entry in self.history_mgr.get_recent(10):
            ts = entry.get("timestamp", "")[:16].replace("T", " ")
            mod = entry.get("module", "?")
            icon = "✓" if entry.get("success") else "✗"
            item = QListWidgetItem(f"{icon}  {ts}  {mod}")
            self.history_list.addItem(item)

    def add_to_history(self, module_name: str, input_file: str,
                       output_file: str, success: bool = True, stats: dict = None):
        self.history_mgr.add(module_name, input_file, output_file,
                             success=success, stats=stats)
        self._refresh_history()
