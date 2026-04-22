"""
Panneau de chat IA — interface utilisateur pour l'agent Llama.
"""
from __future__ import annotations
import threading
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTextEdit,
    QLineEdit, QPushButton, QComboBox, QScrollArea, QFrame, QSizePolicy
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QObject
from PyQt6.QtGui import QTextCursor, QFont


class AgentWorker(QObject):
    """Runs the agent in a background thread and emits signals back to UI."""
    token_ready = pyqtSignal(str)
    tool_called = pyqtSignal(str, dict)
    tool_done = pyqtSignal(str, str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, agent, message: str):
        super().__init__()
        self.agent = agent
        self.message = message

    def run(self):
        try:
            result = self.agent.chat(
                self.message,
                on_tool_call=lambda name, args: self.tool_called.emit(name, args),
                on_tool_result=lambda name, res: self.tool_done.emit(name, res),
            )
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


class AIChatPanel(QWidget):
    """Chat interface for the Llama agent embedded in AuditPro."""

    def __init__(self, agent, parent=None):
        super().__init__(parent)
        self.agent = agent
        self.setObjectName("AIChatPanel")
        self._thread: QThread | None = None
        self._worker: AgentWorker | None = None
        self._models_loaded = False
        self._build_ui()
        self._append_system("Agent IA initialisé. Vérification au premier envoi.")

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        # ── Header ────────────────────────────────────
        header = QHBoxLayout()
        title = QLabel("Agent IA (Llama)")
        title.setObjectName("AssistantTitle")
        header.addWidget(title)
        header.addStretch()

        self.model_combo = QComboBox()
        self.model_combo.setFixedWidth(160)
        self.model_combo.setToolTip("Modèle Ollama à utiliser")
        self.model_combo.addItem(self.agent.model)
        self.model_combo.currentTextChanged.connect(self._on_model_changed)
        header.addWidget(self.model_combo)

        reset_btn = QPushButton("Réinitialiser")
        reset_btn.setFixedWidth(90)
        reset_btn.clicked.connect(self._reset_chat)
        header.addWidget(reset_btn)

        layout.addLayout(header)

        # ── Chat display ──────────────────────────────
        self.chat_display = QTextEdit()
        self.chat_display.setReadOnly(True)
        self.chat_display.setObjectName("ChatDisplay")
        self.chat_display.setFont(QFont("Segoe UI", 9))
        layout.addWidget(self.chat_display, 1)

        # ── Status label ──────────────────────────────
        self.status_label = QLabel("")
        self.status_label.setObjectName("AssistantMessage")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        # ── Input row ─────────────────────────────────
        input_row = QHBoxLayout()
        self.input_field = QLineEdit()
        self.input_field.setPlaceholderText("Posez une question ou donnez une instruction à l'agent...")
        self.input_field.returnPressed.connect(self._send)
        input_row.addWidget(self.input_field)

        self.send_btn = QPushButton("Envoyer")
        self.send_btn.setFixedWidth(80)
        self.send_btn.clicked.connect(self._send)
        input_row.addWidget(self.send_btn)

        layout.addLayout(input_row)

    def _check_agent(self):
        if not self.agent.available:
            self._append_system(
                "⚠️  La bibliothèque <b>ollama</b> n'est pas installée.<br>"
                "Installez-la avec : <code>pip install ollama</code><br>"
                "Ensuite, installez Ollama depuis <b>ollama.com</b> et lancez :<br>"
                "<code>ollama pull llama3.2</code>"
            )
            self.send_btn.setEnabled(False)
            self.input_field.setEnabled(False)
        else:
            self._append_system(
                "Agent IA prêt. Vous pouvez lui demander d'exécuter des modules AuditPro,<br>"
                "analyser des fichiers, ou répondre à vos questions d'audit."
            )

    def _refresh_models(self):
        self.model_combo.blockSignals(True)
        self.model_combo.clear()
        models = self.agent.list_models() if self.agent.available else []
        if models:
            self.model_combo.addItems(models)
            current = self.agent.model
            idx = self.model_combo.findText(current)
            if idx >= 0:
                self.model_combo.setCurrentIndex(idx)
        else:
            self.model_combo.addItem(self.agent.model)
        self.model_combo.blockSignals(False)

    def _on_model_changed(self, model_name: str):
        if model_name:
            self.agent.model = model_name

    def _reset_chat(self):
        self.agent.reset()
        self.chat_display.clear()
        self._append_system("Conversation réinitialisée.")

    def _send(self):
        text = self.input_field.text().strip()
        if not text:
            return
        if self._thread and self._thread.isRunning():
            return

        if not self.agent.available:
            self._append_system(
                "⚠️ Ollama indisponible. Lancez Ollama (ou `ollama serve`) puis réessayez."
            )
            return

        if not self._models_loaded:
            self._refresh_models()
            self._models_loaded = True

        self.input_field.clear()
        self._append_user(text)
        self._set_busy(True)

        self._worker = AgentWorker(self.agent, text)
        self._thread = QThread()
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.tool_called.connect(self._on_tool_called)
        self._worker.tool_done.connect(self._on_tool_done)
        self._worker.finished.connect(self._on_finished)
        self._worker.error.connect(self._on_error)
        self._worker.finished.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)
        self._thread.finished.connect(lambda: self._set_busy(False))

        self._thread.start()

    def _on_tool_called(self, name: str, args: dict):
        self.status_label.setText(f"⚙️  Exécution : {name}...")
        self._append_system(f"🔧 Appel outil : <b>{name}</b>")

    def _on_tool_done(self, name: str, result_json: str):
        import json
        try:
            r = json.loads(result_json)
            if r.get("success"):
                out = r.get("output_path", "")
                stats = r.get("stats", {})
                summary = f"✅ <b>{name}</b> terminé."
                if out:
                    summary += f"<br>📄 Fichier : {out}"
                if stats:
                    for k, v in list(stats.items())[:3]:
                        summary += f"<br>• {k}: {v}"
            else:
                errs = r.get("errors") or [r.get("error", "Erreur inconnue")]
                summary = f"❌ <b>{name}</b> échoué : {'; '.join(str(e) for e in errs)}"
        except Exception:
            summary = f"Résultat outil {name} : {result_json[:200]}"
        self._append_system(summary)

    def _on_finished(self, response: str):
        self.status_label.setText("")
        self._append_assistant(response)

    def _on_error(self, error: str):
        self.status_label.setText("")
        self._append_system(f"❌ Erreur agent : {error}")

    def _set_busy(self, busy: bool):
        self.send_btn.setEnabled(not busy)
        self.input_field.setEnabled(not busy)
        if busy:
            self.status_label.setText("⏳ L'agent réfléchit...")

    def _append_user(self, text: str):
        self.chat_display.append(
            f'<p style="color:#c084fc;font-weight:bold;">Vous</p>'
            f'<p style="color:#e2e8f0;margin-left:8px;">{self._escape(text)}</p>'
        )
        self._scroll_bottom()

    def _append_assistant(self, text: str):
        html = self._escape(text).replace("\n", "<br>")
        self.chat_display.append(
            f'<p style="color:#34d399;font-weight:bold;">Agent IA</p>'
            f'<p style="color:#e2e8f0;margin-left:8px;">{html}</p>'
        )
        self._scroll_bottom()

    def _append_system(self, html: str):
        self.chat_display.append(
            f'<p style="color:#94a3b8;font-style:italic;font-size:8pt;">{html}</p>'
        )
        self._scroll_bottom()

    def _scroll_bottom(self):
        cursor = self.chat_display.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        self.chat_display.setTextCursor(cursor)

    @staticmethod
    def _escape(text: str) -> str:
        return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
