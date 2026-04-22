"""
Workspace — zone centrale de travail.
Contient : zone de drop, inputs dynamiques, aperçu, boutons d'action.
"""
import os
from pathlib import Path
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QDateEdit,
    QProgressBar, QFrame, QMessageBox, QScrollArea, QGridLayout,
    QSizePolicy, QButtonGroup
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QFont
import pandas as pd

from ui.preview_table import PreviewTable
from modules.base_module import BaseModule, ModuleInput


class Workspace(QWidget):
    """Zone centrale de travail."""

    file_loaded = pyqtSignal(str)
    execute_requested = pyqtSignal(dict)
    preview_requested = pyqtSignal(dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("Workspace")
        self.current_module: BaseModule | None = None
        self.input_widgets: dict[str, QWidget] = {}
        self.loaded_files: dict[str, str] = {}
        self._active_etape_id: str = ""
        self._step_buttons: dict[str, QPushButton] = {}
        self._build_ui()

    def _build_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(24, 20, 24, 20)
        main_layout.setSpacing(16)

        # ── Titre du module ───────────────────────────
        self.title = QLabel("Bienvenue dans AuditPro")
        self.title.setObjectName("WorkspaceTitle")
        main_layout.addWidget(self.title)

        self.description = QLabel(
            "Sélectionnez un module dans le menu de gauche pour commencer."
        )
        self.description.setObjectName("WorkspaceDesc")
        self.description.setWordWrap(True)
        main_layout.addWidget(self.description)

        # ── Scroll area pour le contenu ───────────────
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll_content = QWidget()
        self.content_layout = QVBoxLayout(scroll_content)
        self.content_layout.setSpacing(12)
        scroll.setWidget(scroll_content)
        main_layout.addWidget(scroll, 1)


        # ── Sélecteur d'étapes (workflow) ────────────
        self.steps_frame = QFrame()
        self.steps_frame.setObjectName("Card")
        self.steps_layout = QVBoxLayout(self.steps_frame)
        self.steps_layout.setSpacing(6)
        self.steps_layout.setContentsMargins(16, 14, 16, 14)
        self.steps_frame.setVisible(False)
        self.content_layout.addWidget(self.steps_frame)

        # Description de l'étape sélectionnée
        self.step_desc = QLabel("")
        self.step_desc.setObjectName("WorkspaceDesc")
        self.step_desc.setWordWrap(True)
        self.step_desc.setVisible(False)
        self.content_layout.addWidget(self.step_desc)

        # ── Zone d'inputs dynamiques ──────────────────
        self.inputs_frame = QFrame()
        self.inputs_frame.setObjectName("Card")
        self.inputs_layout = QGridLayout(self.inputs_frame)
        self.inputs_layout.setSpacing(10)
        self.inputs_frame.setVisible(False)
        self.content_layout.addWidget(self.inputs_frame)

        # ── Zone de paramètres : masquée définitivement
        self.params_frame = QFrame()
        self.params_frame.setObjectName("Card")
        self.params_layout = QGridLayout(self.params_frame)
        self.params_frame.setVisible(False)  # jamais visible utilisateur
        self.content_layout.addWidget(self.params_frame)

        # ── Zone de sortie ────────────────────────────
        self.output_frame = QFrame()
        self.output_frame.setObjectName("Card")
        output_layout = QVBoxLayout(self.output_frame)
        output_layout.setContentsMargins(16, 12, 16, 12)

        output_title = QLabel("Répertoire de sortie (optionnel)")
        output_title.setStyleSheet("font-weight: bold; font-size: 14px;")
        output_layout.addWidget(output_title)

        output_row = QHBoxLayout()
        self.output_line = QLineEdit()
        self.output_line.setPlaceholderText("Laisser vide pour automatique...")
        self.output_line.setStyleSheet("color: #000000;")
        self.output_line.setReadOnly(True)
        output_row.addWidget(self.output_line, 1)

        self.btn_choose_output = QPushButton("Choisir")
        self.btn_choose_output.setObjectName("SecondaryButton")
        self.btn_choose_output.clicked.connect(self._choose_output_dir)
        output_row.addWidget(self.btn_choose_output)

        output_layout.addLayout(output_row)

        self.output_frame.setVisible(False)  # Masqué définitivement
        self.content_layout.addWidget(self.output_frame)
        self.output_frame.hide()  # Jamais visible

        # ── Aperçu ────────────────────────────────────
        self.preview_label = QLabel("Aperçu des données")
        self.preview_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        self.preview_label.setVisible(False)
        self.content_layout.addWidget(self.preview_label)

        self.preview_table = PreviewTable()
        self.preview_table.setVisible(False)
        self.preview_table.setMaximumHeight(250)
        self.content_layout.addWidget(self.preview_table)

        # ── Boutons d'action ──────────────────────────
        btn_layout = QHBoxLayout()

        self.btn_preview = QPushButton("Aperçu")
        self.btn_preview.setObjectName("SecondaryButton")
        self.btn_preview.setVisible(False)
        self.btn_preview.clicked.connect(self._on_preview)
        btn_layout.addWidget(self.btn_preview)

        self.btn_execute = QPushButton("Exécuter")
        self.btn_execute.setObjectName("PrimaryButton")
        self.btn_execute.setVisible(False)
        self.btn_execute.clicked.connect(self._on_execute)
        btn_layout.addWidget(self.btn_execute)

        btn_layout.addStretch()
        self.content_layout.addLayout(btn_layout)

        # ── Barre de progression ──────────────────────
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        self.content_layout.addWidget(self.progress_bar)

        self.progress_message = QLabel("")
        self.progress_message.setVisible(False)
        self.progress_message.setStyleSheet("color: #6C757D; font-size: 12px;")
        self.content_layout.addWidget(self.progress_message)

        # ── Résultat (avec bouton intégré) ────────────
        self.result_frame = QFrame()
        self.result_frame.setObjectName("Card")
        self.result_layout = QVBoxLayout(self.result_frame)
        self.result_label = QLabel("")
        self.result_label.setWordWrap(True)
        self.result_layout.addWidget(self.result_label)

        # Bouton "Ouvrir" DANS le résultat — toujours visible après exécution
        self.btn_open_output = QPushButton("Ouvrir le résultat")
        self.btn_open_output.setObjectName("PrimaryButton")
        self.btn_open_output.setVisible(False)
        self.btn_open_output.clicked.connect(self._on_open_output)
        self.result_layout.addWidget(self.btn_open_output)

        self.result_frame.setVisible(False)
        self.content_layout.addWidget(self.result_frame)

        self.content_layout.addStretch()

        # ── État interne ──────────────────────────────
        self._output_path = ""

    def set_module(self, module: BaseModule):
        """Configure le workspace pour un module donné."""
        self.current_module = module
        self.title.setText(module.name)
        self.description.setText(module.description)

        # Reset complet
        self._clear_inputs()
        self._step_buttons.clear()
        self._active_etape_id = ""
        self.preview_table.clear_table()
        self.preview_table.setVisible(False)
        self.preview_label.setVisible(False)
        self.result_frame.setVisible(False)
        self.btn_open_output.setVisible(False)
        self.progress_bar.setVisible(False)
        self.progress_message.setVisible(False)
        self.loaded_files.clear()

        # ── Sélecteur d'étapes (si le module en a) ───
        etapes = getattr(module, "ETAPES", None)
        # Chercher aussi dans le module Python importé
        if etapes is None:
            import modules.circularisation.module as circ_mod
            if type(module).__name__ == "CircularisationTiers":
                etapes = circ_mod.ETAPES
        self._build_steps(etapes)

        # La box paramètres peut être visible si le module en a
        self.params_frame.setVisible(False)

        # Afficher les inputs de la 1ère étape (ou du module entier)
        if not etapes:
            required_inputs = module.get_required_inputs()
            if required_inputs:
                self._build_inputs(required_inputs)
                self.inputs_frame.setVisible(True)
            else:
                self.inputs_frame.setVisible(False)

        # Afficher les paramètres du module si disponibles
        params = []
        if hasattr(module, "get_param_schema"):
            params = module.get_param_schema() or []
        if params:
            self._build_params(params)
            self.params_frame.setVisible(True)
        else:
            self.params_frame.setVisible(False)

        self._update_output_dir_display()
        self.output_frame.setVisible(False)  # Toujours masqué
        self.btn_preview.setVisible(True)
        self.btn_execute.setVisible(True)

    # ── Sélecteur d'étapes ────────────────────────────────────

    def _build_steps(self, etapes):
        """Construit les boutons d'étapes si le module en a."""
        # Vider l'ancien sélecteur
        while self.steps_layout.count():
            item = self.steps_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._step_buttons.clear()

        if not etapes:
            self.steps_frame.setVisible(False)
            self.step_desc.setVisible(False)
            return

        # Titre de la section
        lbl_titre = QLabel("Choisissez une étape :")
        lbl_titre.setStyleSheet("font-weight: bold; font-size: 13px; margin-bottom: 6px;")
        self.steps_layout.addWidget(lbl_titre)

        btn_group = QButtonGroup(self)
        btn_group.setExclusive(True)

        for etape in etapes:
            btn = QPushButton(f"  {etape['icon']}.  {etape['label']}")
            btn.setCheckable(True)
            btn.setObjectName("StepButton")
            btn.setMinimumHeight(44)
            btn.setStyleSheet("""
                QPushButton#StepButton {
                    background-color: #F3EEF8;
                    color: #1A1A2E;
                    border: 2px solid #C4B0D8;
                    border-radius: 8px;
                    padding: 8px 16px;
                    text-align: left;
                    font-size: 13px;
                }
                QPushButton#StepButton:hover {
                    background-color: #E0D4EF;
                    border-color: #4B286D;
                }
                QPushButton#StepButton:checked {
                    background-color: #4B286D;
                    color: white;
                    border-color: #4B286D;
                    font-weight: bold;
                }
            """)
            etape_id = etape["id"]
            etape_desc = etape["description"]
            btn.clicked.connect(
                lambda checked, eid=etape_id, edesc=etape_desc:
                    self._on_step_selected(eid, edesc)
            )
            btn_group.addButton(btn)
            self.steps_layout.addWidget(btn)
            self._step_buttons[etape_id] = btn

        self.steps_frame.setVisible(True)

        # Sélectionner la première étape par défaut
        first_id   = etapes[0]["id"]
        first_desc = etapes[0]["description"]
        self._step_buttons[first_id].setChecked(True)
        self._on_step_selected(first_id, first_desc)

    def _on_step_selected(self, etape_id: str, etape_desc: str):
        """Callback : une étape a été sélectionnée → recharger les inputs."""
        self._active_etape_id = etape_id
        self.loaded_files.clear()

        # Mettre à jour la description
        self.step_desc.setText(etape_desc)
        self.step_desc.setVisible(True)

        # Mettre à jour l'étape active dans le module
        if self.current_module:
            self.current_module.etape_active = etape_id

        # Recharger les inputs pour cette étape
        from modules.circularisation.module import INPUTS_PAR_ETAPE
        required_inputs = INPUTS_PAR_ETAPE.get(etape_id, [])

        self._clear_inputs()
        if required_inputs:
            self._build_inputs(required_inputs)
            self.inputs_frame.setVisible(True)
        else:
            self.inputs_frame.setVisible(False)

        # Masquer résultats précédents
        self.result_frame.setVisible(False)
        self.btn_open_output.setVisible(False)
        self.preview_table.setVisible(False)
        self.preview_label.setVisible(False)

    def _update_output_dir_display(self):
        """Met à jour l'affichage du répertoire de sortie."""
        # Répertoire par défaut (sera déterminé automatiquement)
        default_dir = "(Déterminé automatiquement - Documents ou Bureau)"
        self.output_line.setText(default_dir)
        self.output_line.setPlaceholderText(default_dir)

    def _build_inputs(self, inputs: list[ModuleInput]):
        """Génère les widgets d'input selon le schéma du module."""
        self._clear_inputs()

        row = 0  # Compteur de lignes pour le placement correct

        for inp in inputs:
            # Label
            label = QLabel(f"{inp.label} {'*' if inp.required else ''}")
            label.setStyleSheet("font-weight: bold;")
            self.inputs_layout.addWidget(label, row, 0)
            row += 1

            # Input widget
            if inp.input_type == "file":
                row_layout = QHBoxLayout()
                line = QLineEdit()
                line.setPlaceholderText("Cliquer ou glisser un fichier...")
                line.setReadOnly(True)
                line.setStyleSheet("color: #000000;")
                btn = QPushButton("Parcourir")
                btn.setObjectName("SecondaryButton")
                exts = " ".join(f"*{e}" for e in inp.extensions)
                btn.clicked.connect(
                    lambda checked, l=line, k=inp.key, x=exts, m=inp.multiple:
                        self._browse_file(l, k, x, m)
                )
                row_layout.addWidget(line, 1)
                row_layout.addWidget(btn)
                container = QWidget()
                container.setLayout(row_layout)
                self.inputs_layout.addWidget(container, row, 0, 1, 2)  # Span 2 columns
                self.input_widgets[inp.key] = line
                row += 1

            elif inp.input_type == "folder":
                row_layout = QHBoxLayout()
                line = QLineEdit()
                line.setPlaceholderText("Sélectionner un dossier...")
                line.setReadOnly(True)
                line.setStyleSheet("color: #000000;")
                btn = QPushButton("Parcourir")
                btn.setObjectName("SecondaryButton")
                btn.clicked.connect(
                    lambda checked, l=line, k=inp.key:
                        self._browse_folder(l, k)
                )
                row_layout.addWidget(line, 1)
                row_layout.addWidget(btn)
                container = QWidget()
                container.setLayout(row_layout)
                self.inputs_layout.addWidget(container, row, 0, 1, 2)  # Span 2 columns
                self.input_widgets[inp.key] = line
                row += 1

            elif inp.input_type == "text":
                line = QLineEdit()
                line.setPlaceholderText(inp.tooltip or "")
                line.setStyleSheet("color: #000000;")
                if inp.default:
                    line.setText(str(inp.default))
                self.inputs_layout.addWidget(line, row, 1)
                self.input_widgets[inp.key] = line
                row += 1

            elif inp.input_type == "number":
                spin = QSpinBox()
                spin.setRange(0, 999999)
                if inp.default is not None:
                    spin.setValue(int(inp.default))
                self.inputs_layout.addWidget(spin, row, 1)
                self.input_widgets[inp.key] = spin
                row += 1

            elif inp.input_type == "combo":
                combo = QComboBox()
                combo.addItems(inp.options or [])
                if inp.default and inp.default in (inp.options or []):
                    combo.setCurrentText(str(inp.default))
                self.inputs_layout.addWidget(combo, row, 1)
                self.input_widgets[inp.key] = combo
                row += 1

            # Tooltip (sauf pour text où c'est dans le placeholder)
            if inp.tooltip and inp.input_type != "text":
                tip = QLabel(inp.tooltip)
                tip.setStyleSheet("color: #6C757D; font-size: 11px; margin-left: 8px;")
                self.inputs_layout.addWidget(tip, row, 0, 1, 2)  # Span 2 columns
                row += 1

    def _build_params(self, params: list[dict]):
        """Génère les widgets de paramètres ajustables."""
        # Clear existing
        while self.params_layout.count():
            item = self.params_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        title = QLabel("Paramètres")
        title.setStyleSheet("font-weight: bold; font-size: 14px;")
        self.params_layout.addWidget(title, 0, 0, 1, 2)

        for i, p in enumerate(params, 1):
            lbl = QLabel(p.get("label", p["key"]))
            self.params_layout.addWidget(lbl, i, 0)

            if p.get("type") == "number":
                default_val = p.get("default", 0)
                step = p.get("step", 1)
                # Use QDoubleSpinBox when decimals are needed
                if isinstance(default_val, float) or isinstance(step, float):
                    w = QDoubleSpinBox()
                    w.setRange(float(p.get("min", 0)), float(p.get("max", 999999)))
                    w.setValue(float(default_val))
                    w.setSingleStep(float(step))
                    decimals = max(len(str(step).rstrip("0").split(".")[-1]) if "." in str(step) else 0, 2)
                    w.setDecimals(decimals)
                else:
                    w = QSpinBox()
                    w.setRange(int(p.get("min", 0)), int(p.get("max", 999999)))
                    w.setValue(int(default_val))
                self.params_layout.addWidget(w, i, 1)
                self.input_widgets[f"param_{p['key']}"] = w
            elif p.get("type") == "combo":
                w = QComboBox()
                w.addItems(p.get("options", []))
                self.params_layout.addWidget(w, i, 1)
                self.input_widgets[f"param_{p['key']}"] = w
            else:
                w = QLineEdit()
                w.setText(str(p.get("default", "")))
                w.setStyleSheet("color: #000000;")
                self.params_layout.addWidget(w, i, 1)
                self.input_widgets[f"param_{p['key']}"] = w

    def _clear_inputs(self):
        self.input_widgets.clear()
        while self.inputs_layout.count():
            item = self.inputs_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    def _browse_file(self, line_edit: QLineEdit, key: str, extensions: str, multiple: bool = False):
        if multiple:
            paths, _ = QFileDialog.getOpenFileNames(
                self, "Sélectionner un ou plusieurs fichiers", "",
                f"Fichiers ({extensions});;Tous (*.*)"
            )
            if paths:
                display = "; ".join(paths)
                line_edit.setText(display)
                self.loaded_files[key] = paths
                self.file_loaded.emit(paths[0])
        else:
            path, _ = QFileDialog.getOpenFileName(
                self, "Sélectionner un fichier", "", f"Fichiers ({extensions});;Tous (*.*)"
            )
            if path:
                line_edit.setText(path)
                self.loaded_files[key] = path
                self.file_loaded.emit(path)

    def _browse_folder(self, line_edit: QLineEdit, key: str):
        path = QFileDialog.getExistingDirectory(
            self, "Sélectionner un dossier", ""
        )
        if path:
            line_edit.setText(path)
            self.loaded_files[key] = path

    def _choose_output_dir(self):
        """Permet à l'utilisateur de choisir le répertoire de sortie."""
        dir_path = QFileDialog.getExistingDirectory(
            self, "Choisir le répertoire de sortie",
            self.output_line.text() or str(Path.home())
        )
        if dir_path:
            self.output_line.setText(dir_path)

    def _collect_inputs(self):
        """Récupère les valeurs des champs d'input."""
        values = {}

        # Champs d'input normaux
        for key, widget in self.input_widgets.items():
            if isinstance(widget, QLineEdit):
                values[key] = widget.text().strip()
            elif isinstance(widget, QComboBox):
                values[key] = widget.currentText().strip()
            elif isinstance(widget, QDoubleSpinBox):
                values[key] = widget.value()
            elif isinstance(widget, QSpinBox):
                values[key] = widget.value()
            elif isinstance(widget, QDateEdit):
                values[key] = widget.date().toString("yyyy-MM-dd")

        # Fichiers chargés (priorité sur les QLineEdit)
        for key, file_val in self.loaded_files.items():
            values[key] = file_val

        # Injecter l'étape active (pour les modules multi-étapes)
        if self._active_etape_id:
            values["_etape"] = self._active_etape_id

        # Répertoire de sortie (si spécifié manuellement)
        output_dir = self.output_line.text().strip()
        if output_dir and not output_dir.startswith("(Déterminé automatiquement"):
            values["_output_dir"] = output_dir

        return values

    def prefill_inputs(self, saved_params: dict):
        """Pré-remplit les widgets d'input avec les paramètres sauvegardés."""
        for key, value in saved_params.items():
            if key in self.input_widgets:
                widget = self.input_widgets[key]
                if isinstance(widget, QLineEdit):
                    widget.setText(str(value))
                elif isinstance(widget, QComboBox):
                    if value in [widget.itemText(i) for i in range(widget.count())]:
                        widget.setCurrentText(str(value))
                elif isinstance(widget, QSpinBox):
                    try:
                        widget.setValue(int(value))
                    except (ValueError, TypeError):
                        pass
                elif isinstance(widget, QDateEdit):
                    # Assuming date format yyyy-MM-dd
                    from PyQt6.QtCore import QDate
                    date = QDate.fromString(str(value), "yyyy-MM-dd")
                    if date.isValid():
                        widget.setDate(date)

    def _on_preview(self):
        if not self.current_module:
            return

        inputs = self._collect_inputs()

        # Validation
        ok, errors = self.current_module.validate(inputs)
        if not ok:
            QMessageBox.warning(self, "Validation",
                                "Erreurs :\n" + "\n".join(f"• {e}" for e in errors))
            return

        # Aperçu
        try:
            df = self.current_module.preview(inputs)
            if df is not None and not df.empty:
                self.preview_table.load_dataframe(df)
                self.preview_table.setVisible(True)
                self.preview_label.setVisible(True)
            else:
                QMessageBox.information(self, "Aperçu",
                                        "Pas d'aperçu disponible pour ce module.")
        except Exception as e:
            QMessageBox.critical(self, "Erreur aperçu", str(e))

    def _on_execute(self):
        if not self.current_module:
            return

        inputs = self._collect_inputs()

        ok, errors = self.current_module.validate(inputs)
        if not ok:
            QMessageBox.warning(self, "Validation",
                                "Erreurs :\n" + "\n".join(f"• {e}" for e in errors))
            return

        self.execute_requested.emit(inputs)

    def show_progress(self, percent: int, message: str = ""):
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(percent)
        if message:
            self.progress_message.setText(message)
            self.progress_message.setVisible(True)
        self.btn_execute.setEnabled(False)

    def show_result(self, result):
        """Affiche le résultat après exécution."""
        self.progress_bar.setVisible(False)
        self.progress_message.setVisible(False)
        self.btn_execute.setEnabled(True)

        if result.success:
            self.result_label.setStyleSheet("color: #2E7D32;")
            stats_text = ""
            if result.stats:
                stats_lines = [f"  • {k}: {v}" for k, v in result.stats.items()]
                stats_text = "\n" + "\n".join(stats_lines)
            self.result_label.setText(
                f"OK  Traitement réussi\n\n"
                f"{result.message}{stats_text}"
            )
            self._output_path = result.output_path

            # Libellé dynamique du bouton selon le type de résultat
            if result.output_path:
                p = Path(result.output_path)
                if p.is_dir():
                    self.btn_open_output.setText("Ouvrir le dossier")
                elif p.suffix.lower() in (".pdf", ".eml"):
                    self.btn_open_output.setText("Ouvrir le dossier")
                else:
                    self.btn_open_output.setText("Ouvrir le fichier")
                self.btn_open_output.setVisible(True)
            else:
                self.btn_open_output.setVisible(False)
        else:
            self.result_label.setStyleSheet("color: #DC3545;")
            error_text = "\n".join(result.errors) if result.errors else result.message
            self.result_label.setText(f"ERREUR\n\n{error_text}")

        if result.warnings:
            warns = "\n".join(f"Attention : {w}" for w in result.warnings)
            current = self.result_label.text()
            self.result_label.setText(f"{current}\n\n{warns}")

        self.result_frame.setVisible(True)

    def _on_open_output(self):
        """Ouvre le fichier OU le dossier contenant le résultat."""
        if not self._output_path or not os.path.exists(self._output_path):
            return

        import subprocess, sys as _sys
        p    = Path(self._output_path)
        plat = _sys.platform

        # Dossier → ouvrir directement dans l'Explorateur
        if p.is_dir():
            target = str(p)
        # PDF / EML / JSON → ouvrir le dossier parent et sélectionner le fichier
        elif p.suffix.lower() in (".pdf", ".eml", ".json"):
            target = str(p.parent)
        else:
            # Excel, Word → ouvrir le fichier
            target = str(p)

        if plat == "win32":
            if p.is_dir() or p.suffix.lower() in (".pdf", ".eml", ".json"):
                subprocess.Popen(["explorer", target])
            else:
                os.startfile(target)
        elif plat == "darwin":
            subprocess.Popen(["open", target])
        else:
            subprocess.Popen(["xdg-open", target])
