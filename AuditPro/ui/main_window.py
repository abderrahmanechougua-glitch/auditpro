"""
Fenêtre principale d'AuditPro.
Orchestre : sidebar, workspace, assistant, profils.
"""
import os
from pathlib import Path
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QMenuBar, QMenu,
    QStatusBar, QMessageBox, QFileDialog
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QAction, QIcon

from core.module_registry import ModuleRegistry
from core.file_detector import FileDetector
from core.profiles import ProfileManager
from core.history import HistoryManager
from core.worker import Worker
from core.config import APP_NAME, APP_VERSION

from ui.module_panel import ModulePanel
from ui.workspace import Workspace
from ui.assistant_panel import AssistantPanel
from ui.profile_dialog import ProfileDialog
from ui.styles import STYLESHEET


class MainWindow(QMainWindow):
    """Fenêtre principale AuditPro."""

    def __init__(self):
        super().__init__()

        # ── Services ──────────────────────────────────
        self.registry = ModuleRegistry()
        self.detector = FileDetector()
        self.profiles = ProfileManager()
        self.history = HistoryManager()

        self.current_module = None
        self.current_profile = ""
        self.worker = None

        # ── Configuration fenêtre ─────────────────────
        self.setWindowTitle(f"{APP_NAME} — {APP_VERSION}")
        self.setMinimumSize(1100, 700)
        self.resize(1280, 780)
        self.setStyleSheet(STYLESHEET)

        icon_path = Path(__file__).parent.parent / "resources" / "icons" / "logo.png"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        # ── Menu bar ──────────────────────────────────
        self._build_menu()

        # ── Layout principal ──────────────────────────
        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # ── Panneau gauche (modules) ──────────────────
        self.module_panel = ModulePanel()
        self.module_panel.module_selected.connect(self._on_module_selected)
        layout.addWidget(self.module_panel)

        # ── Zone centrale (workspace) ─────────────────
        self.workspace = Workspace()
        self.workspace.file_loaded.connect(self._on_file_loaded)
        self.workspace.execute_requested.connect(self._on_execute)
        layout.addWidget(self.workspace, 1)

        # ── Panneau droit (assistant) ─────────────────
        self.assistant = AssistantPanel(self.history)
        layout.addWidget(self.assistant)

        # ── Status bar ────────────────────────────────
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self._update_status()

        # ── Peupler la sidebar ────────────────────────
        modules = self.registry.get_all()
        self.module_panel.set_modules(modules)

        if not modules:
            self.assistant.show_message(
                "Aucun module détecté.\n\n"
                "Ajoutez vos scripts dans le dossier modules/\n"
                "avec un fichier module.py dans chaque sous-dossier."
            )
        elif self.registry.load_warnings:
            warnings_text = "\n".join(f"• {w}" for w in self.registry.load_warnings)
            self.assistant.show_message(
                f"⚠ {len(self.registry.load_warnings)} module(s) non chargé(s) :\n\n"
                f"{warnings_text}"
            )

    def _build_menu(self):
        menubar = self.menuBar()

        # Fichier
        file_menu = menubar.addMenu("Fichier")

        open_action = QAction("Ouvrir un fichier...", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self._open_file)
        file_menu.addAction(open_action)

        file_menu.addSeparator()

        quit_action = QAction("Quitter", self)
        quit_action.setShortcut("Ctrl+Q")
        quit_action.triggered.connect(self.close)
        file_menu.addAction(quit_action)

        # Profils
        profile_menu = menubar.addMenu("Profils")

        manage_profiles = QAction("Gérer les profils...", self)
        manage_profiles.triggered.connect(self._show_profiles)
        profile_menu.addAction(manage_profiles)

        # Aide
        help_menu = menubar.addMenu("Aide")

        about_action = QAction("À propos", self)
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)

        help_menu.addSeparator()

        contact_action = QAction("Contacter l'auteur", self)
        contact_action.triggered.connect(self._contact_author)
        help_menu.addAction(contact_action)

    def _on_module_selected(self, module_name: str):
        """Un module a été cliqué dans la sidebar."""
        module = self.registry.get(module_name)
        if module:
            self.current_module = module
            self.workspace.set_module(module)
            self.assistant.update_for_module(module)
            self._update_status()

            # Charger les paramètres du profil si actif
            if self.current_profile:
                saved = self.profiles.get_last_params(
                    self.current_profile, module_name
                )
                self.workspace.prefill_inputs(saved)

    def _on_file_loaded(self, file_path: str):
        """Un fichier a été chargé (drop ou browse)."""
        # Détection automatique
        modules = self.registry.get_all()
        detections = self.detector.detect(file_path, modules)
        self.assistant.show_detection(detections)

        # Si un module est détecté avec haute confiance et aucun module actif
        if detections and detections[0]["score"] >= 0.7 and not self.current_module:
            best = detections[0].get("module")
            if best is not None:
                self._on_module_selected(best.name)
                self.module_panel.select_module(best.name)

        file_info = self.detector.get_file_info(file_path)
        self.status.showMessage(
            f"Fichier : {file_info['name']} — "
            f"{file_info['size_kb']} Ko — "
            f"{file_info['total_rows']} lignes — "
            f"Onglets : {', '.join(file_info.get('sheets', []))}"
        )

    def _on_execute(self, inputs: dict):
        """Lance l'exécution du module actif."""
        if not self.current_module:
            return

        # Déterminer le dossier de sortie
        output_dir = self._determine_output_dir(inputs)

        # Vérifier les permissions du répertoire de sortie
        permission_ok, error_msg = self._check_output_permissions(output_dir)
        if not permission_ok:
            QMessageBox.critical(self, "Erreur de permissions",
                                f"Impossible d'écrire dans le dossier de sortie :\n{output_dir}\n\n"
                                f"Erreur : {error_msg}\n\n"
                                f"Solutions :\n"
                                f"• Vérifiez que le disque n'est pas plein\n"
                                f"• Redémarrez l'application\n"
                                f"• Contactez le support si le problème persiste")
            return

        # Afficher un avertissement si le répertoire a été changé
        if error_msg:
            QMessageBox.information(self, "Répertoire de sortie modifié",
                                  f"Le répertoire de sortie a été automatiquement changé :\n\n{error_msg}\n\n"
                                  f"Cliquez sur OK pour continuer.")

        self.workspace.show_progress(0, "Démarrage du traitement...")

        # Exécuter dans un thread séparé
        module = self.current_module

        def run_module(progress_callback=None):
            return module.execute(inputs, output_dir,
                                  progress_callback=progress_callback)

        self.worker = Worker(run_module)
        self.worker.progress.connect(self.workspace.show_progress)
        self.worker.finished.connect(self._on_execution_finished)
        self.worker.error.connect(self._on_execution_error)
        self.worker.start()

    def _determine_output_dir(self, inputs: dict) -> str:
        """Tous les outputs vont dans Téléchargements/AuditPro_Output."""
        # Dossier Téléchargements Windows
        downloads_dir = Path.home() / "Downloads"
        if downloads_dir.exists():
            return str(downloads_dir / "AuditPro_Output")

        # Fallback si Downloads introuvable
        return str(Path.home() / "AuditPro_Output")

    def _check_output_permissions(self, output_dir: str) -> tuple[bool, str]:
        """Vérifie que le répertoire de sortie est accessible en écriture."""
        try:
            # Créer le répertoire
            os.makedirs(output_dir, exist_ok=True)

            # Tester l'écriture d'un fichier temporaire
            test_file = Path(output_dir) / ".auditpro_test_write"
            with open(test_file, 'w') as f:
                f.write("test")

            # Supprimer le fichier de test
            test_file.unlink()

            return True, ""

        except Exception as e:
            # Si ça échoue, essayer un répertoire de fallback
            fallback_dir = str(Path.home() / "AuditPro_Output_Fallback")
            try:
                os.makedirs(fallback_dir, exist_ok=True)
                test_file = Path(fallback_dir) / ".auditpro_test_write"
                with open(test_file, 'w') as f:
                    f.write("test")
                test_file.unlink()
                return True, f"Répertoire de sortie changé vers : {fallback_dir}"
            except Exception as e2:
                return False, f"Impossible d'accéder à aucun répertoire de sortie. Erreur : {str(e)}"

    def _on_execution_finished(self, result):
        """Traitement terminé."""
        self.workspace.show_result(result)

        # Historique
        input_name = ""
        if self.current_module:
            input_name = self.current_module.name

        self.assistant.add_to_history(
            module_name=input_name,
            input_file="",
            output_file=getattr(result, 'output_path', ''),
            success=getattr(result, 'success', False),
            stats=getattr(result, 'stats', {})
        )

        # Sauvegarder les params dans le profil (apprentissage)
        if self.current_profile and self.current_module:
            inputs = self.workspace._collect_inputs()
            self.profiles.save_last_params(
                self.current_profile, self.current_module.name, inputs
            )

        self._update_status()

    def _on_execution_error(self, error_msg: str):
        from modules.base_module import ModuleResult
        result = ModuleResult(success=False, errors=[error_msg])
        self.workspace.show_result(result)

    def _open_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Ouvrir un fichier",
            "", "Excel (*.xlsx *.xls *.xlsm *.csv);;Tous (*.*)"
        )
        if path:
            self._on_file_loaded(path)

    def _show_profiles(self):
        dialog = ProfileDialog(self.profiles, self)
        dialog.profile_selected.connect(self._on_profile_selected)
        dialog.exec()

    def _on_profile_selected(self, name: str):
        self.current_profile = name
        self._update_status()

    def _show_about(self):
        QMessageBox.about(
            self, "À propos",
            f"<h2>{APP_NAME}</h2>"
            f"<p>Version {APP_VERSION}</p>"
            f"<hr>"
            f"<p style='font-style: italic;'>by Abderrahmane Chougua</p>"
            f"<p><a href='mailto:abderrahmanechougua@gmail.com' style='color: #6B21A8;'>abderrahmanechougua@gmail.com</a></p>"
        )

    def _contact_author(self):
        """Ouvre le client mail par défaut."""
        import webbrowser
        webbrowser.open("mailto:abderrahmanechougua@gmail.com")

    def _update_status(self):
        mod_name = self.current_module.name if self.current_module else "—"
        profile = self.current_profile or "Aucun profil"
        n_modules = self.registry.count()
        self.status.showMessage(
            f"Module : {mod_name}  |  Profil : {profile}  |  "
            f"{n_modules} module(s) disponible(s)"
        )
