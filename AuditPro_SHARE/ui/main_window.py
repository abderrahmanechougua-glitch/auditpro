"""
Fenêtre principale d'AuditPro.
Orchestre : sidebar, workspace, assistant, profils.
"""
import os
import traceback
from pathlib import Path
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QSplitter,
    QFrame, QStatusBar, QMessageBox,
    QFileDialog, QTabWidget
)
from PyQt6.QtCore import Qt, QTimer
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
from ui.ai_chat_panel import AIChatPanel
from ui.notifications import NotificationManager, Toast
from ui.profile_dialog import ProfileDialog
from ui.styles import get_stylesheet


class MainWindow(QMainWindow):
    """Fenêtre principale AuditPro."""

    def __init__(self):
        super().__init__()

        # ── Services ──────────────────────────────────
        self.registry = ModuleRegistry()
        self.detector = FileDetector()
        self.profiles = ProfileManager()
        self.history = HistoryManager()

        # ── Agent IA ──────────────────────────────────
        from agent.llama_agent import LlamaAgent
        self._agent = LlamaAgent(self.registry)

        self.current_module = None
        self.current_profile = ""
        self.worker = None
        self.current_file_path = ""
        self.current_file_info = {}
        self.current_detections = []
        self.current_output_dir = ""
        self.last_result = None
        self.last_run_success = None
        self.theme_name = "light"
        self.progress_toast: Toast | None = None

        # ── Configuration fenêtre ─────────────────────
        self.setWindowTitle(f"{APP_NAME} — {APP_VERSION}")
        self.setMinimumSize(1100, 700)
        self.resize(1280, 780)
        self.setStyleSheet(get_stylesheet(self.theme_name))

        icon_path = Path(__file__).parent.parent / "resources" / "icons" / "logo.png"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        # ── Menu bar ──────────────────────────────────
        self._build_menu()
        self._build_shell()
        self._build_status_bar()
        self._wire_workspace_signals()
        self.notifications = NotificationManager(self)

        modules = self._populate_modules()
        self._initialize_empty_state(modules)
        self._update_status()

    def _build_shell(self):
        """Construit la coque principale avec barre contexte + splitter."""
        central = QWidget()
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        self.content_splitter = QSplitter(Qt.Orientation.Horizontal)
        self.content_splitter.setObjectName("MainContentSplitter")

        # ── Panneau gauche (modules) ──────────────────
        self.module_panel = ModulePanel()
        self.module_panel.module_selected.connect(self._on_module_selected)
        self.content_splitter.addWidget(self.module_panel)

        # ── Zone centrale (workspace) ─────────────────
        self.workspace = Workspace()
        self.content_splitter.addWidget(self.workspace)

        # ── Panneau droit (assistant + agent IA) ─────
        self.right_rail = self._build_right_rail()
        self.content_splitter.addWidget(self.right_rail)

        self.content_splitter.setStretchFactor(0, 0)
        self.content_splitter.setStretchFactor(1, 1)
        self.content_splitter.setStretchFactor(2, 0)
        self.content_splitter.setCollapsible(0, False)
        self.content_splitter.setCollapsible(2, True)
        self.content_splitter.setSizes([250, 860, 320])

        root.addWidget(self.content_splitter, 1)

    def _build_right_rail(self) -> QFrame:
        """Construit le rail droit redimensionnable."""
        rail = QFrame()
        rail.setObjectName("RightRail")

        rail_layout = QVBoxLayout(rail)
        rail_layout.setContentsMargins(0, 0, 0, 0)
        rail_layout.setSpacing(0)

        self.right_tabs = QTabWidget()
        self.right_tabs.setObjectName("RightRailTabs")
        self.right_tabs.setMinimumWidth(280)

        self.assistant = AssistantPanel(self.history)
        self.ai_chat = AIChatPanel(self._agent)
        self.right_tabs.addTab(self.assistant, "Assistant")
        self.right_tabs.addTab(self.ai_chat, "Agent IA")
        rail_layout.addWidget(self.right_tabs)

        return rail

    def _build_status_bar(self):
        """Crée la barre de statut dédiée aux messages transitoires."""
        self.status = QStatusBar()
        self.setStatusBar(self.status)

    def _wire_workspace_signals(self):
        self.workspace.file_loaded.connect(self._on_file_loaded)
        self.workspace.file_dropped.connect(self._on_file_dropped)
        self.workspace.execute_requested.connect(self._on_execute)

    def _populate_modules(self):
        modules = self.registry.get_all()
        self.module_panel.set_modules(modules)
        return modules

    def _initialize_empty_state(self, modules):
        if not modules:
            self.assistant.show_message(
                "Aucun module détecté.\n\n"
                "Ajoutez vos scripts dans le dossier modules/\n"
                "avec un fichier module.py dans chaque sous-dossier."
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

        # Affichage
        view_menu = menubar.addMenu("Affichage")

        toggle_theme_action = QAction("Basculer le thème", self)
        toggle_theme_action.setShortcut("Ctrl+T")
        toggle_theme_action.triggered.connect(self._toggle_theme)
        view_menu.addAction(toggle_theme_action)

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
        try:
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
        except Exception as e:
            self._show_transient_status(f"Erreur chargement module '{module_name}': {e}")
            QMessageBox.critical(
                self,
                "Erreur module",
                "Le module n'a pas pu être chargé.\n\n"
                f"Module : {module_name}\n"
                f"Erreur : {e}\n\n"
                "L'application reste ouverte pour corriger le problème."
            )
            print(traceback.format_exc())

    def _on_file_loaded(self, file_path: str):
        """Un fichier a été chargé (drop ou browse)."""
        # Détection automatique
        modules = self.registry.get_all()
        detections = self.detector.detect(file_path, modules)
        self.current_file_path = file_path
        self.current_detections = detections
        self.assistant.show_detection(detections)

        # Si un module est détecté avec haute confiance et aucun module actif
        if detections and detections[0]["score"] >= 0.7 and not self.current_module:
            best = detections[0].get("module")
            if best is not None:
                self._on_module_selected(best.name)
                self.module_panel.select_module(best.name)

        file_info = self.detector.get_file_info(file_path)
        self.current_file_info = file_info
        self._update_status()
        self._show_transient_status(
            f"Fichier chargé : {file_info['name']}"
        )

    def _on_file_dropped(self, file_path: str):
        self.notifications.show_info(
            "Fichier reçu",
            f"{Path(file_path).name} a été injecté dans le workflow actif.",
            duration=2600,
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

        # Conserver le dossier de sortie pour le bouton "Ouvrir le résultat".
        self.workspace.set_output_dir_hint(output_dir)
        self.current_output_dir = output_dir
        self.last_run_success = None
        self._update_status()

        self.workspace.show_progress(0, "Démarrage du traitement...")
        self._show_progress_toast("Traitement en cours", "Initialisation du module...", 0)

        # Exécuter dans un thread séparé
        module = self.current_module

        def run_module(progress_callback=None):
            return module.execute(inputs, output_dir,
                                  progress_callback=progress_callback)

        self.worker = Worker(run_module)
        self.worker.progress.connect(self.workspace.show_progress)
        self.worker.progress.connect(self._on_worker_progress)
        self.worker.finished.connect(self._on_execution_finished)
        self.worker.error.connect(self._on_execution_error)
        self.worker.start()

    def _determine_output_dir(self, inputs: dict) -> str:
        """Détermine le dossier de sortie selon la logique métier."""
        # Priorité 1 : répertoire explicitement choisi par l'utilisateur
        user_dir = inputs.get("_output_dir", "").strip()
        if user_dir and not user_dir.startswith("(Déterminé automatiquement"):
            return user_dir

        # Priorité 2 : Documents/AuditPro_Output
        documents_dir = Path.home() / "Documents"
        if documents_dir.exists():
            return str(documents_dir / "AuditPro_Output")

        # Priorité 3 : Bureau/AuditPro_Output
        desktop_dir = Path.home() / "Desktop"
        if desktop_dir.exists():
            return str(desktop_dir / "AuditPro_Output")

        # Fallback : home
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
        self.last_result = result
        self.last_run_success = bool(getattr(result, "success", False))

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
        self._show_transient_status("Traitement terminé", 4000)
        self._dismiss_progress_toast()
        self.notifications.show_success(
            "Traitement terminé",
            getattr(result, "message", "Le traitement s'est terminé avec succès."),
            duration=3500,
        )

    def _on_execution_error(self, error_msg: str):
        from modules.base_module import ModuleResult
        result = ModuleResult(success=False, errors=[error_msg])
        self.workspace.show_result(result)
        self.last_result = result
        self.last_run_success = False
        self._update_status()
        self._show_transient_status(f"Erreur exécution : {error_msg}", 7000)
        self._dismiss_progress_toast()
        self.notifications.show_error("Erreur d'exécution", error_msg, duration=5500)

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
        profile = self.current_profile or "Aucun"
        n_modules = self.registry.count()
        details = ""
        if self.current_file_info:
            name = self.current_file_info.get("name", "—")
            rows = self.current_file_info.get("total_rows", "—")
            details = f"  |  Fichier : {name} ({rows} lignes)"
        self.status.showMessage(
            f"Module : {mod_name}  |  Profil : {profile}  |  "
            f"{n_modules} module(s) disponible(s){details}"
        )

    def _show_transient_status(self, message: str, timeout_ms: int = 5000):
        self.status.showMessage(message, timeout_ms)

    def _show_progress_toast(self, title: str, message: str, progress: int):
        self._dismiss_progress_toast()
        self.progress_toast = self.notifications.show_progress(title, message)
        self.progress_toast.set_progress(progress)

    def _on_worker_progress(self, percent: int, message: str = ""):
        if self.progress_toast is None:
            self._show_progress_toast("Traitement en cours", message or "Exécution...", percent)
            return
        if message:
            self.progress_toast.update_message(message)
        self.progress_toast.set_progress(percent)
        self.notifications.reposition()

        # Safety net: if the module signals 100 % complete, schedule the toast
        # to dismiss after a short delay so it disappears even when
        # _on_execution_finished is delayed or the normal dismiss path fails.
        if percent >= 100:
            QTimer.singleShot(600, self._dismiss_progress_toast)

    def _dismiss_progress_toast(self):
        if self.progress_toast is not None:
            self.progress_toast.dismiss()
            self.progress_toast = None

    def _toggle_theme(self):
        self.theme_name = "dim" if self.theme_name == "light" else "light"
        self.setStyleSheet(get_stylesheet(self.theme_name))
        theme_label = "Clair" if self.theme_name == "light" else "Dim"
        self._show_transient_status(f"Thème activé : {theme_label}", 2500)
        self.notifications.show_info("Thème changé", f"Mode {theme_label} activé.", duration=2200)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "notifications"):
            self.notifications.reposition()
