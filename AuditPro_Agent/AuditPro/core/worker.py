"""
Worker thread pour exécuter les traitements sans bloquer l'UI.
"""
from PyQt6.QtCore import QThread, pyqtSignal


class Worker(QThread):
    """Exécute une fonction dans un thread séparé."""

    progress = pyqtSignal(int, str)        # (pourcentage, message)
    finished = pyqtSignal(object)           # résultat
    error = pyqtSignal(str)                 # message d'erreur

    def __init__(self, func, *args, **kwargs):
        super().__init__()
        self.func = func
        self.args = args
        self.kwargs = kwargs

    def run(self):
        try:
            # Injecter le callback de progression
            self.kwargs["progress_callback"] = self._on_progress
            result = self.func(*self.args, **self.kwargs)
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(f"{type(e).__name__}: {e}")

    def _on_progress(self, percent: int, message: str = ""):
        self.progress.emit(percent, message)
