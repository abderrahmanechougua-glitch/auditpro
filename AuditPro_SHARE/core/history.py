"""
Historique des traitements récents.
Conserve les N dernières exécutions avec métadonnées.
"""
import json
from datetime import datetime
from .config import HISTORY_FILE, MAX_HISTORY


class HistoryManager:

    def __init__(self):
        self._ensure_file()

    def _ensure_file(self):
        if not HISTORY_FILE.exists():
            HISTORY_FILE.write_text("[]", encoding="utf-8")

    def _load(self) -> list:
        try:
            return json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, FileNotFoundError):
            return []

    def _save(self, data: list):
        HISTORY_FILE.write_text(
            json.dumps(data, indent=2, ensure_ascii=False, default=str),
            encoding="utf-8"
        )

    def add(self, module_name: str, input_file: str, output_file: str,
            profile: str = "", stats: dict = None, success: bool = True):
        """Ajoute une entrée à l'historique."""
        entries = self._load()
        entries.insert(0, {
            "timestamp": datetime.now().isoformat(),
            "module": module_name,
            "input": input_file,
            "output": output_file,
            "profile": profile,
            "stats": stats or {},
            "success": success,
        })
        # Garder uniquement les N derniers
        entries = entries[:MAX_HISTORY]
        self._save(entries)

    def get_recent(self, n: int = 10) -> list[dict]:
        return self._load()[:n]

    def clear(self):
        self._save([])
