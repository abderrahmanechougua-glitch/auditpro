"""
Gestion des profils clients.
Sauvegarde les paramètres préférés par client (module, seuils, chemins…).
"""
import json
from datetime import datetime
from .config import PROFILES_FILE


class ProfileManager:
    """CRUD sur les profils clients stockés en JSON."""

    def __init__(self):
        self._ensure_file()

    def _ensure_file(self):
        if not PROFILES_FILE.exists():
            PROFILES_FILE.write_text("{}", encoding="utf-8")

    def _load(self) -> dict:
        try:
            return json.loads(PROFILES_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, FileNotFoundError):
            return {}

    def _save(self, data: dict):
        PROFILES_FILE.write_text(
            json.dumps(data, indent=2, ensure_ascii=False, default=str),
            encoding="utf-8"
        )

    def list_profiles(self) -> list[str]:
        """Liste les noms de profils existants."""
        return sorted(self._load().keys())

    def get(self, name: str) -> dict:
        """Retourne un profil par nom. {} si inexistant."""
        return self._load().get(name, {})

    def save(self, name: str, settings: dict):
        """Crée ou met à jour un profil."""
        data = self._load()
        existing = data.get(name, {})
        existing.update(settings)
        existing["_updated_at"] = datetime.now().isoformat()
        if "_created_at" not in existing:
            existing["_created_at"] = existing["_updated_at"]
        data[name] = existing
        self._save(data)

    def delete(self, name: str) -> bool:
        data = self._load()
        if name in data:
            del data[name]
            self._save(data)
            return True
        return False

    def get_last_params(self, profile_name: str, module_name: str) -> dict:
        """Retourne les derniers paramètres utilisés pour un module donné."""
        profile = self.get(profile_name)
        return profile.get("modules", {}).get(module_name, {})

    def save_last_params(self, profile_name: str, module_name: str, params: dict):
        """Enregistre les paramètres utilisés (apprentissage)."""
        profile = self.get(profile_name)
        profile.setdefault("modules", {})[module_name] = params
        self.save(profile_name, profile)
