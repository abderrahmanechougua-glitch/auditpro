"""
Registre des modules — scanne automatiquement le dossier modules/
et enregistre tout module qui hérite de BaseModule.
"""
import importlib
import pkgutil
from pathlib import Path
from .config import MODULES_DIR


class ModuleRegistry:
    """Auto-découvre et enregistre tous les modules disponibles."""

    def __init__(self):
        self._modules: dict[str, object] = {}
        self.load_warnings: list[str] = []
        self._discover()

    def _discover(self):
        """Scanne modules/ pour trouver les sous-dossiers avec un module.py."""
        modules_path = MODULES_DIR

        for item in modules_path.iterdir():
            if not item.is_dir() or item.name.startswith("_"):
                continue

            module_file = item / "module.py"
            if not module_file.exists():
                continue

            try:
                # Import dynamique : modules.tva.module
                mod = importlib.import_module(f"modules.{item.name}.module")

                # Cherche la classe qui hérite de BaseModule
                from modules.base_module import BaseModule
                for attr_name in dir(mod):
                    attr = getattr(mod, attr_name)
                    if (isinstance(attr, type)
                            and issubclass(attr, BaseModule)
                            and attr is not BaseModule):
                        instance = attr()
                        self._modules[instance.name] = instance
                        break

            except Exception as e:
                warning = f"Module '{item.name}' non chargé : {type(e).__name__}: {e}"
                self.load_warnings.append(warning)
                print(f"  [!] {warning}")

    def get_all(self) -> dict[str, object]:
        """Retourne tous les modules découverts. {nom: instance}"""
        return dict(self._modules)

    def get(self, name: str):
        """Retourne un module par son nom."""
        return self._modules.get(name)

    def get_by_category(self) -> dict[str, list]:
        """Retourne les modules groupés par catégorie."""
        categories = {}
        for mod in self._modules.values():
            cat = getattr(mod, 'category', 'Général')
            categories.setdefault(cat, []).append(mod)
        return categories

    def names(self) -> list[str]:
        return list(self._modules.keys())

    def count(self) -> int:
        return len(self._modules)
