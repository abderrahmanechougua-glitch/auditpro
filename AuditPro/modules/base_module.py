"""
Contrat abstrait pour tous les modules AuditPro.
Chaque module (TVA, CNSS, Lettrage, etc.) hérite de BaseModule
et implémente les 5 méthodes obligatoires.
"""
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Any
import pandas as pd


@dataclass
class ModuleInput:
    """Décrit un input attendu par le module."""
    key: str                    # Identifiant interne (ex: "declarations")
    label: str                  # Label affiché (ex: "Déclarations TVA")
    input_type: str = "file"    # "file", "folder", "text", "number", "date", "combo"
    extensions: list = field(default_factory=lambda: [".xlsx", ".xls"])
    required: bool = True
    multiple: bool = False      # Accepte plusieurs fichiers
    default: Any = None
    options: list = field(default_factory=list)  # Pour type "combo"
    tooltip: str = ""


@dataclass
class ModuleResult:
    """Résultat standard retourné par execute()."""
    success: bool
    output_path: str = ""
    message: str = ""
    stats: dict = field(default_factory=dict)    # Métriques (total, nb lignes, écarts...)
    warnings: list = field(default_factory=list)  # Alertes non-bloquantes
    errors: list = field(default_factory=list)


class BaseModule(ABC):
    """
    Classe abstraite — contrat pour chaque module.
    
    Pour intégrer un script existant :
    1. Créer modules/mon_module/module.py
    2. Hériter de BaseModule
    3. Implémenter les 5 méthodes
    4. Le registre le détecte automatiquement
    """

    # ── Métadonnées (à surcharger) ───────────────────────
    name: str = "Module sans nom"
    icon: str = ""                      # Nom de fichier icône (ex: "tva.png")
    description: str = "Aucune description."
    category: str = "Général"           # Pour grouper dans le menu
    help_text: str = ""                 # Texte long pour le panneau assistant
    version: str = "1.0"

    # ── Règles de détection automatique ──────────────────
    # Mots-clés dans les noms de colonnes qui déclenchent ce module
    detection_keywords: list = []
    # Score minimum pour considérer une détection valide (0.0 à 1.0)
    detection_threshold: float = 0.5

    @abstractmethod
    def get_required_inputs(self) -> list[ModuleInput]:
        """
        Déclare les inputs nécessaires.
        L'UI génère automatiquement les champs de saisie correspondants.
        """
        ...

    @abstractmethod
    def validate(self, inputs: dict) -> tuple[bool, list[str]]:
        """
        Vérifie que les inputs sont valides avant exécution.
        Retourne (True, []) si OK, ou (False, ["erreur 1", "erreur 2"]).
        """
        ...

    @abstractmethod
    def preview(self, inputs: dict) -> pd.DataFrame | None:
        """
        Retourne un aperçu (5-10 lignes) du résultat attendu.
        Retourne None si pas de preview possible.
        """
        ...

    @abstractmethod
    def execute(self, inputs: dict, output_dir: str,
                progress_callback=None) -> ModuleResult:
        """
        Exécute le traitement principal.
        
        progress_callback(percent, message) : pour mettre à jour la barre.
        Retourne un ModuleResult.
        """
        ...

    def get_default_params(self) -> dict:
        """
        Paramètres par défaut du module (peuvent être surchargés par un profil).
        """
        return {}

    def get_param_schema(self) -> list[dict]:
        """
        Schéma des paramètres ajustables par l'utilisateur.
        Chaque dict : {"key": ..., "label": ..., "type": ..., "default": ...}
        """
        return []
