"""
Configuration centralisée pour le retraitement comptable intelligent.
Définit les seuils, règles de mapping et profils ERP.
"""
import json
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Dict, List, Optional


GL_CONFIG = {
    "tolerance": 0.01,
    "tolerance_balance": 0.01,
    "title_keywords": [
        "opening balance",
        "closing balance",
        "solde initial",
        "solde final",
    ],
    "total_keywords": [
        "total",
        "sous-total",
        "sous total",
        "report",
        "cumul",
    ],
    # Mots-clés à exclure des lignes (lignes structurelles non-comptables).
    # Opening/Closing balance sont CONSERVÉS (TYPE=OPENING/CLOSING BALANCE dans la sortie).
    "exclude_keywords": [
        r"total",
        r"sous[- ]total",
        r"solde initial",
        r"solde final",
    ],
    "date_candidates": ["date", "voucher", "date operation", "date ecriture", "dt"],
    "voucher_candidates": ["voucher", "n° document", "n°document", "numero document", "document", "piece", "n° piece"],
    "account_candidates": ["n°compte", "compte", "account", "account_number", "numero compte"],
    "description_candidates": ["description", "libellé", "libelle", "motif", "narration", "intitulé", "intitule"],
    "debit_candidates": ["débit", "debit", "debit_amount", "montant debit"],
    "credit_candidates": ["crédit", "credit", "credit_amount", "montant credit"],
    "accumulation_candidates": ["accumulation", "accumulated", "accmula", "accumula", "solde cumulé", "solde cumule", "running balance", "cumulative balance"],
    "solde_candidates": ["solde", "balance", "net"],
    "description_clean_regex": r"^[\s;!()]+|[\s;!()]+$",
    "spaces_regex": r"\s+",
}


@dataclass
class Config:
    """Configuration du retraitement comptable."""

    # Tolérance pour les validations numériques (ex: écart d/c)
    tolerance: float = 0.01

    # Seuil de correspondance floue pour le mapping des colonnes (difflib)
    column_match_threshold: float = 0.80

    # Format de détection des montants : auto, european (1.000,00), american (1,000.00)
    amount_format_detection: str = "auto"

    # Nombre de lignes à échantillonner pour détecter le format des montants
    amount_format_sample_size: int = 100

    # Nombre de lignes à échantillonner pour détecter le format des dates
    date_format_sample_size: int = 50

    # Activer la suppression des lignes marquées 'total' (si False, les garder)
    remove_total_rows: bool = False

    # Poids utilisé dans la détection du type de document
    # scores = {col_weight * présence_col + keyword_weight * présence_keyword}
    doc_type_col_weight: int = 10
    doc_type_keyword_weight: int = 2

    # Mots-clés bannissables (lignes contenant ces mots seront marquées)
    total_keywords: List[str] = field(
        default_factory=lambda: [
            "total",
            "sous-total",
            "sous total",
            "report",
            "report à nouveau",
        ]
    )

    # Mappings de colonnes par type de document
    column_mappings: Dict[str, Dict[str, List[str]]] = field(default_factory=dict)

    # Profils ERP : configuration spécifique par ERP
    erp_profiles: Dict[str, Dict] = field(default_factory=dict)

    def __post_init__(self):
        """Initialiser les mappings et profils par défaut."""
        if not self.column_mappings:
            self.column_mappings = {
                "GL": {
                    "n°compte": [
                        "compte",
                        "n° compte",
                        "num_compte",
                        "numero compte",
                        "account",
                        "account_number",
                    ],
                    "date": ["date", "dt", "date écriture", "date ecriture"],
                    "journal": ["journal", "journal_code", "code journal"],
                    "pièce": [
                        "pièce",
                        "piece",
                        "n° pièce",
                        "num piece",
                        "reference",
                        "ref",
                    ],
                    "libellé": ["libellé", "lib", "description", "motif", "narration"],
                    "débit": ["débit", "debit", "déb", "deb", "debit_amount"],
                    "crédit": ["crédit", "credit", "créd", "cred", "credit_amount"],
                    "intitulé": [
                        "intitulé",
                        "intitule",
                        "account_name",
                        "account_description",
                    ],
                },
                "BG": {
                    "n°compte": [
                        "compte",
                        "n° compte",
                        "num_compte",
                        "numero compte",
                        "account",
                    ],
                    "intitulé": ["intitule", "intitulé", "libellé", "description"],
                    "solde_débit": ["solde débit", "sd", "débit", "debit", "debit_balance"],
                    "solde_crédit": [
                        "solde crédit",
                        "sc",
                        "crédit",
                        "credit",
                        "credit_balance",
                    ],
                    "date": ["date", "dt"],
                    "libellé": ["libellé", "lib", "commentaire"],
                },
                "AUX": {
                    "tiers": [
                        "tiers",
                        "fournisseur",
                        "client",
                        "nom",
                        "raison sociale",
                        "supplier",
                        "customer",
                    ],
                    "sf_débit": [
                        "solde final débit",
                        "sf débit",
                        "solde débiteur",
                        "final_debit",
                    ],
                    "sf_crédit": [
                        "solde final crédit",
                        "sf crédit",
                        "solde créditeur",
                        "final_credit",
                    ],
                    "si_débit": [
                        "solde initial débit",
                        "si débit",
                        "opening_debit",
                    ],
                    "si_crédit": [
                        "solde initial crédit",
                        "si crédit",
                        "opening_credit",
                    ],
                    "mvt_débit": ["mouvement débit", "mvt débit", "movement_debit"],
                    "mvt_crédit": ["mouvement crédit", "mvt crédit", "movement_credit"],
                },
            }

        if not self.erp_profiles:
            self.erp_profiles = {
                "default": {},
                "sap": {"amount_format_detection": "auto"},
                "ciel": {"tolerance": 0.005},
                "sage": {"tolerance": 0.01},
            }

    @classmethod
    def load_from_json(cls, json_path: Path) -> "Config":
        """Charge la configuration depuis un fichier JSON."""
        with open(json_path) as f:
            data = json.load(f)
        return cls(**data)

    def save_to_json(self, json_path: Path) -> None:
        """Sauvegarde la configuration dans un fichier JSON."""
        json_path.parent.mkdir(parents=True, exist_ok=True)
        with open(json_path, "w") as f:
            json.dump(asdict(self), f, indent=2, ensure_ascii=False, default=str)

    def apply_erp_profile(self, erp_name: str) -> None:
        """Applique un profil ERP aux paramètres."""
        if erp_name not in self.erp_profiles:
            raise ValueError(
                f"Profil ERP '{erp_name}' non trouvé. "
                f"Disponibles: {list(self.erp_profiles.keys())}"
            )
        profile = self.erp_profiles[erp_name]
        for key, value in profile.items():
            if hasattr(self, key):
                setattr(self, key, value)
