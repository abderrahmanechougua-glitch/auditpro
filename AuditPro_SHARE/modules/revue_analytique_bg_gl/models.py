"""
Dataclasses du domaine — Revue Analytique BG-GL.

Toutes les entités du data-model.md sont représentées ici.
"""
from __future__ import annotations

from dataclasses import dataclass, field


# ── Source ──────────────────────────────────────────────────────────────────


@dataclass
class BGRow:
    """Ligne source depuis la feuille BG."""
    row_index: int
    compte_raw: str | int
    compte_norm: str           # Normalisé (chiffres uniquement, zfill)
    intitule: str
    solde_n: float
    solde_n1: float
    ref: str = ""              # Optionnel — colonne Ref de la BG

    @property
    def is_eligible_4(self) -> bool:
        return len(self.compte_norm) == 4 and self.compte_norm.isdigit()

    @property
    def is_eligible_8(self) -> bool:
        return len(self.compte_norm) == 8 and self.compte_norm.isdigit()


# ── Groupement ──────────────────────────────────────────────────────────────


@dataclass
class SubAccount:
    """Sous-compte 8 chiffres rattaché à un agrégat 4 chiffres."""
    code_8: str
    parent_code_4: str
    source_row_index: int


@dataclass
class AggregateAccount:
    """Compte agrégé 4 chiffres — une feuille analytique par agrégat."""
    code_4: str
    label: str = ""
    sub_accounts: list[SubAccount] = field(default_factory=list)

    @property
    def sheet_name(self) -> str:
        return self.code_4


# ── Feuille analytique ──────────────────────────────────────────────────────


@dataclass
class AnalyticalRow:
    """Ligne sous-compte dans la feuille analytique (toutes valeurs = formules)."""
    compte_formula_or_value: str
    intitule_formula: str
    ref_formula_or_empty: str
    solde_n_formula: str
    solde_n1_formula: str
    variation_formula: str = ""  # Calculée = N - N-1


@dataclass
class AnalyticalTotalRow:
    """Ligne de total en bas du tableau analytique."""
    label: str
    intitule: str = ""
    total_n_formula: str = ""
    total_n1_formula: str = ""
    variation_formula: str = ""  # Calculée = N - N-1
    style_bold: bool = True


@dataclass
class AnalyticalComment:
    """Bloc de commentaire analytique standard sous chaque tableau."""
    account_code: str
    text_rendered: str


@dataclass
class AnalyticalSheet:
    """Feuille analytique générée pour un agrégat."""
    name: str
    columns: tuple = ("Compte", "Intitulé", "Ref", "31/12/N", "31/12/N-1")
    rows_sub_accounts: list[AnalyticalRow] = field(default_factory=list)
    total_row: AnalyticalTotalRow | None = None
    comment_block: AnalyticalComment | None = None


# ── Résumé d'exécution ──────────────────────────────────────────────────────


@dataclass
class GenerationRunSummary:
    """Résumé machine-readable de l'exécution."""
    success: bool
    generated_sheets_count: int = 0
    detected_aggregate_accounts_count: int = 0
    invalid_rows_count: int = 0
    invalid_rows_examples: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
