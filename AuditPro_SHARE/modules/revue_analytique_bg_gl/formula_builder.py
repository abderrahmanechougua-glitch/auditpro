"""
Construction des formules Excel — Revue Analytique BG-GL.

Les formules utilisent INDEX/MATCH pour résoudre les comptes de la BG.
Compatible Excel 2013+ (XLOOKUP n'existe que dans Excel 365).
Le séparateur d'arguments (FORMULA_SEP) est abstrait pour supporter
les locales FR (;) et EN (,).
"""
from __future__ import annotations

import re

from openpyxl.utils import get_column_letter

from .models import AnalyticalRow, AnalyticalTotalRow

# Séparateur d'arguments de formules Excel (locale FR = ";", EN = ",").
FORMULA_SEP = ";"

# Ordre des colonnes dans les feuilles analytiques (1-based index)
COL_COMPTE = 1
COL_INTITULE = 2
COL_REF = 3
COL_SOLDE_N = 4
COL_SOLDE_N1 = 5
COL_VARIATION = 6
COL_VARIATION_PCT = 7

# Colonnes correspondantes dans la feuille BG (fallback)
BG_COL_COMPTE_DEFAULT = 1
BG_COL_INTITULE_DEFAULT = 2
BG_COL_REF_DEFAULT = None
BG_COL_SOLDE_N_DEFAULT = 4
BG_COL_SOLDE_N1_DEFAULT = 5


def _col_letter(idx: int) -> str:
    return get_column_letter(idx)


class FormulaBuilder:
    """
    Construit les formules INDEX/MATCH pour les lignes et totaux analytiques.
    Compatible Excel 2013+ (remplace XLOOKUP qui ne fonctionne qu'en Excel 365).
    """

    def __init__(
        self,
        bg_col_map: dict[str, int],
        bg_last_row: int,
        has_ref_col: bool = False,
        bg_sheet_name: str = "BG",
    ):
        self.bg_col_map = bg_col_map
        self.bg_last_row = bg_last_row
        self.has_ref_col = has_ref_col
        self.bg_sheet_name = bg_sheet_name

        self._bg_compte_col = bg_col_map.get("Compte", BG_COL_COMPTE_DEFAULT)
        self._bg_intitule_col = bg_col_map.get("Intitule", BG_COL_INTITULE_DEFAULT)
        self._bg_ref_col = bg_col_map.get("Ref") if has_ref_col else None
        self._bg_solde_n_col = bg_col_map.get("Solde N", BG_COL_SOLDE_N_DEFAULT)
        self._bg_solde_n1_col = bg_col_map.get("Solde N-1", BG_COL_SOLDE_N1_DEFAULT)

    def _sheet_ref(self) -> str:
        """Nom de feuille prêt pour formule Excel (quote si nécessaire)."""
        name = str(self.bg_sheet_name or "BG")
        if re.match(r"^[A-Za-z0-9_]+$", name):
            return name
        escaped = name.replace("'", "''")
        return f"'{escaped}'"

    def _vlookup_compatible(
        self,
        lookup_value_formula: str,
        bg_target_col: int,
        if_not_found: str = '""',
    ) -> str:
        """
        Construit une formule INDEX/MATCH compatible Excel 2013+.
        Compatible avec toutes les versions Excel, contrairement à XLOOKUP (Excel 365 only).
        Utilise des plages délimitées ($F$3:$I$742) au lieu de colonnes entières.
        """
        lookup_col = _col_letter(self._bg_compte_col)
        target_col = _col_letter(bg_target_col)
        sheet = self._sheet_ref()
        header_row = 2
        first_data_row = header_row + 1
        last_data_row = self.bg_last_row
        return (
            f"=IFERROR(INDEX({sheet}!${target_col}${first_data_row}:${target_col}${last_data_row},"
            f"MATCH({lookup_value_formula},{sheet}!${lookup_col}${first_data_row}:${lookup_col}${last_data_row},0)),"
            f"{if_not_found})"
        )

    def _lookup_key(self, row: int) -> str:
        """Expression Excel pour la clé de lookup : A{row} (sans conversion, texte direct)."""
        compte_cell = f"{_col_letter(COL_COMPTE)}{row}"
        return compte_cell

    def build_subaccount_row(
        self,
        code_8: str,
        sheet_row: int,
    ) -> AnalyticalRow:
        """Construit une AnalyticalRow avec toutes les formules XLOOKUP pointant vers BG."""
        lookup_key = self._lookup_key(sheet_row)

        intitule_formula = self._vlookup_compatible(lookup_key, self._bg_intitule_col)
        solde_n_formula = self._vlookup_compatible(lookup_key, self._bg_solde_n_col)
        solde_n1_formula = self._vlookup_compatible(lookup_key, self._bg_solde_n1_col)

        if self.has_ref_col and self._bg_ref_col:
            ref_formula = self._vlookup_compatible(lookup_key, self._bg_ref_col)
        else:
            ref_formula = ""

        solde_n_col = _col_letter(COL_SOLDE_N)
        solde_n1_col = _col_letter(COL_SOLDE_N1)
        variation_formula = f"={solde_n_col}{sheet_row}-{solde_n1_col}{sheet_row}"

        return AnalyticalRow(
            compte_formula_or_value=code_8,
            intitule_formula=intitule_formula,
            ref_formula_or_empty=ref_formula,
            solde_n_formula=solde_n_formula,
            solde_n1_formula=solde_n1_formula,
            variation_formula=variation_formula,
        )

    def build_total_row(
        self,
        code_4: str,
        first_data_row: int,
        last_data_row: int,
    ) -> AnalyticalTotalRow:
        """Construit la ligne de total avec formules SUM (incluant Variation)."""
        solde_n_col = _col_letter(COL_SOLDE_N)
        solde_n1_col = _col_letter(COL_SOLDE_N1)
        variation_col = _col_letter(COL_VARIATION)

        total_n_formula = f"=SUM({solde_n_col}{first_data_row}:{solde_n_col}{last_data_row})"
        total_n1_formula = f"=SUM({solde_n1_col}{first_data_row}:{solde_n1_col}{last_data_row})"
        total_variation_formula = f"=SUM({variation_col}{first_data_row}:{variation_col}{last_data_row})"

        return AnalyticalTotalRow(
            label=code_4,
            total_n_formula=total_n_formula,
            total_n1_formula=total_n1_formula,
            variation_formula=total_variation_formula,
            style_bold=True,
        )