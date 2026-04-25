"""
Commentaire analytique standard — Revue Analytique BG-GL.

Le commentaire conserve l'unité native de la BG (aucune conversion forcée).
"""
from __future__ import annotations

from .models import AnalyticalComment


class CommentWriter:
    """Construit le bloc de commentaire analytique standard."""

    def build_comment(
        self,
        label: str,
        solde_n: float,
        solde_n1: float,
        exercice: str = "2025",
    ) -> AnalyticalComment:
        """
        Construit un AnalyticalComment avec texte simple inspiré du SRM.

        label    : intitulé du compte agrégé (ex: "Charges d'exploitation")
        solde_n  : solde N (total réel)
        solde_n1 : solde N-1 (total réel)
        exercice : année de l'exercice
        """
        try:
            year_n = int(exercice)
            year_n1 = year_n - 1
        except (ValueError, TypeError):
            year_n = "N"
            year_n1 = "N-1"

        variation = solde_n - solde_n1
        variation_direction = "à la hausse" if variation > 0 else "à la baisse" if variation < 0 else "stable"

        solde_n_fmt = f"{int(solde_n):,}".replace(",", " ")
        solde_n1_fmt = f"{int(solde_n1):,}".replace(",", " ")
        variation_abs_fmt = f"{int(abs(variation)):,}".replace(",", " ")

        pct = 0
        if abs(solde_n1) > 0:
            pct = round((abs(variation) / abs(solde_n1)) * 100, 1)

        text = (
            f"Le compte {label} présente un solde de KMAD {solde_n_fmt} au 31/12/{year_n} "
            f"contre un solde de KMAD {solde_n1_fmt} au 31/12/{year_n1}, "
            f"soit une variation {variation_direction} de KMAD {variation_abs_fmt} soit {pct}% de variation."
        )

        return AnalyticalComment(
            account_code=label,
            text_rendered=text,
        )
