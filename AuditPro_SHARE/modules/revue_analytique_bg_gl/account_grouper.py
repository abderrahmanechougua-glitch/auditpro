"""
Normalisation et groupement des comptes — Revue Analytique BG-GL.
"""
from __future__ import annotations

import re

from .models import BGRow, AggregateAccount, SubAccount


class AccountNormalizer:
    """
    Normalise une valeur brute de compte en chaîne canonique.

    Règles (Decision 2 + R3):
    - Strip whitespace, cast str, supprimer séparateurs (espaces, tirets, points)
    - Comptes 8 chiffres stockés en entier par Excel (ex: 61200000 → "61200000")
      ou avec zéros de tête perdus (ex: 1200000 → "01200000")
    - Résultat eligble_4 : longueur 4, chiffres uniquement
    - Résultat eligible_8 : longueur 8, chiffres uniquement
    - Sinon : invalid
    """

    _SEPARATOR_RE = re.compile(r"[\s\-./]")

    def normalize(self, raw) -> str:
        """Retourne la chaîne normalisée (peut être vide si None)."""
        if raw is None:
            return ""
        s = str(raw).strip()
        s = self._SEPARATOR_RE.sub("", s)
        if s.endswith(".0") and s[:-2].isdigit():
            s = s[:-2]
        if not s.isdigit():
            return s
        if len(s) < 8:
            return s.zfill(8)
        if len(s) == 8:
            return s
        if len(s) > 8:
            return s[:8]
        return s

    def classify(self, norm: str) -> str:
        """Retourne 'eligible_4', 'eligible_8' ou 'invalid'."""
        if not norm or not norm.isdigit():
            return "invalid"
        if len(norm) == 4:
            return "eligible_4"
        if len(norm) == 8:
            return "eligible_8"
        return "invalid"


class AccountGrouper:
    """
    Construit la liste des AggregateAccount (comptes 4 chiffres) depuis les BGRow.

    La classe est déterminée par le premier chiffre non-nul du compte normalisé.
    Seules les classes sélectionnées sont conservées (par défaut: 1-7).
    """

    def __init__(self, classes: list[str] | None = None):
        """
        Initialise le grouper.
        classes: liste des classes à inclure (ex: ["1", "2", "6", "7"]).
                 Si None, inclut toutes les classes 1-7.
        """
        if classes is None:
            self.classes = {"1", "2", "3", "4", "5", "6", "7"}
        else:
            self.classes = set(c[0] for c in classes) if classes else {"1", "2", "3", "4", "5", "6", "7"}

    def _first_real_digit(self, code: str) -> str:
        """Retourne le premier chiffre non-nul du code."""
        for c in code:
            if c.isdigit() and c != "0":
                return c
        return code[0] if code else ""

    def group(self, rows: list[BGRow]) -> list[AggregateAccount]:
        """Retourne la liste triée des AggregateAccount pour les classes sélectionnées."""
        aggregates: dict[str, AggregateAccount] = {}

        for row in rows:
            code = row.compte_norm
            if len(code) < 4:
                continue
            first_digit = self._first_real_digit(code)
            if first_digit not in self.classes:
                continue

            if row.is_eligible_4:
                code_4 = code
                if code_4 not in aggregates:
                    aggregates[code_4] = AggregateAccount(code_4=code_4, label=row.intitule)
                else:
                    if not aggregates[code_4].label:
                        aggregates[code_4].label = row.intitule

            elif row.is_eligible_8:
                code_4 = code[:4]
                if code_4 not in aggregates:
                    aggregates[code_4] = AggregateAccount(code_4=code_4)
                sub = SubAccount(
                    code_8=code,
                    parent_code_4=code_4,
                    source_row_index=row.row_index,
                )
                aggregates[code_4].sub_accounts.append(sub)

        return sorted(aggregates.values(), key=lambda a: a.code_4)
