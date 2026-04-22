#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
patterns.py — Centralized extraction rules for Moroccan invoices.

All PatternRule objects are the single source of truth.
factextv19.py imports from here instead of defining its own PATTERNS_* lists.
"""

from dataclasses import dataclass


# ============================================================
# AMOUNT REGEX — reused in pattern strings via f-strings
# ============================================================

AMOUNT_RE = r"([0-9OIlSbB][0-9OIlSbB\s\.,']{0,40}[0-9OIlSbB])"


# ============================================================
# PATTERN RULE
# ============================================================

@dataclass
class PatternRule:
    """A single extraction rule: a regex pattern and its base confidence score."""
    pattern: str
    confidence: int


# ============================================================
# CONTEXT KEYWORD LISTS
# Used by get_context_score() in factextv19.py
# ============================================================

HT_KEYWORDS  = ["ht", "h.t", "hors taxe", "sous-total", "sous total", "subtotal", "net"]
TVA_KEYWORDS = ["tva", "t.v.a", "taxe", "vat"]
TTC_KEYWORDS = ["ttc", "t.t.c", "total", "net à payer", "total général", "total general",
                "net a payer", "arrêté", "arrete"]


# ============================================================
# RULES — ordered by specificity (highest confidence first)
# ============================================================

RULES_NUM_FACTURE = (
    PatternRule(r"N[°º]\s*facture\s+Date\s+Page[^\n]*\n\s*([A-Z0-9][A-Z0-9\-_\/\. ]{3,25})\s+\d{1,2}[\/'’\-\.]\d{1,2}[\/'’\-\.]\d{2,4}", 95),
    PatternRule(r"NUM[ÉE]RO\s*:\s*([A-Z0-9][A-Z0-9\-_\/\. ]{1,25}?)(?=\s+(?:ICE|I\.?F\.?|RC|TP|PATENTE|CNSS|DATE|CLIENT)\b|$)", 95),
    PatternRule(r"[ËÉE]?\s*NUM[ÉE]RO\s*:\s*([A-Z]{1,3}\d{5,12})", 94),
    PatternRule(r"N[°º]\s*FACTURE\s*:?\s*([A-Z0-9][A-Z0-9\-_\/\. ]{1,25}?)(?=\s+(?:ICE|I\.?F\.?|RC|TP|PATENTE|CNSS|DATE|CLIENT)\b|$)", 95),
    PatternRule(r"N[°º]\s*CLIENT\s*:?\s*[A-Z0-9\-_\/]+\s+N[°º]\s*FACTURE\s*:?\s*([A-Z0-9][A-Z0-9\-_\/\. ]{1,25})", 95),
    PatternRule(r"FACTURE\s+N[°º]\s*:?\s*([A-Z0-9][A-Z0-9\-_\/\. ]{1,25}?)(?=\s+(?:ICE|I\.?F\.?|RC|TP|PATENTE|CNSS|DATE|CLIENT)\b|$)", 95),
    PatternRule(r"(?:FACTURE\s+DE\s+)?BC\s*N[°º]\s*:?\s*([A-Z0-9][A-Z0-9\-_\/\. ]{1,25})", 88),
    PatternRule(r"Commande\s+N[°º]\s*:?\s*([A-Z0-9][A-Z0-9\-_\/\. ]{1,25})", 85),
    PatternRule(r"Facture\s+N[°º]\s*:?\s*([A-Z0-9][A-Z0-9\-_\/\ \.]{1,25})", 90),
    PatternRule(r"Facture\s+n[°º]\s*:?\s*([A-Z0-9][A-Z0-9\-_\/\ \.]{1,25})", 90),
    PatternRule(r"Invoice\s+N[°º]\s*:?\s*([A-Z0-9][A-Z0-9\-_\/\ \.]{1,25})", 85),
    PatternRule(r"N[°º]\s*:?\s*([A-Z]{0,3}\d{2,}[A-Z0-9\-_\/]*)", 80),
)

RULES_DATE = (
    PatternRule(r"DATE\s*:\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})", 95),
    PatternRule(r"DATE\s*:\s*(\d{1,2}[\/'’\-\.]\d{1,2}[\/'’\-\.]\d{2,4})", 95),
    PatternRule(r"DATE\s*:\s*(\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2})", 95),
    PatternRule(r"DATE\s*:\s*(\d{4}[\/'’\-\.]\d{1,2}[\/'’\-\.]\d{1,2})", 95),
    PatternRule(r"Date\s+d[''']?\s*[ée]mission\s*:\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})", 95),
    PatternRule(r"Du\s*:\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})", 90),
    PatternRule(r"Du\s*:\s*(\d{1,2}[\/'’\-\.]\d{1,2}[\/'’\-\.]\d{2,4})", 90),
    PatternRule(r"Du\s*:\s*(\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2})", 90),
    PatternRule(r"Date\s+d[''']?\s*[ée]ch[ée]ance\s*:\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})", 85),
    PatternRule(r"(?:Rabat|Casablanca|Marrakech|Tanger|Fès|Agadir|Tétouan|Kénitra)[,\s]+le\s+(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})", 90),
    PatternRule(r"R[èe]glement\s+le\s*:\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})", 80),
    PatternRule(r"\ble\s+(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})", 70),
    PatternRule(r"(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{4})", 60),
)

RULES_HT = (
    PatternRule(rf"Total\s+HT\s*[:\s]*{AMOUNT_RE}", 95),
    PatternRule(rf"Total\s*\(\s*H\.?T\.?\s*\)\s*[\|:_\-\s]*{AMOUNT_RE}", 95),
    PatternRule(rf"Subtotal\s*[:\s]*(?:MAD\s*)?{AMOUNT_RE}", 93),
    PatternRule(rf"Total\s+net\s+H\.?T\.?\s*[:\s]*(?:MAD\s*)?{AMOUNT_RE}", 93),
    PatternRule(rf"Montant\s+H\.?T\.?\s*[:\s]*(?:MAD\s*)?{AMOUNT_RE}", 92),
    PatternRule(rf"Montant\s+HT\s*[:\s]*(?:DH|MAD)?\s*{AMOUNT_RE}", 92),
    PatternRule(rf"Montant\s+Total\s+H\.?T\.?\s*[:\s]*(?:DH|MAD)?\s*{AMOUNT_RE}", 92),
    PatternRule(rf"Total\s+HT\s*[\|:_\-\s]+{AMOUNT_RE}", 91),
    PatternRule(rf"Montant\s+(?:total\s+)?H\.?T\.?\s*[:\s]*{AMOUNT_RE}", 95),
    PatternRule(rf"Sous[- ]?total\s*[:\s]*(?:MAD\s*)?{AMOUNT_RE}", 90),
    PatternRule(rf"Montant\s+H\.T\.\s*\n?\s*{AMOUNT_RE}", 85),
    PatternRule(rf"H\.?T\.?\s*[:\s]+{AMOUNT_RE}", 80),
    PatternRule(rf"(?:MAD|DH)\s+{AMOUNT_RE}\s*$", 75),
)

RULES_TVA_MONTANT = (
    PatternRule(rf"T\.?V\.?A\.?\s*(?:\(\s*\d+\s*%?\s*\))?\s*[:\s]*(?:MAD\s*)?{AMOUNT_RE}", 90),
    PatternRule(rf"TVA\s+\d+\s*%?\s*[:\s]*{AMOUNT_RE}", 90),
    PatternRule(rf"Total\s+TVA\s*[:\s]*(?:MAD\s*)?{AMOUNT_RE}", 89),
    PatternRule(rf"T\.V\.A\.\s*[:\s]*{AMOUNT_RE}", 85),
    PatternRule(rf"TVA\s*[:\s]+{AMOUNT_RE}", 80),
)

RULES_TVA_TAUX = (
    PatternRule(r"TVA\s*\(?\s*(\d{1,2}(?:[.,]\d{1,2})?)\s*%\s*\)?", 95),
    PatternRule(r"T\.?V\.?A\.?\s+(\d{1,2}(?:[.,]\d{1,2})?)\s*%", 95),
    PatternRule(r"Taux\s+T\.?V\.?A\.?\s*[:\s]*(\d{1,2}(?:[.,]\d{1,2})?)", 90),
    PatternRule(r"(\d{1,2})\s*%\s*$", 60),
)

RULES_TTC = (
    # Montant en lettres (highest confidence — parsed separately)
    PatternRule(r"Arr[êeé]t[ée]e?\s+(?:la\s+pr[ée]sente\s+facture\s+)?[àa]\s+la\s+(?:somme|valeur)\s+de\s*:?\s*(.+?)(?:\.|$)", 98),
    PatternRule(rf"Total\s+G[ée]n[ée]ral\s*[:\s]*{AMOUNT_RE}", 95),
    PatternRule(rf"Total\s+T\.?T\.?C\.?\s*[:\s]*{AMOUNT_RE}", 95),
    PatternRule(rf"Total\s*\(\s*(?:T\.?T\.?C\.?|F\.?T\.?C\.?)\s*\)\s*[\|:_\-\s]*{AMOUNT_RE}", 95),
    PatternRule(rf"Montant\s+(?:total\s+)?T\.?T\.?C\.?\s*[:\s]*{AMOUNT_RE}", 95),
    PatternRule(rf"Total\s+TTC\s*[:\s]*(?:DH|MAD)?\s*{AMOUNT_RE}", 94),
    PatternRule(rf"NET\s+A\s+PAYER\s*[:\s]*(?:DH|MAD)?\s*{AMOUNT_RE}", 94),
    PatternRule(rf"NET\s+APAYER\s*[:\s]*(?:DH|MAD)?\s*{AMOUNT_RE}", 94),
    PatternRule(rf"Net\s+[àa]\s+payer\s*[:\s]*(?:MAD\s*)?{AMOUNT_RE}", 90),
    PatternRule(rf"{AMOUNT_RE}\s*(?:DH|DIRHAMS?|MAD)\s*$", 75),
    PatternRule(rf"MAD\s+{AMOUNT_RE}\s*$", 75),
)

RULES_FOURNISSEUR = (
    PatternRule(r"(?:Soci[ée]t[ée]|Sté|Entreprise|Raison\s+sociale)\s*:?\s*([A-ZÀ-ÿ][A-Za-zÀ-ÿ\s\-&\.]{3,60})", 90),
    PatternRule(r"^([A-Z][A-Z\s\-&\.]{4,50})\s*$", 70),
)

RULES_ICE = (
    PatternRule(r"ICE\s*:?\s*(\d{15})", 95),
    PatternRule(r"I\.?C\.?E\.?\s*:?\s*(\d{15})", 90),
    PatternRule(r"Identifiant\s+commun\s+(?:de\s+l['''])?entreprise\s*(?:\(ICE\))?\s*:?\s*(\d{15})", 90),
)

RULES_IF = (
    PatternRule(r"I\.?F\.?\s*:?\s*(\d{6,10})", 90),
    PatternRule(r"Identifiant\s+fiscal\s*:?\s*(\d{6,10})", 90),
)

RULES_TP = (
    PatternRule(r"(?:PATENTE|TP|Patente|T\.?P\.?)\s*:?\s*(\d{6,15})", 90),
)

RULES_RC = (
    PatternRule(r"R\.?C\.?\s*:?\s*(\d{3,10})", 90),
    PatternRule(r"(?:Registre\s+de\s+Commerce|N[°o]\s+de\s+RC)\s*:?\s*(\d{3,10})", 90),
)

RULES_CNSS = (
    PatternRule(r"CNSS\s*:?\s*(\d{6,15})", 90),
    PatternRule(r"N[°o]\s+CNSS\s*:?\s*(\d{6,15})", 90),
)

RULES_DESCRIPTION = (
    PatternRule(r"(?:Objet|D[ée]signation|Description)\s*:?\s*([^\n]{10,200})", 85),
    PatternRule(r"Projet\s*:?\s*([^\n]{10,200})", 80),
    PatternRule(r"Mission\s+([^\n]{10,150})", 75),
    PatternRule(r"Prestation\s+(?:de\s+)?([^\n]{10,150})", 70),
)
