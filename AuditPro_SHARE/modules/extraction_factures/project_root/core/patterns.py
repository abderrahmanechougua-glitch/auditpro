from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Pattern


@dataclass(frozen=True)
class PatternRule:
    """Compiled regex rule with a base confidence weight."""

    pattern: Pattern[str]
    confidence: int
    method: str


RE_FLAGS = re.IGNORECASE | re.MULTILINE
DOTALL_FLAGS = re.IGNORECASE | re.MULTILINE | re.DOTALL


def _rule(regex: str, confidence: int, method: str, dotall: bool = False) -> PatternRule:
    flags = DOTALL_FLAGS if dotall else RE_FLAGS
    return PatternRule(re.compile(regex, flags), confidence, method)


AMOUNT_RE = r"([0-9OIlSbBZz][0-9OIlSbBZz\s\.,]{0,40}[0-9OIlSbBZz])"

INVOICE_NUMBER_RULES: tuple[PatternRule, ...] = (
    _rule(r"NUM[ÉE]RO\s*:\s*([A-Z0-9][A-Z0-9\-_\/\s\.]{1,25})", 95, "invoice_number:num"),
    _rule(r"N[°º]\s*FACTURE\s*:?\s*([A-Z0-9][A-Z0-9\-_\/\s\.]{1,25})", 95, "invoice_number:n_facture"),
    _rule(r"FACTURE\s+N[°º]\s*:?\s*([A-Z0-9][A-Z0-9\-_\/\s\.]{1,25})", 92, "invoice_number:facture_n"),
    _rule(r"Facture\s+n[°º]\s*:?\s*([A-Z0-9][A-Z0-9\-_\/\s\.]{1,25})", 90, "invoice_number:facture_n_lower"),
    _rule(r"N[°º]\s*:?\s*([A-Z]{0,3}\d{2,}[A-Z0-9\-_\/]*)", 80, "invoice_number:fallback_n"),
)

DATE_RULES: tuple[PatternRule, ...] = (
    _rule(r"DATE\s*:\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})", 95, "date:label"),
    _rule(r"Date\s+d['’]?\s*[ée]mission\s*:\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})", 95, "date:emission"),
    _rule(r"(?:Rabat|Casablanca|Marrakech|Tanger|Fès|Agadir|Tétouan|Kénitra)[,\s]+le\s+(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})", 90, "date:city"),
    _rule(r"\ble\s+(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})", 75, "date:le"),
    _rule(r"(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{4})", 60, "date:fallback"),
)

HT_RULES: tuple[PatternRule, ...] = (
    _rule(rf"Total\s+HT\s*[:\s]*{AMOUNT_RE}", 95, "ht:total_ht"),
    _rule(rf"Montant\s+(?:total\s+)?H\.?T\.?\s*[:\s]*{AMOUNT_RE}", 95, "ht:montant_ht"),
    _rule(rf"Sous[- ]?total\s*[:\s]*(?:MAD\s*)?{AMOUNT_RE}", 90, "ht:sous_total"),
    _rule(rf"H\.?T\.?\s*[:\s]+{AMOUNT_RE}", 80, "ht:short"),
)

TVA_AMOUNT_RULES: tuple[PatternRule, ...] = (
    _rule(rf"T\.?V\.?A\.?\s*(?:\(\s*\d+\s*%?\s*\))?\s*[:\s]*(?:MAD\s*)?{AMOUNT_RE}", 90, "tva:amount"),
    _rule(rf"TVA\s+\d+\s*%?\s*[:\s]*{AMOUNT_RE}", 90, "tva:rate_amount"),
)

TVA_RATE_RULES: tuple[PatternRule, ...] = (
    _rule(r"TVA\s*\(?\s*(\d{1,2}(?:[.,]\d{1,2})?)\s*%\s*\)?", 95, "tva_rate:label"),
    _rule(r"Taux\s+T\.?V\.?A\.?\s*[:\s]*(\d{1,2}(?:[.,]\d{1,2})?)", 90, "tva_rate:taux"),
)

TTC_RULES: tuple[PatternRule, ...] = (
    _rule(r"Arr[êeé]t[ée]e?\s+(?:la\s+pr[ée]sente\s+facture\s+)?[àa]\s+la\s+(?:somme|valeur)\s+de\s*:?\s*(.+?)(?:\.|$)", 98, "ttc:words"),
    _rule(rf"Total\s+G[ée]n[ée]ral\s*[:\s]*{AMOUNT_RE}", 95, "ttc:total_general"),
    _rule(rf"Total\s+T\.?T\.?C\.?\s*[:\s]*{AMOUNT_RE}", 95, "ttc:total_ttc"),
    _rule(rf"Montant\s+(?:total\s+)?T\.?T\.?C\.?\s*[:\s]*{AMOUNT_RE}", 95, "ttc:montant_ttc"),
    _rule(rf"Net\s+[àa]\s+payer\s*[:\s]*(?:MAD\s*)?{AMOUNT_RE}", 90, "ttc:net_payer"),
)

SUPPLIER_LABEL_RULES: tuple[PatternRule, ...] = (
    _rule(r"(?:Soci[ée]t[ée]|Sté|Entreprise|Raison\s+sociale)\s*:?\s*([A-ZÀ-ÿ][A-Za-zÀ-ÿ\s\-&\.]{3,60})", 90, "supplier:label"),
)

JURIDICAL_RULES: dict[str, tuple[PatternRule, ...]] = {
    "ice": (
        _rule(r"ICE\s*:?\s*(\d{15})", 95, "ice"),
        _rule(r"I\.?C\.?E\.?\s*:?\s*(\d{15})", 90, "ice_alt"),
    ),
    "if_fiscal": (
        _rule(r"I\.?F\.?\s*:?\s*(\d{6,10})", 90, "if"),
        _rule(r"Identifiant\s+fiscal\s*:?\s*(\d{6,10})", 90, "if_label"),
    ),
    "patente": (
        _rule(r"(?:PATENTE|TP|Patente|T\.?P\.?)\s*:?\s*(\d{6,15})", 90, "patente"),
    ),
    "rc": (
        _rule(r"R\.?C\.?\s*:?\s*(\d{3,10})", 90, "rc"),
        _rule(r"(?:Registre\s+de\s+Commerce|N[°o]\s+de\s+RC)\s*:?\s*(\d{3,10})", 90, "rc_label"),
    ),
    "cnss": (
        _rule(r"CNSS\s*:?\s*(\d{6,15})", 90, "cnss"),
    ),
}

TABLE_BLOCK_RULE = _rule(
    rf"(?:Total|Montant|Sous[- ]?total)\s+H\.?T\.?\s*[:\s]*{AMOUNT_RE}"
    rf".*?T\.?V\.?A\.?\s*(?:\(?\s*\d+\s*%?\s*\)?)?\s*[:\s]*{AMOUNT_RE}"
    rf".*?(?:Total|Montant)\s+(?:G[ée]n[ée]ral|T\.?T\.?C\.?)\s*[:\s]*{AMOUNT_RE}",
    97,
    "table:block",
    dotall=True,
)

INVOICE_BOUNDARY_RULES: tuple[PatternRule, ...] = (
    _rule(r"\b(FACTURE|Invoice)\b", 0, "boundary:invoice"),
    _rule(r"(?:N[°º]\s*(?:FACTURE|Client)|NUM[ÉE]RO|Facture\s+N)", 0, "boundary:number"),
    _rule(r"(?:Total\s+H\.?T|Montant\s+H\.?T)", 0, "boundary:total_ht"),
    _rule(r"(?:T\.?T\.?C|Net\s+[àa]\s+payer|Total\s+G)", 0, "boundary:ttc"),
)
