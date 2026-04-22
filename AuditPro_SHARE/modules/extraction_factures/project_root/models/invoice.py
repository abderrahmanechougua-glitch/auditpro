from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass(frozen=True)
class ExtractionResult:
    """Represents a single extracted field candidate.

    Attributes:
        value: Extracted value as text.
        confidence: Confidence score in [0, 100].
        method: Extraction method label.
    """

    value: str = ""
    confidence: int = 0
    method: str = ""


@dataclass(frozen=True)
class TVAValidationResult:
    """Structured VAT validation output.

    Attributes:
        ht: Parsed HT amount.
        tva_detected: Parsed detected VAT amount.
        tva_recalculated: Recalculated VAT from TTC-HT.
        ttc: Parsed TTC amount.
        vat_rate_detected: VAT rate read on document.
        vat_rate_recalculated: VAT rate recomputed from amounts.
        status: Validation status code.
        message: Human-readable explanation.
    """

    ht: Optional[float] = None
    tva_detected: Optional[float] = None
    tva_recalculated: Optional[float] = None
    ttc: Optional[float] = None
    vat_rate_detected: Optional[float] = None
    vat_rate_recalculated: Optional[float] = None
    status: str = "INCOMPLETE"
    message: str = ""


@dataclass
class Invoice:
    """Invoice aggregate model used across extraction pipeline."""

    source_file: str = ""
    page_start: int = 0
    page_end: int = 0

    invoice_number: ExtractionResult = field(default_factory=ExtractionResult)
    invoice_date: ExtractionResult = field(default_factory=ExtractionResult)
    supplier: ExtractionResult = field(default_factory=ExtractionResult)
    description: ExtractionResult = field(default_factory=ExtractionResult)

    amount_ht: ExtractionResult = field(default_factory=ExtractionResult)
    amount_tva: ExtractionResult = field(default_factory=ExtractionResult)
    amount_ttc: ExtractionResult = field(default_factory=ExtractionResult)
    vat_rate: ExtractionResult = field(default_factory=ExtractionResult)

    ice: ExtractionResult = field(default_factory=ExtractionResult)
    if_fiscal: ExtractionResult = field(default_factory=ExtractionResult)
    patente: ExtractionResult = field(default_factory=ExtractionResult)
    rc: ExtractionResult = field(default_factory=ExtractionResult)
    cnss: ExtractionResult = field(default_factory=ExtractionResult)

    currency: str = "MAD"
    amount_in_words: str = ""
    tva_validation: TVAValidationResult = field(default_factory=TVAValidationResult)
    anomalies: list[str] = field(default_factory=list)
    confidence_score: int = 0
