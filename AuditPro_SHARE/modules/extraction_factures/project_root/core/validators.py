from __future__ import annotations

from typing import Optional

from ..models.invoice import TVAValidationResult


def validate_tva(
    ht: Optional[float],
    tva_detected: Optional[float],
    ttc: Optional[float],
    vat_rate_detected: Optional[float],
    tolerance_amount: float = 1.0,
    tolerance_rate: float = 1.0,
) -> TVAValidationResult:
    """Validate VAT consistency with configurable tolerances.

    Rules:
    - HT + TVA == TTC within tolerance.
    - Recalculated TVA from TTC-HT compared to detected TVA.
    - Recalculated rate compared to detected rate.

    Args:
        ht: Amount excluding VAT.
        tva_detected: VAT amount extracted from document.
        ttc: Amount including VAT.
        vat_rate_detected: VAT rate extracted from document.
        tolerance_amount: Absolute tolerance for amount checks.
        tolerance_rate: Absolute tolerance for rate checks.

    Returns:
        Structured TVAValidationResult.
    """
    result = TVAValidationResult(
        ht=ht,
        tva_detected=tva_detected,
        ttc=ttc,
        vat_rate_detected=vat_rate_detected,
    )

    if ht is None and ttc is None and tva_detected is None:
        return TVAValidationResult(status="INCOMPLETE", message="HT/TVA/TTC manquants")

    if ht is not None and ttc is not None:
        result.tva_recalculated = round(ttc - ht, 2)
        if ht > 0:
            result.vat_rate_recalculated = round((result.tva_recalculated / ht) * 100, 2)

    if result.tva_recalculated is None and ht is not None and tva_detected is not None:
        result.ttc = round(ht + tva_detected, 2)
        if ht > 0:
            result.vat_rate_recalculated = round((tva_detected / ht) * 100, 2)
        return TVAValidationResult(
            ht=ht,
            tva_detected=tva_detected,
            tva_recalculated=tva_detected,
            ttc=result.ttc,
            vat_rate_detected=vat_rate_detected,
            vat_rate_recalculated=result.vat_rate_recalculated,
            status="TTC_CALCULATED",
            message=f"TTC calculé: {result.ttc:.2f}",
        )

    if result.tva_recalculated is None:
        return TVAValidationResult(
            ht=ht,
            tva_detected=tva_detected,
            ttc=ttc,
            vat_rate_detected=vat_rate_detected,
            status="INCOMPLETE",
            message="Données insuffisantes pour contrôle TVA",
        )

    if tva_detected is None:
        status = "TVA_MISSING"
        message = f"TVA calculée: {result.tva_recalculated:.2f}"
    else:
        amount_gap = abs(result.tva_recalculated - tva_detected)
        sum_gap = abs((ht or 0.0) + tva_detected - (ttc or 0.0))
        if amount_gap <= tolerance_amount and sum_gap <= tolerance_amount:
            status = "OK"
            message = f"HT+TVA=TTC cohérent (écart={amount_gap:.2f})"
        else:
            status = "ANOMALY"
            message = f"Incohérence montants (écart TVA={amount_gap:.2f}, écart somme={sum_gap:.2f})"

    if (
        vat_rate_detected is not None
        and result.vat_rate_recalculated is not None
        and abs(vat_rate_detected - result.vat_rate_recalculated) > tolerance_rate
    ):
        status = "RATE_ANOMALY"
        message += (
            f" | Taux document={vat_rate_detected:.2f}%"
            f", taux recalculé={result.vat_rate_recalculated:.2f}%"
        )

    return TVAValidationResult(
        ht=ht,
        tva_detected=tva_detected,
        tva_recalculated=result.tva_recalculated,
        ttc=ttc,
        vat_rate_detected=vat_rate_detected,
        vat_rate_recalculated=result.vat_rate_recalculated,
        status=status,
        message=message,
    )
