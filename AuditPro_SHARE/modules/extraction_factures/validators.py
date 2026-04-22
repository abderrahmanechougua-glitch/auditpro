#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
validators.py — TVA validation logic for Moroccan invoices.

Single source of truth for TVA coherence checks.
factextv19.py delegates all TVA control to validate_tva() here.
"""

from dataclasses import dataclass
from typing import Optional


# ============================================================
# TVA CONTROL RESULT
# ============================================================

@dataclass
class TVAControl:
    ht: Optional[float] = None
    tva_calculee: Optional[float] = None
    tva_document: Optional[float] = None
    ttc: Optional[float] = None
    taux: Optional[float] = None
    taux_detecte: Optional[float] = None
    statut: str = ""
    message: str = ""


# ============================================================
# VALIDATE TVA
# Accepts pre-normalized float values (normalization done by caller).
# ============================================================

def validate_tva(
    ht_val: Optional[float],
    tva_val: Optional[float],
    ttc_val: Optional[float],
    taux_doc: Optional[float] = None,
    ttc_words_val: Optional[float] = None,
) -> TVAControl:
    """
    Validate TVA coherence for a Moroccan invoice.

    Args:
        ht_val:       Montant HT as float (or None if not found).
        tva_val:      Montant TVA as float (or None if not found).
        ttc_val:      Montant TTC as float (or None if not found).
        taux_doc:     Taux TVA read from the document (0–100) or None.
        ttc_words_val: TTC derived from the French words amount (cross-check).

    Returns:
        TVAControl with statut, taux, tva_calculee and diagnostic message.
    """
    ctrl = TVAControl(
        ht=ht_val,
        tva_document=tva_val,
        ttc=ttc_val,
        taux_detecte=taux_doc,
    )

    if ht_val and ttc_val and ht_val > 0:
        ctrl.tva_calculee = round(ttc_val - ht_val, 2)
        ctrl.taux = round((ctrl.tva_calculee / ht_val) * 100, 2)

        if tva_val is not None:
            ecart_tva = abs(ctrl.tva_calculee - tva_val)
            somme_check = abs((ht_val + tva_val) - ttc_val)
            if ecart_tva <= 1.0 and somme_check <= 1.0:
                ctrl.statut = "OK"
                ctrl.message = (
                    f"HT({ht_val:.2f}) + TVA({tva_val:.2f}) = TTC({ttc_val:.2f})"
                )
            else:
                ctrl.statut = "ANOMALIE"
                ctrl.message = (
                    f"Écart TVA: {ecart_tva:.2f} | Écart somme: {somme_check:.2f}"
                )
        else:
            ctrl.statut = "TVA_MANQUANTE"
            ctrl.message = (
                f"TVA calculée: {ctrl.tva_calculee:.2f} "
                f"(taux ~{ctrl.taux:.1f}%)"
            )

        # Verify declared rate vs computed rate
        if taux_doc and ctrl.taux:
            ecart_taux = abs(ctrl.taux - taux_doc)
            if ecart_taux > 1.0:
                ctrl.statut = "TAUX_ANOMALIE"
                ctrl.message += (
                    f" | Taux doc: {taux_doc}%, Taux calculé: {ctrl.taux:.1f}%"
                )

    elif ht_val and tva_val and not ttc_val:
        ctrl.ttc = round(ht_val + tva_val, 2)
        ctrl.tva_calculee = tva_val
        if ht_val > 0:
            ctrl.taux = round((tva_val / ht_val) * 100, 2)
        ctrl.statut = "TTC_CALCULE"
        ctrl.message = f"TTC calculé: {ctrl.ttc:.2f}"

    else:
        ctrl.statut = "INCOMPLET"
        missing = []
        if not ht_val:
            missing.append("HT")
        if not ttc_val:
            missing.append("TTC")
        if not tva_val:
            missing.append("TVA")
        ctrl.message = f"Manquant: {', '.join(missing)}"

    # Cross-check with French words amount
    if ttc_words_val and ttc_val:
        ecart_lettres = abs(ttc_words_val - ttc_val)
        if ecart_lettres > 1.0:
            ctrl.message += (
                f" | ALERTE lettres: {ttc_words_val:.2f} ≠ TTC {ttc_val:.2f}"
            )

    return ctrl
