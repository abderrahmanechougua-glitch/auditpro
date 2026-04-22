"""
Validator : validation des règles métier comptable.
Génère un rapport de violation avec niveaux de sévérité.
"""
import logging
from typing import Dict, List

import pandas as pd

from .config import Config, GL_CONFIG

logger = logging.getLogger(__name__)


def validate_gl_balance(df: pd.DataFrame) -> List[Dict]:
    """Valide un GL standard (12 colonnes AuditPro) et retourne les violations."""
    violations: List[Dict] = []
    required = ["DATE", "COMPTE", "DESCRIPTION", "DÉBIT", "CRÉDIT", "cumul"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        return [{
            "rule": "GL_missing_standard_columns",
            "severity": "critical",
            "message": f"Colonnes standard manquantes : {missing}",
            "value": None,
            "column": None,
        }]

    tol = float(GL_CONFIG.get("tolerance_balance", GL_CONFIG.get("tolerance", 0.01)))
    debit = pd.to_numeric(df["DÉBIT"], errors="coerce").fillna(0.0)
    credit = pd.to_numeric(df["CRÉDIT"], errors="coerce").fillna(0.0)

    if (debit < 0).any() or (credit < 0).any():
        violations.append({
            "rule": "GL_negative_amounts",
            "severity": "critical",
            "message": "DÉBIT/CRÉDIT contient des valeurs négatives",
            "value": int((debit < 0).sum() + (credit < 0).sum()),
            "column": None,
        })

    # Vérifier l'équilibre uniquement sur les lignes MOUVEMENT (hors Opening/Closing)
    if "TYPE" in df.columns:
        mvt_mask = df["TYPE"] == "MOUVEMENT"
        debit_mvt = debit[mvt_mask]
        credit_mvt = credit[mvt_mask]
    else:
        debit_mvt, credit_mvt = debit, credit

    delta_total = float(debit_mvt.sum() - credit_mvt.sum())
    if abs(delta_total) > tol:
        violations.append({
            "rule": "GL_total_debit_credit_imbalance",
            "severity": "warning",
            "message": f"Total Débit - Crédit (mouvements) = {delta_total:.2f} (tolérance {tol})",
            "value": delta_total,
            "column": None,
        })
    else:
        violations.append({
            "rule": "GL_total_debit_credit_balanced",
            "severity": "info",
            "message": f"Total Débit - Crédit (mouvements) dans la tolérance ({delta_total:.4f})",
            "value": delta_total,
            "column": None,
        })

    logger.info("Validation GL standard: %s violation(s)", len(violations))
    return violations


def validate_document(
    df: pd.DataFrame,
    doc_type: str,
    config: Config,
) -> List[Dict]:
    """
    Valide un document comptable selon les règles métier.
    
    Retourne une liste de violations avec sévérité et message.
    
    Args:
        df: DataFrame normalisé.
        doc_type: Type de document (GL, BG, AUX).
        config: Configuration.
    
    Returns:
        List[Dict] contenant :
        {
            'rule': str (identifiant de la règle),
            'severity': str ('info', 'warning', 'critical'),
            'message': str,
            'value': float ou str (la valeur relevante),
            'column': str (colonne impliquée si applicable)
        }
    """
    violations = []

    # Valider la présence des colonnes requises
    required_cols = {
        "GL": ["n°compte", "date", "libellé"],
        "BG": ["n°compte", "intitulé"],
        "AUX": ["tiers"],
    }

    if doc_type in required_cols:
        missing = [col for col in required_cols[doc_type] if col not in df.columns]
        if missing:
            violations.append({
                "rule": f"{doc_type}_missing_columns",
                "severity": "critical",
                "message": f"Colonnes obligatoires manquantes : {missing}",
                "value": None,
                "column": None,
            })

    # Filtrer les lignes "keep" pour la validation
    if "_flag" in df.columns:
        df_valid = df[df["_flag"] == "keep"].copy()
    else:
        df_valid = df.copy()

    if len(df_valid) == 0:
        violations.append({
            "rule": "no_valid_data",
            "severity": "critical",
            "message": "Aucune ligne valide après nettoyage",
            "value": 0,
            "column": None,
        })
        return violations

    # --- Validations Grand Livre (GL) ---
    if doc_type == "GL":
        # GL-1 : Vérifier que somme débits ≈ somme crédits
        if "débit" in df_valid.columns and "crédit" in df_valid.columns:
            sum_debit = df_valid["débit"].sum()
            sum_credit = df_valid["crédit"].sum()
            diff = abs(sum_debit - sum_credit)

            if diff > config.tolerance:
                violations.append({
                    "rule": "GL_debit_credit_imbalance",
                    "severity": "warning",
                    "message": (
                        f"Écart débit/crédit : {diff:.2f} "
                        f"(Débits: {sum_debit:.2f}, Crédits: {sum_credit:.2f})"
                    ),
                    "value": diff,
                    "column": None,
                })
            else:
                violations.append({
                    "rule": "GL_debit_credit_balanced",
                    "severity": "info",
                    "message": f"Grand livre équilibré (écart: {diff:.4f})",
                    "value": diff,
                    "column": None,
                })

        # GL-2 : Vérifier les comptes valides (au moins 4 chiffres)
        if "n°compte" in df_valid.columns:
            invalid_accounts = (
                df_valid["n°compte"].astype(str).str.match(r"^\d{4,}$", na=False) == False
            )
            if invalid_accounts.sum() > 0:
                pct = 100 * invalid_accounts.sum() / len(df_valid)
                violations.append({
                    "rule": "GL_invalid_accounts",
                    "severity": "warning",
                    "message": (
                        f"{invalid_accounts.sum()} compte(s) invalide(s) ({pct:.1f}%). "
                        f"Format attendu : au moins 4 chiffres"
                    ),
                    "value": invalid_accounts.sum(),
                    "column": "n°compte",
                })

        # GL-3 : Vérifier les montants (pas de valeurs aberrantes)
        for col in ["débit", "crédit"]:
            if col in df_valid.columns:
                amounts = df_valid[col].dropna()
                if len(amounts) > 0:
                    q95 = amounts.quantile(0.95)
                    outliers = (amounts > q95 * 10).sum()
                    if outliers > 0:
                        violations.append({
                            "rule": f"GL_outlier_{col}",
                            "severity": "info",
                            "message": (
                                f"{outliers} valeur(s) aberrante(s) détectée(s) "
                                f"dans '{col}' (>10x le Q95)"
                            ),
                            "value": outliers,
                            "column": col,
                        })

    # --- Validations Balance Générale (BG) ---
    elif doc_type == "BG":
        if "solde_débit" in df_valid.columns and "solde_crédit" in df_valid.columns:
            sum_deb = df_valid["solde_débit"].sum()
            sum_cre = df_valid["solde_crédit"].sum()
            diff = abs(sum_deb - sum_cre)

            if diff > config.tolerance:
                violations.append({
                    "rule": "BG_solde_imbalance",
                    "severity": "warning",
                    "message": (
                        f"Écart soldes débiteur/créditeur : {diff:.2f} "
                        f"(Débits: {sum_deb:.2f}, Crédits: {sum_cre:.2f})"
                    ),
                    "value": diff,
                    "column": None,
                })
            else:
                violations.append({
                    "rule": "BG_solde_balanced",
                    "severity": "info",
                    "message": f"Balance équilibrée (écart: {diff:.4f})",
                    "value": diff,
                    "column": None,
                })

    # --- Validations Balance Auxiliaire (AUX) ---
    elif doc_type == "AUX":
        # AUX-1 : Vérifier cohérence soldes
        # Règle : si_débit + mvt_débit - mvt_crédit = sf_débit
        expected_cols = ["si_débit", "mvt_débit", "mvt_crédit", "sf_débit"]
        if all(col in df_valid.columns for col in expected_cols):
            df_valid_aux = df_valid[expected_cols].fillna(0)
            expected_sf = (
                df_valid_aux["si_débit"] + df_valid_aux["mvt_débit"] - df_valid_aux["mvt_crédit"]
            )
            actual_sf = df_valid_aux["sf_débit"]
            diff = (actual_sf - expected_sf).abs()
            discrepancies = (diff > config.tolerance).sum()

            if discrepancies > 0:
                pct = 100 * discrepancies / len(df_valid_aux)
                violations.append({
                    "rule": "AUX_solde_coherence",
                    "severity": "warning",
                    "message": (
                        f"{discrepancies} ligne(s) ({pct:.1f}%) avec incohérence de solde. "
                        f"Attendu : SI_débit + MVT_débit - MVT_crédit = SF_débit"
                    ),
                    "value": discrepancies,
                    "column": None,
                })

    # --- Validations générales ---
    # Vérifier la présence de données
    if len(df_valid) == 0:
        violations.append({
            "rule": "no_data",
            "severity": "critical",
            "message": "Aucune donnée valide à traiter",
            "value": 0,
            "column": None,
        })

    logger.info(f"Validation effectuée : {len(violations)} violation(s) trouvée(s)")
    return violations


def build_validation_report(violations: List[Dict]) -> pd.DataFrame:
    """
    Convertit la liste de violations en DataFrame pour rapport Excel.
    
    Args:
        violations: Liste des violations.
    
    Returns:
        DataFrame avec colonnes : rule, severity, message, value, column.
    """
    if not violations:
        return pd.DataFrame(columns=["rule", "severity", "message", "value", "column"])

    df_violations = pd.DataFrame(violations)

    # Trier par sévérité (critical, warning, info)
    severity_order = {"critical": 0, "warning": 1, "info": 2}
    df_violations["_severity_order"] = df_violations["severity"].map(severity_order)
    df_violations = df_violations.sort_values("_severity_order").drop("_severity_order", axis=1)

    return df_violations.reset_index(drop=True)
