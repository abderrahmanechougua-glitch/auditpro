"""
Reporter : génération des rapports Excel avec audit trail.
Organise les résultats dans plusieurs feuilles thématiques.
"""

# Guard contre l'exécution directe avec relative imports
if __name__ == '__main__':
    import sys
    print("[x] Erreur : Ce module doit être importé, pas exécuté directement.")
    print("[x] Utilisez : python -c \"from modules.retraitement import process_file\"")
    sys.exit(1)
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, List

import pandas as pd

logger = logging.getLogger(__name__)


def build_transformation_log(
    df: pd.DataFrame,
    metadata: Dict,
    doc_type_explanation: Dict,
    column_mappings: Dict,
    normalization_summary: Dict,
    validation_report: pd.DataFrame,
) -> pd.DataFrame:
    """
    Construit un log d'audit complet pour traçabilité.
    
    Combine toutes les décisions de transformation en une seule trace.
    
    Args:
        df: DataFrame final.
        metadata: Métadonnées du fichier (source, feuille, en-tête).
        doc_type_explanation: Résultat de la détection de type.
        column_mappings: Dict des mappages appliqués.
        normalization_summary: Dict des conversions effectuées.
        validation_report: DataFrame des violations.
    
    Returns:
        DataFrame contenant l'audit trail complet.
    """
    log_entries = []

    # 1. Métadonnées du fichier
    log_entries.append({
        "timestamp": datetime.now().isoformat(),
        "phase": "input",
        "action": "file_loaded",
        "details": f"Fichier : {metadata.get('file_path')} | Feuille : {metadata.get('sheet_name')}",
        "value": metadata.get("original_rows"),
    })

    # 2. Détection du type
    log_entries.append({
        "timestamp": datetime.now().isoformat(),
        "phase": "detection",
        "action": "doc_type_detected",
        "details": (
            f"Type : {doc_type_explanation.get('type')} | "
            f"Confiance : {doc_type_explanation.get('confidence'):.1%} | "
            f"Raison : {doc_type_explanation.get('reason')}"
        ),
        "value": doc_type_explanation.get("confidence"),
    })

    # 3. Standardisation des colonnes
    for orig_col, std_col in column_mappings.items():
        log_entries.append({
            "timestamp": datetime.now().isoformat(),
            "phase": "normalization",
            "action": "column_renamed",
            "details": f"'{orig_col}' → '{std_col}'",
            "value": None,
        })

    # 4. Conversions de types
    for key, count in normalization_summary.items():
        if count > 0:
            log_entries.append({
                "timestamp": datetime.now().isoformat(),
                "phase": "normalization",
                "action": "data_converted",
                "details": f"{key}: {count} colonne(s)",
                "value": count,
            })

    # 5. Violations de validation
    if not validation_report.empty:
        for _, violation in validation_report.iterrows():
            log_entries.append({
                "timestamp": datetime.now().isoformat(),
                "phase": "validation",
                "action": f"validation_{violation['rule']}",
                "details": f"[{violation['severity'].upper()}] {violation['message']}",
                "value": violation.get("value"),
            })

    # 6. Statut des lignes (flagging)
    if "_flag" in df.columns:
        flag_counts = df["_flag"].value_counts()
        for flag_name, count in flag_counts.items():
            log_entries.append({
                "timestamp": datetime.now().isoformat(),
                "phase": "cleaning",
                "action": f"rows_flagged_{flag_name}",
                "details": f"{count} ligne(s) marquée(s) comme '{flag_name}'",
                "value": count,
            })

    return pd.DataFrame(log_entries)


def generate_report(
    df: pd.DataFrame,
    metadata: Dict,
    doc_type_explanation: Dict,
    column_mappings: Dict,
    normalization_summary: Dict,
    validation_report: pd.DataFrame,
    output_dir: Path,
    filename_stem: str = "Retraitement",
) -> str:
    """
    Génère un rapport Excel multi-feuilles avec traçabilité complète.
    
    Feuilles générées :
    - Données_Normalisées : données traitées (avec colonnes _flag)
    - Validation : rapport des règles violées
    - Transformation_Log : audit trail complet
    - Métadonnées : informations sur le fichier source et le traitement
    
    Args:
        df: DataFrame normalisé.
        metadata: Métadonnées du fichier.
        doc_type_explanation: Résultat de la détection.
        column_mappings: Mappages de colonnes appliqués.
        normalization_summary: Résumé des conversions.
        validation_report: Rapport de validation.
        output_dir: Répertoire de sortie.
        filename_stem: Préfixe du nom du fichier.
    
    Returns:
        Chemin du fichier généré.
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = output_dir / f"{filename_stem}_{timestamp}.xlsx"

    # Construire le log de transformation
    transformation_log = build_transformation_log(
        df,
        metadata,
        doc_type_explanation,
        column_mappings,
        normalization_summary,
        validation_report,
    )

    logger.info(f"Génération du rapport : {output_file}")

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        # Feuille 1 : Données normalisées
        df.to_excel(writer, sheet_name="Données_Normalisées", index=False)
        logger.info(f"  - Feuille 'Données_Normalisées' : {len(df)} lignes")

        # Feuille 2 : Validation
        if not validation_report.empty:
            validation_report.to_excel(writer, sheet_name="Validation", index=False)
            logger.info(f"  - Feuille 'Validation' : {len(validation_report)} violation(s)")
        else:
            pd.DataFrame({"Message": ["Aucune violation détectée"]}).to_excel(
                writer, sheet_name="Validation", index=False
            )

        # Feuille 3 : Log de transformation
        transformation_log.to_excel(writer, sheet_name="Transformation_Log", index=False)
        logger.info(
            f"  - Feuille 'Transformation_Log' : {len(transformation_log)} entrée(s)"
        )

        # Feuille 4 : Métadonnées
        metadata_df = pd.DataFrame([
            {
                "Propriété": "Fichier source",
                "Valeur": metadata.get("file_path", "N/A"),
            },
            {
                "Propriété": "Feuille utilisée",
                "Valeur": metadata.get("sheet_name", "N/A"),
            },
            {
                "Propriété": "Ligne d'en-tête détectée",
                "Valeur": metadata.get("header_row_idx", "N/A"),
            },
            {
                "Propriété": "Lignes brutes",
                "Valeur": metadata.get("original_rows", "N/A"),
            },
            {
                "Propriété": "Colonnes brutes",
                "Valeur": metadata.get("original_columns", "N/A"),
            },
            {
                "Propriété": "Lignes finales",
                "Valeur": len(df),
            },
            {
                "Propriété": "Colonnes finales",
                "Valeur": len(df.columns),
            },
            {
                "Propriété": "Type détecté",
                "Valeur": doc_type_explanation.get("type", "N/A"),
            },
            {
                "Propriété": "Confiance détection",
                "Valeur": f"{doc_type_explanation.get('confidence', 0):.1%}",
            },
            {
                "Propriété": "Raison détection",
                "Valeur": doc_type_explanation.get("reason", "N/A"),
            },
            {
                "Propriété": "Colonnes mappées",
                "Valeur": len(column_mappings),
            },
            {
                "Propriété": "Rapport généré à",
                "Valeur": datetime.now().isoformat(),
            },
        ])
        metadata_df.to_excel(writer, sheet_name="Métadonnées", index=False)
        logger.info("  - Feuille 'Métadonnées'")

    logger.info(f"Rapport généré avec succès : {output_file}")
    return str(output_file)
