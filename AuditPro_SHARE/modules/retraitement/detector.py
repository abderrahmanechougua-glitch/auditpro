"""
Detector : détection du type de document (GL, BG, AUX).
Basée sur un scoring pondéré des colonnes et mots-clés.
"""
import logging
from typing import Dict, Tuple

import pandas as pd

from .config import Config

logger = logging.getLogger(__name__)


def detect_document_type(
    df: pd.DataFrame,
    config: Config,
    sample_rows: int = 100,
) -> Tuple[str, float, Dict]:
    """
    Détecte le type de document comptable (GL, BG, ou AUX).
    
    Stratégie :
    1. Comparer les noms de colonnes aux variantes connus par type.
    2. Chercher des mots-clés métier dans les premières lignes.
    3. Calculer un score pondéré pour chaque type.
    4. Retourner le type gagnant + explication.
    
    Args:
        df: DataFrame avec des colonnes nommées (après détection de l'en-tête).
        config: Configuration contenant les mappings et poids.
        sample_rows: Nombre de lignes à échantillonner pour les mots-clés.
    
    Returns:
        (type_detecté, confiance, explanation_dict)
        Exemple :
        {
            'type': 'GL',
            'confidence': 0.92,
            'reason': 'Colonnes trouvées: journal, pièce, débit, crédit',
            'scores': {'GL': 92, 'BG': 12, 'AUX': 5}
        }
    """
    col_names = [str(c).lower().strip() for c in df.columns]
    col_set = set(col_names)

    # Normaliser les colonnes du DataFrame pour faciliter la comparaison
    df_sample = df.iloc[:min(sample_rows, len(df))].copy()
    df_text = df_sample.astype(str).apply(lambda col: col.str.lower())

    scores = {"GL": 0, "BG": 0, "AUX": 0}
    matched_columns = {"GL": set(), "BG": set(), "AUX": set()}

    # 1. Scoring par présence de colonnes
    for doc_type, mapping in config.column_mappings.items():
        for std_col, variants in mapping.items():
            variants_lower = [v.lower() for v in variants]
            for col in col_names:
                if col in variants_lower or any(v == col for v in variants_lower):
                    scores[doc_type] += config.doc_type_col_weight
                    matched_columns[doc_type].add(std_col)
                    logger.debug(f"Colonne '{col}' → {doc_type}.{std_col}")
                    break

    # 2. Scoring par mots-clés dans le contenu des données
    # Certaines cellules peuvent etre numeriques (float/int) : coercition defensive en str.
    flat_values = df_text.values.flatten()
    content_text = " ".join(str(v) for v in flat_values if pd.notna(v)).lower()

    keyword_groups = {
        "GL": [
            "journal",
            "pièce",
            "piece",
            "écriture",
            "ecriture",
            "contrepartie",
            "débit",
            "credit",
        ],
        "BG": [
            "balance générale",
            "balance generale",
            "solde débit",
            "solde credit",
            "cumul débit",
            "cumul credit",
        ],
        "AUX": [
            "tiers",
            "fournisseur",
            "client",
            "solde final",
            "mouvement",
            "auxiliaire",
        ],
    }

    for doc_type, keywords in keyword_groups.items():
        for keyword in keywords:
            count = content_text.count(keyword)
            scores[doc_type] += count * config.doc_type_keyword_weight

    # 3. Déterminer le gagnant
    best_type = max(scores, key=scores.get)
    total_score = sum(scores.values())
    confidence = scores[best_type] / max(total_score, 1)

    # 4. Générer l'explication
    reason_parts = []
    for doc_type in ["GL", "BG", "AUX"]:
        if matched_columns[doc_type]:
            reason_parts.append(f"{doc_type}: colonnes {sorted(matched_columns[doc_type])}")

    explanation = {
        "type": best_type,
        "confidence": round(confidence, 3),
        "reason": " ; ".join(reason_parts) if reason_parts else "Détection par score",
        "scores": {k: int(v) for k, v in scores.items()},
        "matched_columns": {k: sorted(v) for k, v in matched_columns.items()},
    }

    logger.info(f"Document détecté : {best_type} (confiance: {confidence:.1%})")
    logger.debug(f"Explication : {explanation}")

    return best_type, confidence, explanation
