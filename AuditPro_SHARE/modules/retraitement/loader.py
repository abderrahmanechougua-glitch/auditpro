"""
Loader : chargement et détection de l'en-tête des fichiers Excel.
Vectorisé pour éviter les itérations cellule par cellule.
"""
import logging
from pathlib import Path
from typing import Optional, Tuple

import pandas as pd

logger = logging.getLogger(__name__)


def find_header_row(df: pd.DataFrame, max_rows: int = 50) -> Optional[int]:
    """
    Détecte la ligne d'en-tête probable dans un DataFrame brut (sans header).
    
    Stratégie vectorisée :
    - Calculer pour chaque ligne : ratio texte vs numérique, présence de mots-clés métier.
    - Sélectionner la ligne avec le meilleur score (sans parcourir cellule par cellule).
    
    Args:
        df: DataFrame brut (header=None)
        max_rows: nombre max de lignes à examiner (performance)
    
    Returns:
        Index de la ligne d'en-tête (0-based), ou None si non détectée.
    """
    df_sample = df.iloc[:min(max_rows, len(df))].copy()

    # 1. Compter les valeurs non-NaN par ligne
    non_empty_counts = df_sample.notna().sum(axis=1)

    # 2. Convertir les cellules en strings et analyser le contenu
    df_str = df_sample.astype(str)

    # Ratio de cellules "textuelles" (contiennent des lettres) vs numériques
    # Une ligne d'en-tête typique a beaucoup de texte, peu de nombres purs
    is_text = df_str.apply(lambda s: s.str.contains(r"[a-zA-ZÀ-ÿ]", regex=True)).sum(axis=1)
    is_numeric = df_str.apply(lambda s: s.str.match(r"^-?\d+[\d,.\s]*$", na=False)).sum(axis=1)
    text_ratio = is_text / (non_empty_counts + 1)  # +1 pour éviter division par 0

    # 3. Détecter les mots-clés métier comptable dans chaque ligne
    header_keywords = [
        "compte",
        "date",
        "libelle",
        "solde",
        "debit",
        "credit",
        "tiers",
        "journal",
        "piece",
        "montant",
        "ref",
        "pièce",
        "intitule",
    ]
    keyword_matches = pd.Series(0, index=df_sample.index, dtype=int)
    for keyword in header_keywords:
        keyword_matches += (
            df_str.astype(str)
            .apply(lambda row: row.str.lower().str.contains(keyword, na=False).any(), axis=1)
            .astype(int)
        )

    # 4. Les lignes suivantes contiennent plus de données numériques (validation)
    next_rows_numeric = pd.Series(0, index=df_sample.index, dtype=int)
    for i in range(len(df_sample)):
        next_rows = df_sample.iloc[i + 1 : i + 4]
        if not next_rows.empty:
            next_numeric_count = (
                next_rows.astype(str)
                .apply(lambda row: row.str.match(r"^-?\d+[\d,.\s]*$", na=False).sum(), axis=1)
                .sum()
            )
            next_rows_numeric[i] = next_numeric_count

    # 5. Score composite
    # Poids : texte + keywords + min(non_empty, 6) + présence données après
    scores = (
        text_ratio * 5
        + (keyword_matches > 0).astype(int) * 8
        + (non_empty_counts / non_empty_counts.max() if non_empty_counts.max() > 0 else 0) * 3
        + (next_rows_numeric > 2).astype(int) * 4
    )

    # 6. Sélectionner le meilleur candidat
    if scores.max() <= 0:
        return None

    best_idx = scores.idxmax()
    logger.info(f"En-tête détecté à la ligne {best_idx} (score: {scores[best_idx]:.2f})")
    return best_idx


def load_excel(
    file_path: Path,
    header_detection: bool = True,
    encoding: str = "utf-8",
) -> Tuple[pd.DataFrame, Optional[int], dict]:
    """
    Charge un fichier Excel et détecte la ligne d'en-tête.
    
    Args:
        file_path: Chemin du fichier Excel.
        header_detection: Si True, détecte automatiquement la ligne d'en-tête.
        encoding: Encodage (non utilisé pour Excel, pour compatibilité future CSV).
    
    Returns:
        (df_raw, header_row_idx, metadata)
        - df_raw : DataFrame brut (sans header si detection, sinon avec header=0)
        - header_row_idx : Index de la ligne d'en-tête (0-based), ou None
        - metadata : Dict avec infos (sheet_name, original_rows, etc.)
    """
    file_path = Path(file_path)

    if not file_path.exists():
        raise FileNotFoundError(f"Fichier non trouvé : {file_path}")

    ext = file_path.suffix.lower()
    if ext not in (".xlsx", ".xls", ".xlsm", ".xlsb"):
        raise ValueError(f"Format de fichier non supporté : {file_path.suffix}")

    try:
        if ext == ".xlsb":
            try:
                import pyxlsb  # noqa: F401
            except Exception as dep_error:
                raise RuntimeError(
                    "Format .xlsb détecté mais le package 'pyxlsb' est absent. "
                    "Installez-le (pip install pyxlsb) ou convertissez le fichier en .xlsx."
                ) from dep_error

        # Lire le fichier pour déterminer les feuilles disponibles
        xl_kwargs = {"engine": "pyxlsb"} if ext == ".xlsb" else {}
        xl_file = pd.ExcelFile(file_path, **xl_kwargs)
        sheet_names = xl_file.sheet_names
        logger.info(f"Feuilles trouvées : {sheet_names}")

        # Essayer chaque feuille
        df_raw = None
        selected_sheet = None

        for sheet_name in sheet_names:
            try:
                read_kwargs = {"engine": "pyxlsb"} if ext == ".xlsb" else {}
                df_candidate = pd.read_excel(file_path, sheet_name=sheet_name, header=None, **read_kwargs)

                # Filtrer les lignes complètement vides (heuristique : doit avoir au moins 2 colonnes non-vides)
                rows_with_data = df_candidate.notna().sum(axis=1) >= 2
                if rows_with_data.any():
                    df_raw = df_candidate
                    selected_sheet = sheet_name
                    logger.info(f"Utilisée la feuille : {sheet_name}")
                    break
            except Exception as e:
                logger.warning(f"Erreur lecture feuille '{sheet_name}' : {e}")
                continue

        if df_raw is None:
            raise ValueError(
                f"Aucune feuille valide trouvée dans {file_path}. "
                f"Feuilles disponibles : {sheet_names}"
            )

        original_rows = len(df_raw)
        logger.info(f"Données brutes chargées : {original_rows} lignes, {len(df_raw.columns)} colonnes")

        # Détecter la ligne d'en-tête
        header_row_idx = None
        if header_detection:
            header_row_idx = find_header_row(df_raw)
            if header_row_idx is not None:
                # Réassigner les colonnes avec les valeurs de la ligne d'en-tête
                header_row = df_raw.iloc[header_row_idx]
                df_raw.columns = header_row.fillna("").astype(str)
                logger.info(f"Colonnes assignées depuis la ligne {header_row_idx}")

        metadata = {
            "file_path": str(file_path),
            "sheet_name": selected_sheet,
            "header_row_idx": header_row_idx,
            "original_rows": original_rows,
            "original_columns": len(df_raw.columns),
        }

        return df_raw, header_row_idx, metadata

    except Exception as e:
        logger.error(f"Erreur chargement {file_path} : {e}")
        raise
