"""
Cleaner : marquage des lignes problématiques sans suppression.
Utilisation de colonnes de drapeau (_flag) pour traçabilité complète.
"""
import logging
import re
from typing import Optional

import pandas as pd

from .config import Config, GL_CONFIG

logger = logging.getLogger(__name__)


def _find_description_column(df: pd.DataFrame) -> str | None:
    """Retourne la colonne la plus probable pour la description/libelle GL."""
    keywords = GL_CONFIG.get("description_candidates", [])

    for col in df.columns:
        col_lower = str(col).lower().strip()
        if any(k in col_lower for k in keywords):
            return col

    # Fallback: premiere colonne texte avec des libelles suffisamment longs.
    for col in df.columns:
        series = df[col]
        if str(series.dtype) in ("object", "string"):
            sample = series.dropna().astype(str).head(30)
            if not sample.empty and sample.str.len().mean() > 10:
                return col
    return None


def _find_date_column(df: pd.DataFrame) -> Optional[str]:
    """Retourne la colonne la plus probable pour la date/voucher."""
    candidates = ["date", "voucher", "date ecriture", "date écriture", "dt", "jour"]
    for col in df.columns:
        col_lower = str(col).lower().strip()
        if any(c == col_lower or c in col_lower for c in candidates):
            return col
    return None


def _find_amount_column(df: pd.DataFrame) -> Optional[str]:
    """Retourne une colonne montant/solde/accumulation probable (préfère solde cumulé)."""
    # Priorité: accumulated/solde avant debit/credit pour la détection des titres sans montant.
    candidates_priority = ["accum", "solde", "montant", "debit", "credit", "crédit", "débit"]
    for c in candidates_priority:
        for col in df.columns:
            col_lower = str(col).lower().strip()
            if c in col_lower:
                return col
    return None


def flag_gl_lines(df: pd.DataFrame) -> pd.DataFrame:
    """Marque les lignes GL en keep/exclude pour filtrer les lignes non comptables."""
    df_out = df.copy()

    # Recalcul force pour eviter tout decalage avec d'anciens exports contenant keep/_flag.
    if "_flag" in df_out.columns:
        df_out = df_out.drop(columns=["_flag"])
    df_out["_flag"] = "keep"

    # Règles structurantes spécifiques Sage injectées par le loader dédié.
    if "_SAGE_ROW_TYPE" in df_out.columns:
        sage_exclude = df_out["_SAGE_ROW_TYPE"].astype(str).str.lower().isin({"title", "total", "empty", "header"})
        df_out.loc[sage_exclude, "_flag"] = "exclude"

    desc_col = _find_description_column(df_out)
    date_col = _find_date_column(df_out)
    amount_col = _find_amount_column(df_out)

    # 1) Lignes avec pseudo-date invalide (#######, 1120000000--, lignes meta Nominal/Sector...).
    if date_col is not None:
        date_vals = df_out[date_col].astype(str).str.strip()
        # Pattern: ###### ou code_compte-- ou texte non-date evidents
        invalid_date_mask = date_vals.str.contains(
            r"^#+$|^\d+--$|^nominal|^sector|charge type|^period$|^primary dimension",
            na=False, regex=True, flags=re.IGNORECASE
        )
        # Toute valeur non vide non parseable comme date est structurelle pour GL.
        parsed_dates = pd.to_datetime(date_vals, errors="coerce")
        non_empty_date = date_vals.str.strip().ne("") & date_vals.str.lower().ne("nan")
        non_date_mask = non_empty_date & parsed_dates.isna()
        # Sous-en-tetes parasites: valeur de colonne Date == 'date' exactement
        subheader_date_mask = date_vals.str.lower().str.strip() == "date"
        df_out.loc[invalid_date_mask | subheader_date_mask | non_date_mask, "_flag"] = "exclude"
    else:
        invalid_date_mask = pd.Series(False, index=df_out.index)

    if desc_col is not None:
        desc_vals = df_out[desc_col].astype(str).str.lower().str.strip()
    else:
        text_df = df_out.fillna("").astype(str)
        desc_vals = text_df.agg(" ".join, axis=1).str.lower().str.strip()

    # 2) Mots-cles d'exclusion metier + en-tetes residuels.
    exclude_keywords = GL_CONFIG.get("exclude_keywords", [])
    keyword_pattern = "|".join(exclude_keywords)
    keyword_mask = desc_vals.str.contains(keyword_pattern, na=False, regex=True) if keyword_pattern else pd.Series(False, index=df_out.index)
    # En-tetes parasites et lignes footer (ex: "Debit / Credit / Net difference")
    header_mask = desc_vals.str.contains(
        r"^voucher$|^description$|^credit$|^debit$|^net\s", na=False, regex=True
    )
    df_out.loc[keyword_mask | header_mask, "_flag"] = "exclude"

    # 3) Lignes titre de compte (description se terminant par - ou --) sans montant exploitable.
    # Chercher toutes les colonnes numeriques pour verifier absence de montant.
    numeric_cols = [c for c in df_out.columns if c != "_flag" and
                    pd.to_numeric(df_out[c], errors="coerce").notna().sum() > 0]
    if numeric_cols:
        any_amount = df_out[numeric_cols].apply(lambda s: pd.to_numeric(s, errors="coerce")).notna().any(axis=1)
    else:
        any_amount = pd.Series(False, index=df_out.index)
    # Titres: description se termine par un ou plusieurs tirets (- ou --)
    title_desc_mask = desc_vals.str.contains(r"-+\s*$", na=False, regex=True) & ~any_amount
    df_out.loc[title_desc_mask, "_flag"] = "exclude"

    # 4) Lignes completement vides (ignorer les colonnes internes propagées).
    data_only = df_out.drop(columns=["_flag", "_COMPTE", "_LIBELLE_COMPTE"], errors="ignore")
    empty_mask = data_only.replace(r"^\s*$", pd.NA, regex=True).isna().all(axis=1)
    df_out.loc[empty_mask, "_flag"] = "exclude"

    logger.info(
        "GL cleaner: desc_col=%s | date_col=%s | amount_col=%s | keep=%s | exclude=%s",
        desc_col,
        date_col,
        amount_col,
        int((df_out["_flag"] == "keep").sum()),
        int((df_out["_flag"] == "exclude").sum()),
    )
    return df_out


def _row_as_text(row: pd.Series) -> str:
    """Convertit une ligne heterogene en texte robuste pour la recherche de mots-cles."""
    return " ".join(str(v) for v in row.tolist() if pd.notna(v))


def flag_rows(df: pd.DataFrame, config: Config) -> pd.DataFrame:
    """
    Ajoute une colonne '_flag' pour marquer les lignes problématiques.
    
    Valeurs possibles :
    - 'keep' : ligne normale à traiter
    - 'empty' : ligne entièrement vide
    - 'total' : ligne contenant un mot-clé total/sous-total/report
    - 'header_like' : ligne ressemblant à un en-tête supplémentaire
    - 'noise' : ligne détectée comme bruit (peu de données)
    
    Aucune suppression n'est effectuée. La colonne '_flag' peut être utilisée
    pour le filtrage ultérieur.
    
    Args:
        df: DataFrame à marquer.
        config: Configuration avec les mots-clés bannissables.
    
    Returns:
        DataFrame avec la colonne '_flag' ajoutée.
    """
    df = df.copy()

    # Initialiser le flag par défaut
    df["_flag"] = "keep"

    # 1. Marquer les lignes complètement vides
    is_empty = df.isna().all(axis=1)
    df.loc[is_empty, "_flag"] = "empty"
    logger.info(f"Marquées {is_empty.sum()} lignes vides")

    # 2. Marquer les lignes contenant des mots-clés 'total'
    if len(config.total_keywords) > 0:
        df_str = df.astype(str).apply(lambda col: col.str.upper())
        # Créer un masque pour chaque keyword
        total_mask = pd.Series(False, index=df.index)
        for keyword in config.total_keywords:
            total_mask |= df_str.apply(lambda row: keyword.upper() in _row_as_text(row), axis=1)

        # Ne marquer comme 'total' que les lignes non vides
        total_mask &= ~is_empty
        df.loc[total_mask, "_flag"] = "total"
        logger.info(f"Marquées {total_mask.sum()} lignes 'total'")

    # 3. Marquer les lignes qui ressemblent à un bruit (heuristique)
    # Une ligne de bruit : peu de colonnes remplies ET peu de caractères
    non_null_count = df.notna().sum(axis=1)
    text_length = df.apply(lambda row: len(_row_as_text(row)), axis=1)
    is_noise = (non_null_count <= 1) & (text_length < 20)
    is_noise &= ~is_empty  # Ne pas surcharger 'noise' sur 'empty'
    df.loc[is_noise, "_flag"] = "noise"
    logger.info(f"Marquées {is_noise.sum()} lignes bruit")

    # 4. Marquer les lignes ressemblant à un en-tête supplémentaire
    # Heuristique : beaucoup de texte, peu de nombres, peu de colonnes non-null
    is_text_heavy = (
        df.astype(str).apply(lambda row: row.str.contains(r"[a-zA-ZÀ-ÿ]", regex=True).sum(), axis=1)
        >= (len(df.columns) * 0.7)
    )
    is_sparse = non_null_count <= (len(df.columns) * 0.5)
    is_header_like = is_text_heavy & is_sparse & ~is_empty & ~(df["_flag"] != "keep")
    df.loc[is_header_like, "_flag"] = "header_like"
    logger.info(f"Marquées {is_header_like.sum()} lignes ressemblant à un en-tête")

    # Statistiques finales
    flag_counts = df["_flag"].value_counts()
    logger.info(f"Distribution des drapeaux :\n{flag_counts}")

    return df


def get_clean_data(df: pd.DataFrame, remove_empty: bool = True, remove_total: bool = False) -> pd.DataFrame:
    """
    Filtre les lignes marquées selon les paramètres.
    
    Utilité : pour les utilisateurs qui veulent un DataFrame "propre" sans
    les lignes problématiques.
    
    Args:
        df: DataFrame avec colonne '_flag'.
        remove_empty: Si True, supprimer les lignes 'empty'.
        remove_total: Si True, supprimer les lignes 'total'.
    
    Returns:
        DataFrame filtré.
    """
    df_clean = df.copy()

    filters_to_remove = []
    if remove_empty:
        filters_to_remove.append("empty")
    if remove_total:
        filters_to_remove.append("total")

    removed_count = 0
    for filter_name in filters_to_remove:
        mask = df_clean["_flag"] == filter_name
        removed_count += mask.sum()
        df_clean = df_clean[~mask]

    logger.info(f"Supprimées {removed_count} lignes selon les filtres")
    return df_clean.reset_index(drop=True)
