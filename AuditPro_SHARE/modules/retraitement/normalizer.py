"""
Normalizer : standardisation des colonnes, parsing des montants et dates.
Entièrement vectorisé pour performance sur 500k+ lignes.
"""
import logging
import re
import unicodedata
from difflib import get_close_matches
from typing import Dict, List, Optional, Tuple

import pandas as pd

from .config import Config, GL_CONFIG

logger = logging.getLogger(__name__)


def _normalize_label(text: str) -> str:
    text = "" if text is None else str(text)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]", "", text.lower())


def _find_matching_column(columns: List[str], candidates: List[str]) -> Optional[str]:
    if not columns:
        return None
    normalized_candidates = {_normalize_label(c) for c in candidates}
    for col in columns:
        if _normalize_label(col) in normalized_candidates:
            return col
    for col in columns:
        col_norm = _normalize_label(col)
        if any(cand in col_norm for cand in normalized_candidates if cand):
            return col
    return None


def _parse_numeric_series(series: pd.Series) -> pd.Series:
    values = series.astype(str).str.strip()
    values = values.str.replace(r"[\s\u00A0]", "", regex=True)
    values = values.str.replace(r"[€$MADdhDHS]", "", regex=True)
    values = values.str.replace(r"\(([^)]+)\)", r"-\1", regex=True)

    has_comma = values.str.contains(",", regex=False, na=False)
    has_dot = values.str.contains("\\.", regex=True, na=False)

    euro_mask = has_comma & has_dot
    values.loc[euro_mask] = values.loc[euro_mask].str.replace(".", "", regex=False).str.replace(",", ".", regex=False)

    comma_only = has_comma & ~has_dot
    values.loc[comma_only] = values.loc[comma_only].str.replace(",", ".", regex=False)

    values = values.str.replace(r"[^0-9.\-]", "", regex=True)
    return pd.to_numeric(values, errors="coerce").fillna(0.0)


def _parse_date_series(series: pd.Series) -> pd.Series:
    raw = series.astype(str).str.strip()
    # Format ISO avec ou sans heure : 2025-12-31 ou 2025-12-31 00:00:00
    iso_mask = raw.str.match(r"^\d{4}[-/]\d{2}[-/]\d{2}(\s+\d{2}:\d{2}(:\d{2})?)?$", na=False)
    iso_dates = pd.to_datetime(raw.where(iso_mask), errors="coerce", dayfirst=False)
    local_dates = pd.to_datetime(raw.where(~iso_mask), errors="coerce", dayfirst=True)
    return iso_dates.combine_first(local_dates)


def normalize_gl_to_standard(
    df: pd.DataFrame,
    compte_col: Optional[str] = None,
    compte_value: Optional[str] = None,
) -> pd.DataFrame:
    """
    Normalise un GL vers le format standard 12 colonnes AuditPro :
    C1, C4, COMPTE, LIBELLÉ COMPTE, DATE, N° DOCUMENT, DESCRIPTION,
    TYPE, DÉBIT, CRÉDIT, cumul, solde.
    """
    work = df.copy().astype(object)
    columns = [str(c) for c in work.columns]

    date_col = _find_matching_column(columns, GL_CONFIG["date_candidates"])
    desc_col = _find_matching_column(columns, GL_CONFIG["description_candidates"])
    debit_col = _find_matching_column(columns, GL_CONFIG["debit_candidates"])
    credit_col = _find_matching_column(columns, GL_CONFIG["credit_candidates"])
    accumulation_col = _find_matching_column(columns, GL_CONFIG["accumulation_candidates"])
    voucher_col = _find_matching_column(columns, GL_CONFIG.get("voucher_candidates", ["voucher", "n° document", "piece"]))
    account_col = _find_matching_column(columns, GL_CONFIG["account_candidates"])

    column_mapping = {
        "DATE": date_col,
        "N° DOCUMENT": voucher_col,
        "DESCRIPTION": desc_col,
        "DÉBIT": debit_col,
        "CRÉDIT": credit_col,
        "cumul": accumulation_col,
    }
    logger.info("GL normalizer mapping: %s", column_mapping)

    if date_col is None:
        raise ValueError("Colonne date introuvable pour le GL")

    # --- DATE ---
    date_values = _parse_date_series(work[date_col])

    # --- DESCRIPTION ---
    description_raw = work[desc_col] if desc_col and desc_col in work.columns else pd.Series("", index=work.index)
    description = (
        pd.Series(description_raw, index=work.index)
        .astype(str)
        .str.replace(GL_CONFIG["description_clean_regex"], " ", regex=True)
        .str.replace(GL_CONFIG["spaces_regex"], " ", regex=True)
        .str.strip()
    )

    # --- N° DOCUMENT ---
    if voucher_col and voucher_col in work.columns:
        n_document = work[voucher_col].astype(str).str.strip().replace("nan", "").replace("NaN", "")
    else:
        n_document = pd.Series("", index=work.index, dtype=object)

    # --- COMPTE & LIBELLÉ COMPTE : préférer colonnes propagées par _load_gl_d365 ---
    if "_COMPTE" in work.columns:
        compte = work["_COMPTE"].astype(str).str.strip()
    elif account_col and account_col in work.columns:
        compte = work[account_col].astype(str).str.strip().replace("nan", "")
    else:
        compte = pd.Series(compte_value or "", index=work.index, dtype=object)

    if "_LIBELLE_COMPTE" in work.columns:
        libelle_compte = work["_LIBELLE_COMPTE"].astype(str).str.strip()
    else:
        libelle_compte = pd.Series("", index=work.index, dtype=object)

    # --- C1 (classe comptable) et C4 (sous-classe 4 chiffres) ---
    compte_digits = compte.str.extract(r"^(\d+)", expand=False).fillna("")
    c1 = compte_digits.str[:1]
    c4 = compte_digits.str[:4]

    # --- MONTANTS ---
    format_detected = "unknown"
    if debit_col and credit_col and debit_col in work.columns and credit_col in work.columns:
        format_detected = "debit_credit"
        debit = _parse_numeric_series(work[debit_col]).clip(lower=0)
        credit = _parse_numeric_series(work[credit_col]).clip(lower=0)
    elif accumulation_col and accumulation_col in work.columns:
        format_detected = "accumulation"
        acc = _parse_numeric_series(work[accumulation_col])
        delta = acc.diff().fillna(acc.iloc[0] if len(acc) else 0.0)
        debit = delta.clip(lower=0)
        credit = (-delta).clip(lower=0)
    else:
        raise ValueError("Impossible de détecter un format GL (débit/crédit ou accumulation)")

    # --- CUMUL (solde cumulé brut depuis la colonne Accumulated) ---
    if accumulation_col and accumulation_col in work.columns:
        cumul = _parse_numeric_series(work[accumulation_col])
    else:
        cumul = (debit - credit).cumsum()

    # --- SOLDE (impact net de l'écriture : Crédit - Débit) ---
    solde = credit - debit

    # --- TYPE ---
    desc_lower = description.str.lower().str.strip()
    type_col_vals = pd.Series("MOUVEMENT", index=work.index, dtype=object)
    type_col_vals[desc_lower.str.contains("opening balance", na=False)] = "OPENING BALANCE"
    type_col_vals[desc_lower.str.contains("closing balance", na=False)] = "CLOSING BALANCE"

    standard = pd.DataFrame(
        {
            "C1": c1,
            "C4": c4,
            "COMPTE": compte,
            "LIBELLÉ COMPTE": libelle_compte,
            "DATE": date_values,
            "N° DOCUMENT": n_document,
            "DESCRIPTION": description,
            "TYPE": type_col_vals,
            "DÉBIT": debit.astype(float),
            "CRÉDIT": credit.astype(float),
            "cumul": cumul.astype(float),
            "solde": solde.astype(float),
        },
        index=work.index,
    )

    # Conserver l'ordre original du fichier (déjà trié par compte puis date dans D365)
    standard = standard.reset_index(drop=True)

    standard.attrs["gl_format_detected"] = format_detected
    standard.attrs["column_mapping"] = column_mapping
    logger.info("GL normalizer format detecte: %s | lignes: %d", format_detected, len(standard))
    return standard


def _column_positions(df: pd.DataFrame, col_name: str) -> List[int]:
    """Retourne toutes les positions d'une colonne (gere les doublons de noms)."""
    return [i for i, c in enumerate(df.columns) if c == col_name]


def _detect_amount_format(df: pd.DataFrame, sample_size: int = 100) -> str:
    """
    Détecte automatiquement le format dominant des montants.
    
    Stratégie : examiner les premières colonnes suspectes, chercher le format
    dominant (européen 1.234,56 vs américain 1,234.56).
    
    Args:
        df: DataFrame.
        sample_size: Nombre de lignes à examiner.
    
    Returns:
        Format détecté : "european", "american", ou "numeric" (pas de séparateur).
    """
    # Sélectionner les colonnes potentiellement numéritiques
    numeric_cols = []
    for i, col in enumerate(df.columns):
        # Chercher des colonnes contenant des patterns numériques
        col_str = df.iloc[:, i].astype(str).str.lower()
        if (col_str.str.contains(r"\d", regex=True).sum() > (len(df) * 0.3)):
            numeric_cols.append(i)

    if not numeric_cols:
        return "numeric"

    # Examiner les premières lignes de ces colonnes
    sample = df.iloc[:, numeric_cols].head(min(sample_size, len(df))).astype(str)

    # Compter les formats
    european_count = 0  # Pattern : 1.234,56 ou 1,00
    american_count = 0  # Pattern : 1,234.56

    for i in range(len(numeric_cols)):
        values = sample.iloc[:, i]
        for val in values:
            if pd.isna(val) or val in ("nan", ""):
                continue
            # Chercher les patterns
            if re.search(r"\d+\.\d+,\d{2}", val):  # 1.234,56
                european_count += 1
            elif re.search(r"\d+,\d+\.\d{2}", val):  # 1,234.56
                american_count += 1

    if european_count > american_count:
        logger.info(f"Format détecté : européen (virgule décimale)")
        return "european"
    elif american_count > european_count:
        logger.info(f"Format détecté : américain (point décimal)")
        return "american"
    else:
        logger.info(f"Format détecté : numérique pur (pas de séparateur)")
        return "numeric"


def _parse_amounts(
    series: pd.Series,
    format_detected: str = "european",
    column_name: str = "",
) -> Tuple[pd.Series, List[str]]:
    """
    Parse les montants en floats, gérant les formats européen/américain.
    
    Vectorisé : aucune boucle Python sur les cellules.
    
    Args:
        series: Série pandas contenant les montants à parser.
        format_detected: Format à utiliser.
        column_name: Nom de la colonne (pour les logs).
    
    Returns:
        (series_parsed, warnings_list)
        Les valeurs non parsées deviennent NaN.
    """
    warnings = []
    result = series.copy()

    # Convertir en strings
    result_str = result.astype(str).str.strip()

    # Normaliser les valeurs "nan"
    mask_nan = result_str.isin(("nan", "none", "", "-", "–", "–", "NaN", "None"))
    result_str[mask_nan] = None

    # Detecter les montants négatifs entre parenthèses
    mask_negative = result_str.str.match(r"^\(\s*[\d,.\s]+\s*\)$", na=False)
    result_str[mask_negative] = "-" + result_str[mask_negative].str.replace(r"[()]", "", regex=True)

    # Nettoyer les caractères non essentiels
    result_str = result_str.str.replace(r"[MAD€\s]", "", regex=True)

    # Conversion selon le format
    if format_detected == "european":
        # Format : 1.234,56 → 1234.56
        result_str = result_str.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    elif format_detected == "american":
        # Format : 1,234.56 → 1234.56
        result_str = result_str.str.replace(",", "", regex=False)
    # Sinon, format numérique pur, pas de modification

    # Conversion en float
    result_parsed = pd.to_numeric(result_str, errors="coerce")

    # Logger les valeurs non converties
    mask_failed = result.notna() & result_parsed.isna() & ~mask_nan
    failed_count = mask_failed.sum()
    if failed_count > 0:
        failed_indices = result.index[mask_failed].tolist()[:5]  # Premiers 5
        warnings.append(
            f"Colonne '{column_name}' : {failed_count} valeur(s) non parsée(s). "
            f"Exemples (indices {failed_indices}) : {result[mask_failed].head(3).tolist()}"
        )
        logger.warning(warnings[-1])

    return result_parsed, warnings


def _detect_date_format(df: pd.DataFrame, date_col: str, sample_size: int = 50) -> bool:
    """
    Détecte si le format de date est jour-en-premier (dayfirst=True).
    
    Heuristique : si la plupart des valeurs ont le premier nombre > 12, c'est jour-en-premier.
    
    Args:
        df: DataFrame.
        date_col: Nom de la colonne.
        sample_size: Nombre de lignes à examiner.
    
    Returns:
        dayfirst : bool.
    """
    date_positions = _column_positions(df, date_col)
    if not date_positions:
        return True

    # En cas de doublon de colonnes 'date', on prend la premiere colonne pour l'heuristique.
    sample = df.iloc[:, date_positions[0]].head(min(sample_size, len(df))).astype(str)

    # Extraire les premiers nombres
    first_numbers = sample.str.extract(r"(\d+)", expand=False).astype(float)
    valid_first = first_numbers.dropna()

    if len(valid_first) > 0:
        # Si la majorité des premiers nombres > 12, c'est jour-en-premier
        dayfirst = (valid_first > 12).sum() > (len(valid_first) * 0.5)
        logger.info(f"Format date détecté : dayfirst={dayfirst}")
        return dayfirst

    return True  # Par défaut


def standardize_columns(
    df: pd.DataFrame,
    doc_type: str,
    config: Config,
) -> Tuple[pd.DataFrame, Dict[str, str], List[str]]:
    """
    Standardise les noms des colonnes selon le type de document.
    
    Stratégie :
    1. Correspondance exacte (après normalisation).
    2. Correspondance floue (difflib) si seuil atteint.
    3. Loguer chaque décision.
    
    Args:
        df: DataFrame avec colonnes brutes.
        doc_type: Type détecté (GL, BG, AUX).
        config: Configuration.
    
    Returns:
        (df_renamed, mappings_applied, warnings)
        - df_renamed : DataFrame avec colonnes standardisées.
        - mappings_applied : Dict des renommages appliqués.
        - warnings : Liste des avertissements.
    """
    # Travailler en dtype object evite les erreurs d'affectation (StringDtype -> float/datetime).
    df = df.copy().astype(object)
    warnings = []
    mappings = {}

    if doc_type not in config.column_mappings:
        warnings.append(f"Type de document '{doc_type}' non reconnu")
        return df, mappings, warnings

    mapping_rules = config.column_mappings[doc_type]
    col_names = [str(c).strip() for c in df.columns]

    # Normaliser les noms pour comparaison
    col_names_lower = [c.lower() for c in col_names]
    col_names_normalized = [
        re.sub(r"[^\w]", "", c, flags=re.UNICODE) for c in col_names_lower
    ]

    used_cols = set()  # Éviter les collisions

    for std_col, variants in mapping_rules.items():
        variants_lower = [v.lower() for v in variants]
        variants_normalized = [
            re.sub(r"[^\w]", "", v, flags=re.UNICODE) for v in variants_lower
        ]

        matched_col = None

        # 1. Correspondance exacte
        for i, (col_orig, col_norm) in enumerate(zip(col_names_lower, col_names_normalized)):
            if col_orig in variants_lower or col_norm in variants_normalized:
                if col_names[i] not in used_cols:
                    matched_col = col_names[i]
                    logger.debug(f"Correspondance exacte : '{col_names[i]}' → '{std_col}'")
                    break

        # 2. Correspondance floue (difflib)
        if not matched_col:
            close_matches = get_close_matches(
                col_names_lower if col_names_lower else [],
                variants_lower,
                n=1,
                cutoff=config.column_match_threshold,
            )
            if close_matches:
                # Retrouver la colonne d'origine
                for i, col_lower in enumerate(col_names_lower):
                    if col_lower == close_matches[0]:
                        if col_names[i] not in used_cols:
                            matched_col = col_names[i]
                            logger.debug(
                                f"Correspondance floue : '{col_names[i]}' (score {config.column_match_threshold}) → '{std_col}'"
                            )
                        break

        if matched_col:
            mappings[matched_col] = std_col
            used_cols.add(matched_col)
        else:
            warnings.append(f"Colonne standard '{std_col}' non trouvée dans le fichier")
            logger.warning(warnings[-1])

    # Appliquer les renommages
    df = df.rename(columns=mappings)
    logger.info(f"Standardisation colonnes : {len(mappings)} mappages appliqués")

    return df, mappings, warnings


def normalize_data_types(
    df: pd.DataFrame,
    doc_type: str,
    config: Config,
) -> Tuple[pd.DataFrame, Dict, List[str]]:
    """
    Normalise les types de données (montants → float, dates → datetime).
    
    Vectorisé : pas de apply(axis=1).
    
    Args:
        df: DataFrame avec colonnes standardisées.
        doc_type: Type détecté (GL, BG, AUX).
        config: Configuration.
    
    Returns:
        (df_normalized, transformation_summary, warnings)
    """
    df = df.copy().astype(object)
    warnings = []
    summary = {"amounts_parsed": 0, "dates_parsed": 0, "columns_processed": 0}

    # Déterminer le format des montants (une fois pour tout le fichier)
    if config.amount_format_detection == "auto":
        format_detected = _detect_amount_format(df, config.amount_format_sample_size)
    else:
        format_detected = config.amount_format_detection

    # Sélectionner les colonnes à traiter selon le type
    amount_cols = []
    if doc_type == "GL":
        amount_cols = ["débit", "crédit"]
    elif doc_type == "BG":
        amount_cols = ["solde_débit", "solde_crédit"]
    elif doc_type == "AUX":
        amount_cols = [
            "sf_débit",
            "sf_crédit",
            "si_débit",
            "si_crédit",
            "mvt_débit",
            "mvt_crédit",
        ]

    # Parser les montants (vectorisé)
    for col in amount_cols:
        col_positions = _column_positions(df, col)
        for pos in col_positions:
            parsed, col_warnings = _parse_amounts(df.iloc[:, pos], format_detected, col)
            df.iloc[:, pos] = parsed
            warnings.extend(col_warnings)
            summary["amounts_parsed"] += 1
            summary["columns_processed"] += 1
            logger.info(f"Colonne '{col}' convertie en float")

    # Parser les dates
    date_col = "date"
    date_positions = _column_positions(df, date_col)
    if date_positions:
        dayfirst = _detect_date_format(df, date_col, config.date_format_sample_size)
        for pos in date_positions:
            try:
                df.iloc[:, pos] = pd.to_datetime(df.iloc[:, pos], errors="coerce", dayfirst=dayfirst)
                summary["dates_parsed"] += 1
                summary["columns_processed"] += 1
                logger.info(f"Colonne '{date_col}' convertie en datetime (dayfirst={dayfirst})")
            except Exception as e:
                warnings.append(f"Erreur conversion '{date_col}' : {e}")
                logger.warning(warnings[-1])

    # Nettoyer les numéros de compte (si présent)
    compte_positions = _column_positions(df, "n°compte")
    for pos in compte_positions:
        df.iloc[:, pos] = (
            df.iloc[:, pos]
            .astype(str)
            .str.strip()
            .str.replace(" ", "", regex=False)
            .str.replace(".0", "", regex=False)
        )
        summary["columns_processed"] += 1

    logger.info(f"Normalisation effectuée : {summary}")
    return df, summary, warnings
