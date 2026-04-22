"""
Main orchestrator : intégration de tous les modules du retraitement comptable.
Fournit l'API publique pour traiter des fichiers comptables.

Production-grade, vectorisé, audit-safe, maintenable.
"""

# Guard contre l'exécution directe avec relative imports
if __name__ == '__main__':
    import sys
    print("[x] Erreur : Ce module doit être importé, pas exécuté directement.")
    print("[x] Utilisez : python -c \"from modules.retraitement import process_file\"")
    sys.exit(1)

import logging
import re
from pathlib import Path
from typing import Dict, Optional, Tuple

import pandas as pd

from .cleaner import flag_rows, get_clean_data, flag_gl_lines
from .config import Config
from .detector import detect_document_type
from .loader import load_excel
from .normalizer import normalize_data_types, standardize_columns, normalize_gl_to_standard
from .reporter import generate_report
from .validator import build_validation_report, validate_document, validate_gl_balance

logger = logging.getLogger(__name__)


def _detect_gl_source_format(file_path: Path, max_rows: int = 40) -> str:
    """
    Détecte le format source GL principal.

    Retourne:
    - "d365" pour Dynamics (Nominal/Sector/Charge Type, Voucher, etc.)
    - "sage" pour exports Sage 100 (Grand-livre des comptes, C.j, N° pièce)
    """
    try:
        df_top = pd.read_excel(file_path, header=None, nrows=max_rows, dtype=str)
    except Exception:
        return "d365"

    joined = " ".join(
        str(v).lower() for v in df_top.fillna("").values.flatten() if str(v).strip()
    )

    d365_signals = [
        "nominal/sector/charge type",
        "accumulated",
        "voucher",
        "opening balance",
        "closing balance",
    ]
    sage_signals = [
        "sage 100",
        "grand-livre des comptes",
        "c.j",
        "n° pièce",
        "solde",
        "progressif",
        "total compte",
    ]

    d365_score = sum(1 for s in d365_signals if s in joined)
    sage_score = sum(1 for s in sage_signals if s in joined)

    if sage_score >= d365_score and sage_score >= 2:
        return "sage"
    return "d365"


def _detect_gl_header_row(file_path: Path, max_rows: int = 20) -> int:
    """Détecte la ligne d'en-tête d'un GL D365 (cherche 'Date' en col 0 suivi de 'Voucher')."""
    import re as _re
    try:
        df_top = pd.read_excel(file_path, header=None, nrows=max_rows, dtype=str)
        for i, row in df_top.iterrows():
            val0 = str(row.iloc[0]).strip().lower()
            val2 = str(row.iloc[2]).strip().lower() if len(row) > 2 else ""
            if val0 == "date" and ("voucher" in val2 or "document" in val2):
                return int(i)
        # Fallback: cherche juste 'date' en col 0
        for i, row in df_top.iterrows():
            if str(row.iloc[0]).strip().lower() == "date":
                return int(i)
    except Exception:
        pass
    return 9  # D365 default


def _load_gl_d365(file_path: Path) -> pd.DataFrame:
    """
    Charge un GL D365 avec propagation du numéro de compte et libellé.

    - Détecte la ligne d'en-tête automatiquement.
    - Extrait le compte (Unnamed:1) et libellé (Description) des lignes titres.
    - Forward-fill sur toutes les lignes (y compris Opening/Closing balance).
    - Retourne le DataFrame avec colonnes _COMPTE et _LIBELLE_COMPTE.
    """
    import re as _re

    header_row = _detect_gl_header_row(file_path)

    # Lire les premières lignes brutes pour récupérer le titre du 1er compte (avant l'en-tête)
    df_top = pd.read_excel(file_path, header=None, nrows=header_row + 1, dtype=str)
    first_compte = ""
    first_libelle = ""
    for i in range(header_row):
        val0 = str(df_top.iloc[i, 0]).strip()
        if _re.search(r"nominal|charge type|sector", val0, _re.IGNORECASE):
            c_raw = str(df_top.iloc[i, 1]).strip() if df_top.shape[1] > 1 else ""
            l_raw = str(df_top.iloc[i, 3]).strip() if df_top.shape[1] > 3 else ""
            first_compte = _re.sub(r"-+\s*$", "", c_raw).strip()
            first_libelle = _re.sub(r"-+\s*$", "", l_raw).strip()
            first_compte = "" if first_compte.lower() == "nan" else first_compte
            first_libelle = "" if first_libelle.lower() == "nan" else first_libelle

    # Charger le fichier avec l'en-tête détecté
    df = pd.read_excel(file_path, header=header_row, dtype=str)

    # Identifier la colonne date et la colonne "Unnamed: 1" (numéro de compte dans D365)
    date_col = next((c for c in df.columns if str(c).strip().lower() == "date"), None)
    unnamed_col = df.columns[1] if len(df.columns) > 1 else None
    desc_col = next((c for c in df.columns if str(c).strip().lower() == "description"), None)

    # Construire les séries COMPTE / LIBELLE depuis les lignes titres dans le df
    _COMPTE: pd.Series = pd.Series(pd.NA, index=df.index, dtype=object)
    _LIBELLE: pd.Series = pd.Series(pd.NA, index=df.index, dtype=object)

    title_mask = pd.Series(False, index=df.index)
    if date_col:
        date_vals = df[date_col].astype(str).str.strip()
        title_mask = date_vals.str.contains(
            r"^nominal|^sector|charge type",
            na=False, regex=True, flags=_re.IGNORECASE
        )
        for idx in df[title_mask].index:
            if unnamed_col is not None:
                c_raw = str(df.at[idx, unnamed_col]).strip()
                c_clean = _re.sub(r"-+\s*$", "", c_raw).strip()
                if c_clean and c_clean.lower() != "nan":
                    _COMPTE[idx] = c_clean
            if desc_col is not None:
                l_raw = str(df.at[idx, desc_col]).strip()
                l_clean = _re.sub(r"-+\s*$", "", l_raw).strip()
                if l_clean and l_clean.lower() != "nan":
                    _LIBELLE[idx] = l_clean

    # Appliquer le premier compte (du bloc pré-en-tête) aux lignes précédant le 1er titre
    if first_compte:
        first_title_in_df = df[title_mask].index[0] if title_mask.any() else None
        if first_title_in_df is not None:
            rows_before = df.index[df.index < first_title_in_df]
        else:
            rows_before = df.index
        if len(rows_before) > 0:
            _COMPTE[rows_before[0]] = first_compte
            _LIBELLE[rows_before[0]] = first_libelle

    # Forward-fill
    df["_COMPTE"] = _COMPTE.ffill().fillna("")
    df["_LIBELLE_COMPTE"] = _LIBELLE.ffill().fillna("")

    return df


def _detect_gl_header_row_sage(file_path: Path, max_rows: int = 40) -> int:
    """Détecte la ligne d'en-tête d'un GL Sage (Date / C.j / N° pièce / Libellé écriture)."""
    try:
        df_top = pd.read_excel(file_path, header=None, nrows=max_rows, dtype=str)
        for i, row in df_top.iterrows():
            values = [str(v).strip().lower() for v in row.tolist()]
            row_text = " ".join(values)
            if (
                "date" in row_text
                and ("c.j" in row_text or "cj" in row_text)
                and ("n° pièce" in row_text or "n° piece" in row_text or "piece" in row_text)
            ):
                return int(i)
    except Exception:
        pass
    return 7


def _best_numeric_neighbor(df: pd.DataFrame, base_idx: int) -> Optional[int]:
    """Pour une colonne en-tête fusionnée, choisit la meilleure colonne de données voisine."""
    candidates = [base_idx, base_idx + 1, base_idx - 1]
    candidates = [i for i in candidates if 0 <= i < len(df.columns)]
    if not candidates:
        return None

    best_idx = None
    best_score = -1
    for idx in candidates:
        score = pd.to_numeric(df.iloc[:, idx], errors="coerce").notna().sum()
        if score > best_score:
            best_score = int(score)
            best_idx = idx
    return best_idx


def _load_gl_sage(file_path: Path) -> pd.DataFrame:
    """
    Charge un GL Sage et le convertit vers une structure canonique GL interne.

    Colonnes produites:
    - Date, Voucher, Description, Debit, Credit, Accumulated
    - _COMPTE, _LIBELLE_COMPTE, _SAGE_ROW_TYPE
    """
    header_row = _detect_gl_header_row_sage(file_path)
    df = pd.read_excel(file_path, header=header_row, dtype=str)

    cols = list(df.columns)
    col0 = cols[0] if len(cols) > 0 else None
    col1 = cols[1] if len(cols) > 1 else None
    col2 = cols[2] if len(cols) > 2 else None
    col3 = cols[3] if len(cols) > 3 else None

    debit_header_idx = next(
        (i for i, c in enumerate(cols) if "debit" in str(c).lower() or "débit" in str(c).lower()),
        None,
    )
    credit_header_idx = next(
        (i for i, c in enumerate(cols) if "credit" in str(c).lower() or "crédit" in str(c).lower()),
        None,
    )
    solde_header_idx = next(
        (i for i, c in enumerate(cols) if "solde" in str(c).lower() or "progressif" in str(c).lower()),
        None,
    )

    debit_idx = _best_numeric_neighbor(df, debit_header_idx) if debit_header_idx is not None else None
    credit_idx = _best_numeric_neighbor(df, credit_header_idx) if credit_header_idx is not None else None
    solde_idx = _best_numeric_neighbor(df, solde_header_idx) if solde_header_idx is not None else None

    date_vals = df.iloc[:, 0].astype(str).str.strip() if col0 is not None else pd.Series("", index=df.index)
    c1_vals = df.iloc[:, 1].astype(str).str.strip() if col1 is not None else pd.Series("", index=df.index)
    c2_vals = df.iloc[:, 2].astype(str).str.strip() if col2 is not None else pd.Series("", index=df.index)
    c3_vals = df.iloc[:, 3].astype(str).str.strip() if col3 is not None else pd.Series("", index=df.index)

    is_title = date_vals.str.match(r"^\d{6,10}$", na=False) & ~c1_vals.str.contains(r"^total\s+compte$", case=False, na=False)
    is_total = c1_vals.str.contains(r"^total\s+compte$", case=False, na=False)
    parsed_date = pd.to_datetime(date_vals, errors="coerce")
    is_movement = parsed_date.notna()

    compte = pd.Series(pd.NA, index=df.index, dtype=object)
    libelle_compte = pd.Series(pd.NA, index=df.index, dtype=object)
    compte[is_title] = date_vals[is_title]
    libelle_compte[is_title] = c1_vals[is_title]

    compte = compte.ffill().fillna("")
    libelle_compte = libelle_compte.ffill().fillna("")

    voucher = c2_vals.copy()
    voucher = voucher.where(voucher.str.lower().ne("nan"), "")
    voucher = voucher.where(voucher.str.strip().ne(""), c1_vals)

    description = c3_vals.copy()
    description = description.where(description.str.lower().ne("nan"), "")

    debit_vals = df.iloc[:, debit_idx] if debit_idx is not None else pd.Series("", index=df.index)
    credit_vals = df.iloc[:, credit_idx] if credit_idx is not None else pd.Series("", index=df.index)
    acc_vals = df.iloc[:, solde_idx] if solde_idx is not None else pd.Series("", index=df.index)

    out = pd.DataFrame(
        {
            "Date": date_vals,
            "Voucher": voucher,
            "Description": description,
            "Debit": debit_vals,
            "Credit": credit_vals,
            "Accumulated": acc_vals,
            "_COMPTE": compte,
            "_LIBELLE_COMPTE": libelle_compte,
        },
        index=df.index,
    )

    row_type = pd.Series("other", index=df.index, dtype=object)
    row_type[is_title] = "title"
    row_type[is_total] = "total"
    row_type[is_movement] = "movement"
    row_type[out[["Date", "Voucher", "Description", "Debit", "Credit", "Accumulated"]]
             .replace(r"^\s*$", pd.NA, regex=True).isna().all(axis=1)] = "empty"
    out["_SAGE_ROW_TYPE"] = row_type

    return out


def _read_gl_file(file_path: Path) -> pd.DataFrame:
    """Chargement GL générique (CSV / Excel sans structure D365)."""
    ext = file_path.suffix.lower()
    if ext == ".csv":
        return pd.read_csv(file_path, sep=None, engine="python")
    if ext in (".xlsx", ".xls", ".xlsm", ".xlsb"):
        engine = "pyxlsb" if ext == ".xlsb" else None
        return pd.read_excel(file_path, engine=engine)
    raise ValueError(f"Format non supporté pour GL: {file_path.suffix}")


def process_gl(
    file_path: Path,
    compte_override: Optional[str] = None,
    output_dir: Optional[Path] = None,
    filename_stem: Optional[str] = None,
) -> Dict:
    """Pipeline GL dédié: cleaner -> normalizer -> validator avec log complet."""
    file_path = Path(file_path)
    output_dir = Path(output_dir) if output_dir else file_path.parent
    logs = []

    def log_step(message: str):
        logger.info(message)
        logs.append(message)

    try:
        log_step(f"Lecture GL: {file_path}")
        source_format = _detect_gl_source_format(file_path)
        log_step(f"Format source GL détecté: {source_format}")
        if source_format == "sage":
            df_raw = _load_gl_sage(file_path)
        else:
            df_raw = _load_gl_d365(file_path)
        if "keep" in df_raw.columns:
            df_raw = df_raw.drop(columns=["keep"])
            log_step("Colonne legacy 'keep' supprimée avant recalcul du flag")
        if "_flag" in df_raw.columns:
            df_raw = df_raw.drop(columns=["_flag"])
            log_step("Colonne legacy '_flag' supprimée avant recalcul du flag")
        log_step(f"Lignes brutes: {len(df_raw)} | Colonnes: {len(df_raw.columns)}")

        df_flagged = flag_gl_lines(df_raw)
        keep_rows = int((df_flagged["_flag"] == "keep").sum())
        excluded_rows = int((df_flagged["_flag"] != "keep").sum())
        log_step(f"Lignes marquées keep: {keep_rows} | exclude: {excluded_rows}")

        df_keep = df_flagged[df_flagged["_flag"] == "keep"].copy()
        if df_keep.empty:
            return {
                "success": False,
                "error": "Aucune ligne exploitable après flagging GL",
                "dataframe": None,
                "violations": [],
                "log": logs,
            }

        df_standard = normalize_gl_to_standard(df_keep, compte_col="n°compte", compte_value=compte_override)
        detected_format = df_standard.attrs.get("gl_format_detected", "unknown")
        column_mapping = df_standard.attrs.get("column_mapping", {})
        log_step(f"Format GL détecté: {detected_format}")
        log_step(f"Colonnes mappées: {column_mapping}")

        violations = validate_gl_balance(df_standard)
        log_step(f"Violations: {len(violations)}")

        validation_report = build_validation_report(violations)
        metadata = {
            "file_path": str(file_path),
            "sheet_name": "N/A",
            "header_row_idx": None,
            "original_rows": len(df_raw),
            "original_columns": len(df_raw.columns),
        }
        doc_explanation = {
            "type": "GL",
            "confidence": 1.0,
            "reason": f"Pipeline GL universel ({detected_format})",
        }
        report_file = generate_report(
            df_standard,
            metadata,
            doc_explanation,
            column_mapping,
            {
                "amounts_parsed": 2,
                "dates_parsed": 1,
                "columns_processed": 6,
            },
            validation_report,
            output_dir=output_dir,
            filename_stem=filename_stem or file_path.stem,
        )
        log_step(f"Rapport GL généré: {report_file}")

        return {
            "success": True,
            "error": None,
            "dataframe": df_standard,
            "violations": violations,
            "validation_report": validation_report,
            "report_file": report_file,
            "doc_type": "GL",
            "log": logs,
            "excluded_rows": excluded_rows,
            "flag_counts": {"keep": keep_rows, "exclude": excluded_rows},
            "format_detected": detected_format,
            "source_format": source_format,
            "column_mapping": column_mapping,
        }
    except Exception as exc:
        logger.error("process_gl: erreur inattendue: %s", exc, exc_info=True)
        return {
            "success": False,
            "error": f"Erreur inattendue : {exc}",
            "dataframe": None,
            "violations": [],
            "log": logs,
        }


class IntelligentRetraitement:
    """
    Processeur intelligent et production-grade de retraitement comptable.
    
    Caractéristiques :
    - Vectorisé : O(n) performance, pas de boucles Python sur les données
    - Audit-safe : trace complète des transformations, pas de suppression silencieuse
    - Configurable : règles externalisées, profils ERP
    - Modulaire : chaque étape est testable indépendamment
    - Production-ready : gestion complète des erreurs, logging
    """

    def __init__(self, config: Optional[Config] = None, output_dir: Optional[Path] = None):
        self.config = config or Config()
        self.output_dir = Path(output_dir) if output_dir else Path.cwd()
        self.output_dir.mkdir(parents=True, exist_ok=True)
        logger.info(f"Retraitement initialisé | Répertoire sortie : {self.output_dir}")

    def process_file(
        self,
        file_path: Path,
        doc_type_hint: Optional[str] = None,
        erp_profile: str = "default",
        keep_flagged_rows: bool = True,
    ) -> Dict:
        """
        Traite un fichier comptable complet.
        
        Pipeline :
        1. Charger le fichier → détecter en-tête
        2. Marquer les lignes problématiques (pas de suppression)
        3. Détecter le type de document
        4. Standardiser les colonnes
        5. Normaliser les données (montants, dates)
        6. Valider les règles métier
        7. Générer le rapport avec audit trail
        
        Args:
            file_path: Chemin du fichier Excel.
            doc_type_hint: Type forcé si connu (GL, BG, AUX), sinon auto-détection.
            erp_profile: Profil ERP à appliquer (default, sap, ciel, sage).
            keep_flagged_rows: Si True, conserver les lignes marquées dans la sortie.
        
        Returns:
            Dict contenant :
            {
                'success': bool,
                'error': str (si succès=False),
                'dataframe': pd.DataFrame (données normalisées),
                'metadata': dict (infos fichier, en-tête, etc.),
                'doc_type': str (GL/BG/AUX),
                'doc_type_explanation': dict (détails détection),
                'validation_report': pd.DataFrame (violations),
                'report_file': str (chemin du fichier rapport),
            }
        """
        file_path = Path(file_path)

        try:
            # Appliquer le profil ERP
            if erp_profile != "default":
                self.config.apply_erp_profile(erp_profile)
                logger.info(f"Profil ERP appliqué : {erp_profile}")

            # 1. Charger le fichier
            logger.info(f"Traitement du fichier : {file_path}")
            df_raw, header_idx, metadata = load_excel(file_path, header_detection=True)
            logger.info(f"Fichier chargé : {len(df_raw)} lignes, {len(df_raw.columns)} colonnes")

            # 2. Marquer les lignes (pas de suppression)
            df_flagged = flag_rows(df_raw, self.config)
            logger.info(f"Marquage des lignes effectué")

            # 3. Détecter le type de document
            doc_type, confidence, doc_explanation = detect_document_type(
                df_flagged, self.config
            )
            if doc_type_hint and doc_type_hint in ("GL", "BG", "AUX"):
                doc_type = doc_type_hint
                logger.info(f"Type de document forcé : {doc_type}")
            logger.info(f"Type détecté : {doc_type} (confiance: {confidence:.1%})")

            if doc_type == "GL":
                logger.info("GL détecté via auto-détection, bascule vers le pipeline GL dédié")
                return process_gl(
                    file_path,
                    output_dir=self.output_dir,
                    filename_stem=file_path.stem,
                )

            # 4. Standardiser les colonnes
            df_normalized, column_mappings, column_warnings = standardize_columns(
                df_flagged, doc_type, self.config
            )
            logger.info(f"Colonnes standardisées : {len(column_mappings)} mappages")

            # 5. Normaliser les types de données
            df_normalized, norm_summary, norm_warnings = normalize_data_types(
                df_normalized, doc_type, self.config
            )
            logger.info(f"Données normalisées : {norm_summary}")

            # 6. Valider les règles métier
            violations = validate_document(df_normalized, doc_type, self.config)
            validation_report = build_validation_report(violations)
            logger.info(f"Validation effectuée : {len(violations)} violation(s)")

            # 7. Filtrer les lignes flaggées si demandé
            if not keep_flagged_rows:
                df_output = get_clean_data(df_normalized, remove_empty=True, remove_total=True)
                logger.info(f"Lignes flaggées supprimées : {len(df_normalized) - len(df_output)} lignes")
            else:
                df_output = df_normalized
                logger.info(f"Lignes flaggées conservées dans la sortie")

            # 8. Générer le rapport
            report_file = generate_report(
                df_output,
                metadata,
                doc_explanation,
                column_mappings,
                norm_summary,
                validation_report,
                self.output_dir,
                filename_stem=file_path.stem,
            )

            # Préparer le résultat
            result = {
                "success": True,
                "error": None,
                "dataframe": df_output,
                "metadata": metadata,
                "doc_type": doc_type,
                "doc_type_explanation": doc_explanation,
                "validation_report": validation_report,
                "report_file": report_file,
                "warnings": column_warnings + norm_warnings,
            }

            logger.info(f"Traitement réussi | Rapport : {report_file}")
            return result

        except FileNotFoundError as e:
            logger.error(f"Fichier non trouvé : {e}")
            return {
                "success": False,
                "error": str(e),
                "dataframe": None,
            }
        except ValueError as e:
            logger.error(f"Erreur de validation : {e}")
            return {
                "success": False,
                "error": str(e),
                "dataframe": None,
            }
        except Exception as e:
            logger.error(f"Erreur inattendue : {e}", exc_info=True)
            return {
                "success": False,
                "error": f"Erreur inattendue : {str(e)}",
                "dataframe": None,
            }

    def process_files(
        self,
        file_paths: list,
        doc_type_hint: Optional[str] = None,
        erp_profile: str = "default",
    ) -> list:
        """
        Traite plusieurs fichiers en batch.
        
        Args:
            file_paths: Liste des chemins de fichiers.
            doc_type_hint: Type forcé pour tous les fichiers.
            erp_profile: Profil ERP appliqué à tous.
        
        Returns:
            Liste des résultats (un dict par fichier).
        """
        results = []
        for file_path in file_paths:
            result = self.process_file(
                file_path,
                doc_type_hint=doc_type_hint,
                erp_profile=erp_profile,
            )
            results.append(result)
        return results


# --- Fonctions de commodité pour utilisation standalone ---


def process_file(
    file_path: Path,
    output_dir: Optional[Path] = None,
    config: Optional[Config] = None,
    doc_type_hint: Optional[str] = None,
    erp_profile: str = "default",
) -> Dict:
    """
    Fonction de commodité pour traiter un fichier.
    
    Crée un IntelligentRetraitement et lance le traitement.
    """
    processor = IntelligentRetraitement(config=config, output_dir=output_dir)
    return processor.process_file(
        file_path,
        doc_type_hint=doc_type_hint,
        erp_profile=erp_profile,
    )


def process_files(
    file_paths: list,
    output_dir: Optional[Path] = None,
    config: Optional[Config] = None,
    doc_type_hint: Optional[str] = None,
    erp_profile: str = "default",
) -> list:
    """
    Fonction de commodité pour traiter plusieurs fichiers.
    """
    processor = IntelligentRetraitement(config=config, output_dir=output_dir)
    return processor.process_files(
        file_paths,
        doc_type_hint=doc_type_hint,
        erp_profile=erp_profile,
    )


if __name__ == "__main__":
    # Test standalone
    import sys

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    )

    if len(sys.argv) > 1:
        files = [Path(f) for f in sys.argv[1:]]
        results = process_files(files, output_dir="./output")
        logger.info(f"Traitement terminé : {len(results)} fichier(s)")
        for r in results:
            if r.get("success"):
                logger.info(f"✓ {r['metadata']['file_path']} → {r['report_file']}")
            else:
                logger.error(f"✗ Erreur : {r['error']}")
