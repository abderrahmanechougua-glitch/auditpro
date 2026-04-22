"""
Excel Analysis Skill Implementation for AuditPro modules.

Provides enhanced Excel capabilities:
- Data validation and anomaly detection
- Pivot table generation
- Format detection and normalization
- Audit trail and reporting
"""
from __future__ import annotations
from pathlib import Path
from typing import Optional, Dict, Any, List
import logging
import pandas as pd

logger = logging.getLogger(__name__)


class ExcelSkillHandler:
    """Handles Excel-based skill operations for audit modules."""
    
    def __init__(self):
        self.pandas_available = True  # Pandas is in requirements

    @staticmethod
    def _find_matching_column(columns, candidates: List[str]) -> str | None:
        normalized = {str(col).lower().strip(): col for col in columns}
        for candidate in candidates:
            if candidate in normalized:
                return normalized[candidate]
        for col in columns:
            col_lower = str(col).lower().strip()
            if any(candidate in col_lower for candidate in candidates):
                return col
        return None

    @staticmethod
    def _to_numeric(series: pd.Series) -> pd.Series:
        cleaned = (
            series.astype(str)
            .str.replace("\u00a0", "", regex=False)
            .str.replace(" ", "", regex=False)
            .str.replace("MAD", "", regex=False)
            .str.replace("€", "", regex=False)
            .str.replace(",", ".", regex=False)
        )
        return pd.to_numeric(cleaned, errors="coerce")
    
    def validate_data(
        self,
        df: pd.DataFrame,
        rules: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        Validate Excel data against business rules.
        
        Args:
            df: DataFrame to validate
            rules: Dict of validation rules (e.g., {"debit_sum": "equals", "credit_sum"})
        
        Returns:
            Dict with validation results
        """
        result = {
            "valid": True,
            "errors": [],
            "warnings": [],
            "stats": {}
        }
        
        try:
            # Basic data quality checks
            result["stats"]["total_rows"] = len(df)
            result["stats"]["total_cols"] = len(df.columns)
            result["stats"]["null_values"] = df.isnull().sum().to_dict()
            result["stats"]["duplicates"] = df.duplicated().sum()
            
            if result["stats"]["null_values"]:
                result["warnings"].append(
                    f"Null values detected: {result['stats']['null_values']}"
                )
            
            if result["stats"]["duplicates"] > 0:
                result["warnings"].append(
                    f"{result['stats']['duplicates']} duplicate rows detected"
                )
            
            # Apply custom validation rules
            if rules:
                for rule_name, rule_config in rules.items():
                    try:
                        if rule_config.get("type") == "formula":
                            # Example: check debit = credit balance
                            result["stats"][rule_name] = eval(
                                rule_config["formula"],
                                {"df": df}
                            )
                    except Exception as e:
                        result["errors"].append(f"Rule '{rule_name}' failed: {e}")
                        result["valid"] = False
            
        except Exception as e:
            logger.error(f"Error validating data: {e}")
            result["errors"].append(str(e))
            result["valid"] = False
        
        return result
    
    def create_pivot_table(
        self,
        df: pd.DataFrame,
        index: str | List[str],
        values: str | List[str],
        aggfunc: str = "sum"
    ) -> Dict[str, Any]:
        """
        Create pivot table for GL analysis.
        
        Args:
            df: Source DataFrame
            index: Column(s) for index
            values: Column(s) for values
            aggfunc: Aggregation function (sum, mean, count, etc.)
        
        Returns:
            Dict with pivot table and metadata
        """
        result = {
            "success": False,
            "pivot_data": None,
            "shape": None,
            "summary": {}
        }
        
        try:
            pivot = pd.pivot_table(
                df,
                index=index,
                values=values,
                aggfunc=aggfunc,
                fill_value=0
            )
            
            result["pivot_data"] = pivot.to_dict()
            result["shape"] = pivot.shape
            result["success"] = True
            
            # Add summary statistics
            result["summary"] = {
                "total": pivot.sum().sum() if isinstance(pivot.sum(), pd.Series) else pivot.sum(),
                "mean": float(pivot.mean().mean()) if hasattr(pivot, 'mean') else 0,
                "rows": len(pivot),
                "cols": len(pivot.columns)
            }
            
        except Exception as e:
            logger.error(f"Error creating pivot table: {e}")
            result["error"] = str(e)
        
        return result
    
    def detect_format(
        self,
        file_path: str
    ) -> Dict[str, Any]:
        """
        Auto-detect accounting file format (Sage, D365, AX).
        
        Args:
            file_path: Path to Excel file
        
        Returns:
            Dict with detected format and confidence
        """
        result = {
            "detected_format": None,
            "confidence": 0.0,
            "characteristics": [],
            "file_path": str(file_path)
        }
        
        try:
            # Read first sheet to analyze structure
            df = pd.read_excel(file_path, sheet_name=0, nrows=5)
            
            columns = df.columns.tolist()
            
            # Sage 100 detection (merged cells, specific columns)
            if any(col in str(columns) for col in ["C.j", "Intitulé", "Solde initial"]):
                result["detected_format"] = "Sage 100"
                result["confidence"] = 0.85
                result["characteristics"] = ["Sage-specific column naming", "Likely merged cells"]
            
            # Dynamics 365 detection
            elif any(col in str(columns) for col in ["Main Account", "Posting Layer", "Amount"]):
                result["detected_format"] = "Dynamics 365"
                result["confidence"] = 0.9
                result["characteristics"] = ["D365-specific naming convention"]
            
            # AX detection
            elif any(col in str(columns) for col in ["MainAccount", "Dimension", "AmountCurDebit"]):
                result["detected_format"] = "Dynamics AX"
                result["confidence"] = 0.9
                result["characteristics"] = ["AX naming pattern"]
            
            # Generic GL detection
            elif any(col in str(columns) for col in ["Account", "Debit", "Credit"]):
                result["detected_format"] = "Generic GL"
                result["confidence"] = 0.7
                result["characteristics"] = ["Standard GL columns"]
            
            logger.info(f"Detected format: {result['detected_format']} ({result['confidence']})")
            
        except Exception as e:
            logger.error(f"Error detecting format from {file_path}: {e}")
            result["error"] = str(e)
        
        return result
    
    def normalize_gl(
        self,
        file_path: str,
        target_format: str = "standard"
    ) -> Dict[str, Any]:
        """
        Normalize GL from various formats to standard format.
        
        Args:
            file_path: Path to GL file
            target_format: Target format (default: "standard")
        
        Returns:
            Dict with normalized data and mapping
        """
        result = {
            "success": False,
            "normalized_data": None,
            "column_mapping": {},
            "rows_processed": 0,
            "file_path": str(file_path)
        }
        
        try:
            # Detect format first
            format_result = self.detect_format(file_path)
            source_format = format_result.get("detected_format")
            
            # Read data
            df = pd.read_excel(file_path)
            
            # Define standard GL columns
            standard_cols = ["Account", "Description", "Debit", "Credit", "Balance"]
            
            # Create mapping based on detected format
            if source_format == "Sage 100":
                mapping = {
                    "C.j": "Account",
                    "Intitulé": "Description",
                    "Débit": "Debit",
                    "Crédit": "Credit"
                }
            elif source_format == "Dynamics 365":
                mapping = {
                    "Main Account": "Account",
                    "Posting Layer": "Description",
                    "Debit": "Debit",
                    "Credit": "Credit"
                }
            else:
                mapping = {}
                # Auto-map columns by similarity
                for col in df.columns:
                    col_lower = col.lower()
                    if "account" in col_lower or "compte" in col_lower:
                        mapping[col] = "Account"
                    elif "debit" in col_lower:
                        mapping[col] = "Debit"
                    elif "credit" in col_lower or "crédit" in col_lower:
                        mapping[col] = "Credit"
            
            # Apply mapping
            normalized = df.rename(columns=mapping)
            
            # Ensure numeric columns
            for col in ["Debit", "Credit", "Balance"]:
                if col in normalized.columns:
                    normalized[col] = pd.to_numeric(normalized[col], errors='coerce')
            
            # Calculate balance if not present
            if "Balance" not in normalized.columns and "Debit" in normalized.columns:
                normalized["Balance"] = normalized["Debit"] - normalized["Credit"]
            
            result["normalized_data"] = normalized.to_dict()
            result["column_mapping"] = mapping
            result["rows_processed"] = len(normalized)
            result["success"] = True
            
        except Exception as e:
            logger.error(f"Error normalizing GL from {file_path}: {e}")
            result["error"] = str(e)
        
        return result
    
    def detect_anomalies(
        self,
        df: pd.DataFrame,
        numeric_column: str = "Amount",
        threshold: float = 2.0
    ) -> Dict[str, Any]:
        """
        Detect data anomalies using statistical methods.
        
        Args:
            df: DataFrame to analyze
            numeric_column: Column to analyze for anomalies
            threshold: Standard deviations for outlier detection
        
        Returns:
            Dict with anomalies and statistics
        """
        result = {
            "anomalies_found": 0,
            "outliers": [],
            "statistics": {}
        }
        
        try:
            if numeric_column not in df.columns:
                result["error"] = f"Column '{numeric_column}' not found"
                return result
            
            # Convert to numeric
            col_data = pd.to_numeric(df[numeric_column], errors='coerce')
            
            # Calculate statistics
            mean = col_data.mean()
            std = col_data.std()
            
            result["statistics"] = {
                "mean": float(mean),
                "std": float(std),
                "min": float(col_data.min()),
                "max": float(col_data.max()),
                "median": float(col_data.median())
            }
            
            # Detect outliers
            outlier_mask = (col_data > mean + threshold * std) | (col_data < mean - threshold * std)
            
            if outlier_mask.any():
                outliers = df[outlier_mask].copy()
                result["outliers"] = outliers.to_dict(orient='records')
                result["anomalies_found"] = len(outliers)
            
        except Exception as e:
            logger.error(f"Error detecting anomalies: {e}")
            result["error"] = str(e)
        
        return result

    def analyze_accounting_data(
        self,
        df: pd.DataFrame,
        key_columns: Optional[Dict[str, str]] = None
    ) -> Dict[str, Any]:
        """Run accounting-oriented audit checks on a ledger-like dataset."""
        result = {
            "success": False,
            "findings": [],
            "warnings": [],
            "stats": {
                "total_rows": len(df),
                "total_columns": len(df.columns),
                "duplicate_rows": 0,
                "duplicate_entries": 0,
                "missing_key_values": 0,
                "debit_total": 0.0,
                "credit_total": 0.0,
                "balance_gap": 0.0,
                "outlier_rows": 0,
            },
            "resolved_columns": {}
        }

        try:
            resolved = dict(key_columns or {})
            if not resolved.get("account"):
                resolved["account"] = self._find_matching_column(
                    df.columns, ["account", "compte", "n°compte", "numero compte"]
                )
            if not resolved.get("debit"):
                resolved["debit"] = self._find_matching_column(
                    df.columns, ["debit", "débit", "solde_débit", "mvt_débit"]
                )
            if not resolved.get("credit"):
                resolved["credit"] = self._find_matching_column(
                    df.columns, ["credit", "crédit", "solde_crédit", "mvt_crédit"]
                )
            if not resolved.get("date"):
                resolved["date"] = self._find_matching_column(
                    df.columns, ["date", "date piece", "date ecriture", "date écriture"]
                )
            if not resolved.get("piece"):
                resolved["piece"] = self._find_matching_column(
                    df.columns, ["piece", "pièce", "reference", "référence", "ref"]
                )

            result["resolved_columns"] = {k: v for k, v in resolved.items() if v}

            duplicate_rows = int(df.duplicated().sum())
            result["stats"]["duplicate_rows"] = duplicate_rows
            if duplicate_rows:
                result["findings"].append(f"{duplicate_rows} ligne(s) dupliquée(s) exactes détectées.")

            subset = [resolved.get(name) for name in ("account", "date", "piece", "debit", "credit") if resolved.get(name) in df.columns]
            if len(subset) >= 2:
                duplicate_entries = int(df.duplicated(subset=subset).sum())
                result["stats"]["duplicate_entries"] = duplicate_entries
                if duplicate_entries:
                    result["warnings"].append(f"{duplicate_entries} écriture(s) potentiellement dupliquée(s) sur les colonnes clés.")

            key_value_cols = [resolved.get(name) for name in ("account", "debit", "credit") if resolved.get(name) in df.columns]
            if key_value_cols:
                missing_key_values = int(df[key_value_cols].isna().sum().sum())
                result["stats"]["missing_key_values"] = missing_key_values
                if missing_key_values:
                    result["findings"].append(f"{missing_key_values} valeur(s) manquante(s) sur les colonnes clés.")

            debit_col = resolved.get("debit")
            credit_col = resolved.get("credit")
            if debit_col in df.columns and credit_col in df.columns:
                debit_values = self._to_numeric(df[debit_col]).fillna(0)
                credit_values = self._to_numeric(df[credit_col]).fillna(0)
                debit_total = float(debit_values.sum())
                credit_total = float(credit_values.sum())
                balance_gap = round(abs(debit_total - credit_total), 2)
                result["stats"]["debit_total"] = debit_total
                result["stats"]["credit_total"] = credit_total
                result["stats"]["balance_gap"] = balance_gap
                if balance_gap > 0.01:
                    result["findings"].append(
                        f"Déséquilibre débit/crédit détecté : écart de {balance_gap:,.2f}."
                    )

                movement_mask = (debit_values.abs() + credit_values.abs()) > 0
                movement_values = (debit_values.abs() + credit_values.abs())[movement_mask]
                if not movement_values.empty:
                    q1 = movement_values.quantile(0.25)
                    q3 = movement_values.quantile(0.75)
                    iqr = q3 - q1
                    if iqr > 0:
                        upper_bound = q3 + 1.5 * iqr
                        outlier_rows = int((movement_values > upper_bound).sum())
                        result["stats"]["outlier_rows"] = outlier_rows
                        if outlier_rows:
                            result["warnings"].append(f"{outlier_rows} écriture(s) à montant atypique détectée(s).")

            date_col = resolved.get("date")
            if date_col in df.columns:
                dates = pd.to_datetime(df[date_col], errors="coerce")
                valid_dates = dates.dropna()
                if not valid_dates.empty:
                    result["stats"]["date_min"] = valid_dates.min().strftime("%Y-%m-%d")
                    result["stats"]["date_max"] = valid_dates.max().strftime("%Y-%m-%d")
                    invalid_dates = int(dates.isna().sum())
                    if invalid_dates:
                        result["warnings"].append(f"{invalid_dates} date(s) invalide(s) ou non interprétable(s).")

            account_col = resolved.get("account")
            if account_col in df.columns:
                account_values = df[account_col].astype(str).str.strip()
                invalid_accounts = int((account_values == "").sum())
                if invalid_accounts:
                    result["findings"].append(f"{invalid_accounts} numéro(s) de compte vide(s) détecté(s).")

            result["success"] = True

        except Exception as e:
            logger.error(f"Error analyzing accounting data: {e}")
            result["findings"].append(str(e))

        return result

    def analyze_srm_data(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Run SRM-oriented checks on a summary review memorandum source table."""
        result = {
            "success": False,
            "findings": [],
            "warnings": [],
            "stats": {
                "total_rows": len(df),
                "total_columns": len(df.columns),
                "duplicate_labels": 0,
                "missing_labels": 0,
                "missing_numeric_values": 0,
                "variation_mismatches": 0,
                "outlier_variations": 0,
            },
            "resolved_columns": {},
        }

        try:
            label_col = self._find_matching_column(
                df.columns,
                ["label", "libellé", "libelle", "poste", "rubrique", "description", "intitulé", "intitule"],
            )
            current_col = self._find_matching_column(
                df.columns,
                ["n", "montant n", "année n", "annee n", "valeur n"],
            )
            prior_col = self._find_matching_column(
                df.columns,
                ["n-1", "n_1", "n 1", "montant n-1", "année n-1", "annee n-1"],
            )
            variation_col = self._find_matching_column(
                df.columns,
                ["variation", "% variation", "ecart", "écart"],
            )

            result["resolved_columns"] = {
                key: value
                for key, value in {
                    "label": label_col,
                    "n": current_col,
                    "n_1": prior_col,
                    "variation": variation_col,
                }.items()
                if value
            }

            if label_col and label_col in df.columns:
                labels = df[label_col].astype(str).str.strip()
                meaningful_labels = labels[(labels != "") & (~labels.str.lower().isin(["nan", "none"]))]
                missing_labels = int(len(labels) - len(meaningful_labels))
                duplicate_labels = int(meaningful_labels.duplicated().sum())
                result["stats"]["missing_labels"] = missing_labels
                result["stats"]["duplicate_labels"] = duplicate_labels
                if missing_labels:
                    result["findings"].append(f"{missing_labels} libellé(s) SRM manquant(s) ou vides détectés.")
                if duplicate_labels:
                    result["warnings"].append(f"{duplicate_labels} libellé(s) SRM dupliqué(s) détectés.")

            numeric_columns = [col for col in [current_col, prior_col, variation_col] if col in df.columns]
            numeric_data = {}
            if numeric_columns:
                for col in numeric_columns:
                    numeric_data[col] = self._to_numeric(df[col])
                missing_numeric_values = int(pd.DataFrame(numeric_data).isna().sum().sum())
                result["stats"]["missing_numeric_values"] = missing_numeric_values
                if missing_numeric_values:
                    result["warnings"].append(f"{missing_numeric_values} cellule(s) numérique(s) SRM manquante(s) ou non lisible(s).")

            if current_col in numeric_data and prior_col in numeric_data:
                current_values = numeric_data[current_col].fillna(0)
                prior_values = numeric_data[prior_col].fillna(0)
                delta = current_values - prior_values

                if variation_col in numeric_data:
                    variation_values = numeric_data[variation_col]
                    mismatch_mask = variation_values.notna() & (delta - variation_values).abs().gt(1)
                    mismatches = int(mismatch_mask.sum())
                    result["stats"]["variation_mismatches"] = mismatches
                    if mismatches:
                        result["findings"].append(
                            f"{mismatches} ligne(s) avec variation incohérente entre N, N-1 et la colonne variation."
                        )

                non_zero_delta = delta[delta != 0]
                if not non_zero_delta.empty:
                    q1 = non_zero_delta.quantile(0.25)
                    q3 = non_zero_delta.quantile(0.75)
                    iqr = q3 - q1
                    if iqr > 0:
                        lower_bound = q1 - 1.5 * iqr
                        upper_bound = q3 + 1.5 * iqr
                        outlier_variations = int(((non_zero_delta < lower_bound) | (non_zero_delta > upper_bound)).sum())
                        result["stats"]["outlier_variations"] = outlier_variations
                        if outlier_variations:
                            result["warnings"].append(
                                f"{outlier_variations} variation(s) atypique(s) détectée(s) dans le tableau SRM."
                            )

            result["success"] = True

        except Exception as e:
            logger.error(f"Error analyzing SRM data: {e}")
            result["findings"].append(str(e))

        return result


# Singleton instance
_excel_handler = ExcelSkillHandler()


def get_excel_skill_handler() -> ExcelSkillHandler:
    """Get the Excel skill handler instance."""
    return _excel_handler
