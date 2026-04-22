"""
XLSX Skill Implementation for AuditPro modules.

Provides spreadsheet-focused capabilities:
- workbook profiling (sheets, rows, columns)
- sheet normalization (cleanup + standardized headers)
- lightweight XLSX export helpers
"""
from __future__ import annotations

from pathlib import Path
from typing import Dict, Any, Optional
import logging

import pandas as pd


logger = logging.getLogger(__name__)


class XLSXSkillHandler:
    """Handles XLSX-specific processing for AuditPro workflows."""

    def profile_workbook(self, file_path: str) -> Dict[str, Any]:
        """Return structure and quality summary for an XLSX workbook."""
        result = {
            "success": False,
            "file_path": str(file_path),
            "sheet_count": 0,
            "total_rows": 0,
            "total_columns": 0,
            "sheets": [],
            "warnings": [],
        }

        try:
            workbook = pd.ExcelFile(file_path)
            result["sheet_count"] = len(workbook.sheet_names)

            for sheet in workbook.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet)
                rows = len(df)
                cols = len(df.columns)
                null_cells = int(df.isna().sum().sum())
                duplicate_rows = int(df.duplicated().sum())

                result["sheets"].append(
                    {
                        "name": sheet,
                        "rows": rows,
                        "columns": cols,
                        "null_cells": null_cells,
                        "duplicate_rows": duplicate_rows,
                        "columns_preview": [str(c) for c in df.columns[:15]],
                    }
                )

                result["total_rows"] += rows
                result["total_columns"] = max(result["total_columns"], cols)

                if duplicate_rows:
                    result["warnings"].append(
                        f"Feuille '{sheet}': {duplicate_rows} ligne(s) dupliquée(s)."
                    )

            result["success"] = True

        except Exception as e:
            logger.error(f"Error profiling workbook {file_path}: {e}")
            result["error"] = str(e)

        return result

    def normalize_sheet(
        self,
        file_path: str,
        output_path: str,
        sheet_name: Optional[str] = None,
        header_row: int = 0,
    ) -> Dict[str, Any]:
        """Normalize one sheet and save to a new XLSX file."""
        result = {
            "success": False,
            "input_file": str(file_path),
            "output_file": str(output_path),
            "sheet_name": sheet_name,
        }

        try:
            sheet = sheet_name or 0
            df = pd.read_excel(file_path, sheet_name=sheet, header=header_row)

            before_rows = len(df)
            before_cols = len(df.columns)

            df = df.dropna(how="all").dropna(axis=1, how="all")
            df.columns = [str(c).strip() for c in df.columns]
            df = df.reset_index(drop=True)

            out = Path(output_path)
            out.parent.mkdir(parents=True, exist_ok=True)
            df.to_excel(out, index=False)

            result["success"] = True
            result["rows_before"] = before_rows
            result["rows_after"] = len(df)
            result["cols_before"] = before_cols
            result["cols_after"] = len(df.columns)
            result["rows_removed"] = max(before_rows - len(df), 0)

        except Exception as e:
            logger.error(f"Error normalizing sheet from {file_path}: {e}")
            result["error"] = str(e)

        return result


_xlsx_handler = XLSXSkillHandler()


def get_xlsx_skill_handler() -> XLSXSkillHandler:
    """Get the XLSX skill handler instance."""
    return _xlsx_handler