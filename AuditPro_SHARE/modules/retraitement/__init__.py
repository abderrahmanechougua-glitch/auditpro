"""
Module de retraitement comptable intelligent - Refactorisé v2.0
Production-grade, vectorisé, audit-safe, configurable.

API Publique:
    from modules.retraitement import (
        IntelligentRetraitement,
        Config,
        process_file,
        process_files,
    )

Exemple d'utilisation:
    # Simple
    result = process_file("data/gl.xlsx")
    if result["success"]:
        df = result["dataframe"]
        print(result["report_file"])
    
    # Avancé
    config = Config(tolerance=0.005, erp_profile="sap")
    processor = IntelligentRetraitement(config=config, output_dir="./reports")
    result = processor.process_file("data/bg.xlsx", doc_type_hint="BG")
"""

from .config import Config
from .processor import IntelligentRetraitement, process_file, process_files, process_gl
from .cleaner import flag_gl_lines
from .normalizer import normalize_gl_to_standard
from .validator import validate_gl_balance

__all__ = [
    "Config",
    "IntelligentRetraitement",
    "process_file",
    "process_files",
    "process_gl",
    "flag_gl_lines",
    "normalize_gl_to_standard",
    "validate_gl_balance",
]

__version__ = "2.0.0"
__author__ = "AuditPro Engine"
