"""
Module : SRM Generator
Wrapper pour srmgenV7.py — Génération du Summary Review Memorandum
"""
import os
import sys
import importlib.util
import io
from contextlib import redirect_stdout
from pathlib import Path

import pandas as pd

from agent.skills_bridge import get_skills_bridge
from modules.base_module import BaseModule, ModuleInput, ModuleResult


class SRMGenerator(BaseModule):

    name = "SRM Generator"
    description = "Génère le SRM (Word) à partir d'un tableau Excel de synthèse d'audit."
    category = "Audit"
    version = "1.0"
    help_text = (
        "ENTRÉE : Tableau Excel avec colonnes N, N-1, variation.\n"
        "Les blocs sont détectés automatiquement (CPC, Bilan, Détails).\n\n"
        "SORTIE : Document Word avec tableaux natifs + commentaires variés\n"
        "+ guide de remplacement par images Excel en annexe."
    )
    detection_keywords = ["milliers", "en kmad", "résultat d'exploitation",
                          "total de l'actif", "produits d'exploitation"]
    detection_threshold = 0.3

    def get_required_inputs(self):
        return [
            ModuleInput(key="tableau_srm", label="Tableau de synthèse SRM",
                        input_type="file", extensions=[".xlsx", ".xls"],
                        tooltip="Fichier Excel contenant les données SRM"),
        ]

    def validate(self, inputs):
        errors = []
        f = inputs.get("tableau_srm", "")
        if not f:
            errors.append("Le tableau SRM est requis.")
        elif not Path(f).exists():
            errors.append("Fichier introuvable.")
        return (not errors, errors)

    def preview(self, inputs):
        p = inputs.get("tableau_srm", "")
        if p and Path(p).exists():
            try:
                return pd.read_excel(p, nrows=10)
            except Exception:
                pass
        return None

    def _load_srm_engine(self):
        """Charge le moteur SRM via spec_from_file_location (compatible PyInstaller)."""
        script_dir = Path(__file__).parent
        for candidate in ["srmgenV7.py", "srmgen.py"]:
            script_path = script_dir / candidate
            if script_path.exists():
                spec = importlib.util.spec_from_file_location(
                    "srm_engine", str(script_path)
                )
                mod = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(mod)
                return mod
        return None

    def execute(self, inputs, output_dir, progress_callback=None):
        try:
            bridge = get_skills_bridge()
            input_path = inputs["tableau_srm"]

            if progress_callback:
                progress_callback(5, "Chargement du moteur SRM...")

            mod = self._load_srm_engine()
            if mod is None:
                return ModuleResult(
                    success=False,
                    message="Aucun script SRM trouvé. Placez srmgenV7.py dans modules/srm_generator/"
                )

            if progress_callback:
                progress_callback(15, "Vérification du répertoire de sortie...")

            # Garantir que le répertoire de sortie existe
            os.makedirs(output_dir, exist_ok=True)

            # Tester les droits d'écriture
            test_file = Path(output_dir) / ".srm_write_test"
            try:
                test_file.write_text("ok")
                test_file.unlink()
            except OSError as e:
                return ModuleResult(
                    success=False,
                    errors=[
                        f"Impossible d'écrire dans le répertoire de sortie :\n{output_dir}\n\n"
                        f"Erreur : {e}\n\nSolutions :\n"
                        "• Vérifiez que le disque n'est pas plein\n"
                        "• Redémarrez l'application\n"
                        "• Choisissez un autre répertoire de sortie"
                    ]
                )

            if progress_callback:
                progress_callback(30, "Détection des blocs dans le fichier Excel...")

            try:
                audit_df = pd.read_excel(input_path)
                audit_report = bridge.audit_srm_data(audit_df)
            except Exception as audit_error:
                audit_report = {
                    "success": False,
                    "findings": [],
                    "warnings": [f"Analyse SRM impossible avant génération : {audit_error}"],
                    "stats": {},
                }

            captured = io.StringIO()
            with redirect_stdout(captured):
                result_path = mod.process_file(input_path, output_dir)

            if progress_callback:
                progress_callback(100, "SRM généré !")

            # Analyser les logs capturés pour info
            debug_lines = [
                l.strip() for l in captured.getvalue().splitlines()
                if l.strip() and any(
                    kw in l.lower()
                    for kw in ["tab", "feuille", "tableau", "detecte", "total", "srm", "ok"]
                )
            ]

            if result_path and Path(result_path).exists():
                msg = f"SRM généré : {Path(result_path).name}"
                if debug_lines:
                    msg += "\n\nTraitement :\n" + "\n".join(
                        f"• {l}" for l in debug_lines[:6]
                    )
                return ModuleResult(
                    success=True,
                    output_path=str(result_path),
                    message=msg,
                    warnings=audit_report.get("findings", [])[:5] + audit_report.get("warnings", [])[:5],
                    stats={
                        "Lignes source": audit_report.get("stats", {}).get("total_rows", 0),
                        "Colonnes source": audit_report.get("stats", {}).get("total_columns", 0),
                        "Libellés manquants": audit_report.get("stats", {}).get("missing_labels", 0),
                        "Libellés dupliqués": audit_report.get("stats", {}).get("duplicate_labels", 0),
                        "Valeurs numériques manquantes": audit_report.get("stats", {}).get("missing_numeric_values", 0),
                        "Variations incohérentes": audit_report.get("stats", {}).get("variation_mismatches", 0),
                        "Variations atypiques": audit_report.get("stats", {}).get("outlier_variations", 0),
                    }
                )
            else:
                err = "Aucun tableau détecté dans le fichier Excel."
                if debug_lines:
                    err += "\n\nDétails :\n" + "\n".join(f"• {l}" for l in debug_lines)
                return ModuleResult(
                    success=False,
                    message=err,
                    warnings=audit_report.get("findings", [])[:5] + audit_report.get("warnings", [])[:5],
                    stats={
                        "Lignes source": audit_report.get("stats", {}).get("total_rows", 0),
                        "Variations incohérentes": audit_report.get("stats", {}).get("variation_mismatches", 0),
                    },
                )

        except Exception as e:
            msg = str(e)
            if "File is not a zip file" in msg:
                msg = (
                    "Le fichier sélectionné n'est pas un fichier Excel valide (.xlsx).\n"
                    "Assurez-vous d'utiliser un fichier Excel créé avec Excel ou "
                    "un logiciel compatible."
                )
            elif "Permission" in msg or "permission" in msg.lower() or "access" in msg.lower():
                msg = (
                    f"Erreur de permissions lors de l'écriture dans :\n{output_dir}\n\n"
                    f"Erreur technique : {msg}\n\nSolutions :\n"
                    "• Vérifiez que le disque n'est pas plein\n"
                    "• Redémarrez l'application\n"
                    "• Essayez un autre répertoire de sortie"
                )
            else:
                msg = f"Erreur lors du traitement SRM :\n{msg}"
            return ModuleResult(success=False, errors=[msg])
