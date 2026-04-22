"""
Module : Réconciliation BG vs Liasse Fiscale
"""
import datetime
from pathlib import Path

from modules.base_module import BaseModule, ModuleInput, ModuleResult
from modules.reconciliation_bg_liasse.reconciliation import (
    export_reconciliation_report,
    run_reconciliation,
)


class ReconciliationBGLiasse(BaseModule):
    name = "Réconciliation BG vs Liasse"
    description = (
        "Rapprochement automatique entre Balance Générale et Liasse Fiscale.\n"
        "Extraction des rubriques clés (ACTIF/PASSIF/CPC) et calcul des écarts."
    )
    category = "Comptabilité"
    version = "1.0"
    help_text = (
        "ENTRÉES :\n"
        "• Balance Générale (Excel)\n"
        "• Liasse Fiscale (Excel ou PDF)\n\n"
        "SORTIES :\n"
        "• Tableau de réconciliation BG vs Liasse\n"
        "• Rapport Excel formaté (OK / Attention / Critique)"
    )
    detection_keywords = [
        "balance générale", "compte", "liasse", "actif", "passif", "cpc", "chiffre d'affaires"
    ]
    detection_threshold = 0.3

    def get_required_inputs(self):
        return [
            ModuleInput(
                key="fichier_bg",
                label="Balance Générale (Excel)",
                input_type="file",
                extensions=[".xlsx", ".xls", ".xlsm"],
                tooltip="Sélectionnez le fichier Excel de la Balance Générale.",
            ),
            ModuleInput(
                key="fichier_liasse",
                label="Liasse Fiscale (Excel ou PDF)",
                input_type="file",
                extensions=[".xlsx", ".xls", ".pdf"],
                tooltip="Sélectionnez la liasse fiscale (Excel/PDF).",
            ),
        ]

    def validate(self, inputs):
        errors = []
        bg = inputs.get("fichier_bg", "")
        liasse = inputs.get("fichier_liasse", "")

        if not bg:
            errors.append("Le fichier Balance Générale est requis.")
        elif not Path(bg).exists():
            errors.append(f"Fichier BG introuvable : {bg}")

        if not liasse:
            errors.append("Le fichier Liasse Fiscale est requis.")
        elif not Path(liasse).exists():
            errors.append(f"Fichier Liasse introuvable : {liasse}")

        if bg and Path(bg).suffix.lower() not in {".xlsx", ".xls", ".xlsm"}:
            errors.append("Le fichier BG doit être au format Excel.")
        if liasse and Path(liasse).suffix.lower() not in {".xlsx", ".xls", ".pdf"}:
            errors.append("Le fichier liasse doit être au format Excel ou PDF.")

        return (not errors, errors)

    def preview(self, inputs):
        bg = inputs.get("fichier_bg", "")
        liasse = inputs.get("fichier_liasse", "")
        if not bg or not liasse or not Path(bg).exists() or not Path(liasse).exists():
            return None
        try:
            bundle = run_reconciliation(bg, liasse)
            return bundle.reconciliation_df
        except Exception:
            return None

    def execute(self, inputs, output_dir, progress_callback=None):
        try:
            bg_path = inputs["fichier_bg"]
            liasse_path = inputs["fichier_liasse"]

            if progress_callback:
                progress_callback(10, "Chargement Balance Générale...")

            bundle = run_reconciliation(bg_path, liasse_path)

            if progress_callback:
                progress_callback(75, "Génération du rapport Excel...")

            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = (
                Path(output_dir) / f"Reconciliation_BG_Liasse_{ts}.xlsx"
            )
            report_path = export_reconciliation_report(
                bundle.reconciliation_df,
                output_file,
                bg_count=len(bundle.bg_accounts),
                liasse_count=len(bundle.liasse_rubrics),
            )

            if progress_callback:
                progress_callback(100, "Réconciliation terminée.")

            crit = int((bundle.reconciliation_df["Sévérité"] == "CRITIQUE").sum())
            warn = int((bundle.reconciliation_df["Sévérité"] == "ATTENTION").sum())

            return ModuleResult(
                success=True,
                output_path=str(report_path),
                message="Réconciliation BG vs Liasse terminée avec succès.",
                stats={
                    "Comptes BG": len(bundle.bg_accounts),
                    "Rubriques liasse": len(bundle.liasse_rubrics),
                    "Rubriques rapprochées": len(bundle.reconciliation_df),
                    "Écarts critiques": crit,
                    "Écarts attention": warn,
                    "Ecart absolu cumulé": f"{bundle.reconciliation_df['Ecart'].abs().sum():,.2f}",
                    "Rapport": str(report_path),
                },
            )
        except Exception as exc:
            return ModuleResult(success=False, errors=[str(exc)])
