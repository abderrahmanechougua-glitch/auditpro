"""
Module : Centralisation CNSS
Wrapper pour extract_cnss.py — Extraction bordereaux CNSS (RG + AMO)
Place extract_cnss.py dans ce dossier pour activer.
"""
import sys, importlib
import pandas as pd
from pathlib import Path
from modules.base_module import BaseModule, ModuleInput, ModuleResult


class CentralisationCNSS(BaseModule):

    name = "Centralisation CNSS"
    description = (
        "Extraction automatique des bordereaux CNSS (PDF).\n"
        "Régime Général (AF, PS, TFP) + AMO par mois.\n"
        "Cross-checks automatiques des taux et totaux."
    )
    category = "Social"
    version = "2.0"
    help_text = (
        "ENTRÉE : PDF des bordereaux CNSS (2 pages par fichier)\n"
        "Page 1 = Régime Général (AF 6.40%, PS 13.46%, TFP 1.60%)\n"
        "Page 2 = AMO (1.85% + 4.52%)\n\n"
        "SORTIE : Excel avec 12 mois, masses salariales, cotisations,\n"
        "totaux par ligne + formules de contrôle.\n\n"
        "PRÉREQUIS : Tesseract OCR + Poppler (pour PDF scannés)"
    )
    detection_keywords = ["CNSS", "cotisation", "salaire brut", "AMO",
                          "taxe FP", "bordereau", "allocations familiales",
                          "prestations sociales"]
    detection_threshold = 0.3

    def get_required_inputs(self):
        return [
            ModuleInput(key="fichier_cnss", label="Fichier(s) bordereau CNSS (PDF)",
                        input_type="file", extensions=[".pdf"],
                        required=True, multiple=True,
                        tooltip="Sélectionnez un ou plusieurs fichiers PDF CNSS."),
            ModuleInput(key="societe", label="Nom de la société",
                        input_type="text", tooltip="Nom pour le fichier de sortie"),
        ]

    def get_param_schema(self):
        return []  # Masqué pour l'utilisateur

    def validate(self, inputs):
        errors = []
        fichiers = inputs.get("fichier_cnss", "")
        if not fichiers:
            errors.append("Sélectionnez au moins un fichier CNSS.")
        if fichiers:
            if isinstance(fichiers, list):
                for path in fichiers:
                    if not Path(path).exists():
                        errors.append(f"Fichier introuvable : {path}")
            elif not Path(fichiers).exists():
                errors.append(f"Fichier introuvable : {fichiers}")
        if not inputs.get("societe"): errors.append("Le nom de la société est requis.")
        return (not errors, errors)

    def preview(self, inputs):
        return None

    def execute(self, inputs, output_dir, progress_callback=None):
        try:
            script = Path(__file__).parent / "extract_cnss.py"
            if not script.exists():
                return ModuleResult(success=False,
                    message="extract_cnss.py introuvable. Placez-le dans modules/cnss/")

            if progress_callback: progress_callback(5, "Chargement module CNSS...")
            sys.path.insert(0, str(script.parent))
            import extract_cnss as cnss_mod
            importlib.reload(cnss_mod)

            tess = inputs.get("param_tesseract_path", r"C:\Program Files\Tesseract-OCR\tesseract.exe")
            popp = inputs.get("param_poppler_path", r"C:\poppler\poppler-25.12.0\Library\bin")
            cnss_mod.TESSERACT_PATH = tess
            cnss_mod.POPPLER_PATH = popp
            cnss_mod.OUTPUT_DIR = Path(output_dir)

            fichiers = inputs.get("fichier_cnss", "")
            pdf_files = []
            if isinstance(fichiers, list):
                pdf_files = [Path(p) for p in fichiers if Path(p).exists()]
            elif fichiers:
                pdf_files = [Path(fichiers)]

            if not pdf_files:
                return ModuleResult(success=False, message="Aucun PDF CNSS trouvé.")

            declarations = []
            for i, pdf in enumerate(pdf_files):
                if progress_callback:
                    progress_callback(10 + int(70 * i / len(pdf_files)),
                                      f"{pdf.name} ({i+1}/{len(pdf_files)})...")
                d = cnss_mod.process_cnss_file(pdf)
                if d: declarations.append(d)

            if not declarations:
                return ModuleResult(success=False, message="Aucune donnée extraite.")

            if progress_callback: progress_callback(90, "Génération Excel...")
            out = cnss_mod.generate_excel(declarations, inputs["societe"], Path(output_dir))

            if progress_callback: progress_callback(100, "Terminé !")

            tot = sum(d.montant_total for d in declarations)
            return ModuleResult(
                success=True, output_path=str(out),
                message=f"CNSS extraite pour {len(declarations)} mois.",
                stats={
                    "Mois traités": len(declarations),
                    "Total global": f"{tot:,.2f} MAD",
                    "Total AF": f"{sum(d.af_montant for d in declarations):,.2f}",
                    "Total AMO": f"{sum(d.montant_global_amo for d in declarations):,.2f}",
                }
            )
        except Exception as e:
            return ModuleResult(success=False, errors=[str(e)])
