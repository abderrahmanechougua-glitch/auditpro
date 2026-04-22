"""
Module : Centralisation TVA
Wrapper pour extract_tva_v3.py — Extraction déclarations TVA DGI Maroc
Place extract_tva_v3.py dans ce dossier pour activer.
"""
import sys, os, importlib
import importlib.util
import pandas as pd
from pathlib import Path
from modules.base_module import BaseModule, ModuleInput, ModuleResult


class CentralisationTVA(BaseModule):

    name = "Centralisation TVA"
    description = (
        "Extraction automatique des déclarations TVA (PDF DGI Maroc).\n"
        "Supporte les formats Anglo-saxon (brouillon) et Européen (déposée).\n"
        "Génère un canva Excel d'audit + feuille de détail par mois."
    )
    category = "Fiscalité"
    version = "3.0"
    help_text = (
        "ENTRÉE : PDF des déclarations mensuelles TVA (DGI Maroc)\n"
        "Chaque PDF = 1 mois. Formats natifs et scannés supportés.\n\n"
        "SORTIE :\n"
        "• Canva TVA : CA par taux (20%, 14%, 10%, 7%), TVA collectée/déductible\n"
        "• Détail : toutes les lignes extraites + anomalies\n"
        "• Cross-checks : L132 vs somme taux, TVA due = collectée - déductible\n\n"
        "PRÉREQUIS (pour PDF scannés) :\n"
        "• Tesseract OCR installé\n"
        "• Poppler (pdf2image)"
    )
    detection_keywords = ["TVA", "TVA due", "TVA déductible", "base imposable",
                          "prorata", "CA exonéré", "CA imposable", "modèle 10",
                          "déclaration mensuelle"]
    detection_threshold = 0.3

    def get_required_inputs(self):
        return [
            ModuleInput(key="fichier_tva", label="Fichier(s) TVA PDF",
                        input_type="file", extensions=[".pdf"],
                        required=True, multiple=True,
                        tooltip="Sélectionnez un ou plusieurs fichiers PDF de déclarations TVA."),
        ]

    def get_param_schema(self):
        return []  # Masqué pour l'utilisateur

    def validate(self, inputs):
        errors = []
        fichier = inputs.get("fichier_tva", "")
        if not fichier:
            errors.append("Sélectionnez au moins un fichier PDF TVA.")
        if fichier:
            if isinstance(fichier, list):
                for path in fichier:
                    if not Path(path).exists():
                        errors.append(f"Fichier introuvable : {path}")
            elif not Path(fichier).exists():
                errors.append(f"Fichier introuvable : {fichier}")
        return (not errors, errors)

    def preview(self, inputs):
        return None  # PDF → pas de preview tableau

    def execute(self, inputs, output_dir, progress_callback=None):
        try:
            script_candidates = [
                "extract_tva_v3.py",
                "tvaV55.py",
                "centralisation_tva.py"
            ]
            script = None
            for candidate in script_candidates:
                path = Path(__file__).parent / candidate
                if path.exists():
                    script = path
                    break

            if script is None:
                return ModuleResult(success=False,
                    message="Aucun script TVA trouvé. Placez tvaV55.py ou extract_tva_v3.py dans modules/tva/")

            tess = inputs.get("param_tesseract_path", r"C:\Program Files\Tesseract-OCR\tesseract.exe")
            popp = inputs.get("param_poppler_path", r"C:\poppler\poppler-25.12.0\Library\bin")

            if progress_callback: progress_callback(5, "Chargement du module TVA...")

            spec = importlib.util.spec_from_file_location("modules.tva.tva_script", str(script))
            if spec is None or spec.loader is None:
                return ModuleResult(success=False,
                    errors=[f"Impossible de charger le script TVA : {script.name}"])
            tva_mod = importlib.util.module_from_spec(spec)
            sys.modules[spec.name] = tva_mod
            spec.loader.exec_module(tva_mod)

            tva_mod.TESSERACT_PATH = tess
            tva_mod.POPPLER_PATH = popp
            tva_mod.OUTPUT_DIR = Path(output_dir)

            # Collecter les PDF
            fichier = inputs.get("fichier_tva", "")
            pdf_files = []

            if isinstance(fichier, list):
                pdf_files = [Path(p) for p in fichier if Path(p).exists()]
            elif fichier:
                if Path(fichier).exists():
                    pdf_files = [Path(fichier)]

            if not pdf_files:
                return ModuleResult(success=False, message="Aucun PDF trouvé.")

            # Traitement
            declarations = []
            total = len(pdf_files)
            warnings = []
            for i, pdf in enumerate(pdf_files):
                if progress_callback:
                    pct = 10 + int(70 * i / total)
                    progress_callback(pct, f"Traitement {pdf.name} ({i+1}/{total})...")
                try:
                    d = tva_mod.process_file(pdf)
                    if d: declarations.append(d)
                except Exception as e:
                    warnings.append(f"Erreur sur {pdf.name} : {e}")

            if not declarations:
                msg = "Aucune déclaration extraite."
                if warnings:
                    msg += "\n" + "\n".join(warnings)
                return ModuleResult(success=False, message=msg, warnings=warnings)

            if progress_callback: progress_callback(85, "Génération du canva Excel...")

            import datetime
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            soc = declarations[0].raison_sociale or "SOCIETE"

            canva_out = Path(output_dir) / f"TVA_Canva_{ts}.xlsx"
            tva_mod.fill_canva(declarations, None, canva_out)

            detail_out = Path(output_dir) / f"TVA_Detail_{ts}.xlsx"
            tva_mod.generate_detail(declarations, detail_out)

            if progress_callback: progress_callback(100, "Terminé !")

            tot_coll = sum(d.ligne_132 or 0 for d in declarations)
            tot_ded = sum(d.ligne_190 or 0 for d in declarations)

            return ModuleResult(
                success=True,
                output_path=str(canva_out),
                message=f"TVA extraite pour {len(declarations)} mois.",
                warnings=warnings,
                stats={
                    "Mois traités": len(declarations),
                    "TVA collectée totale": f"{tot_coll:,.2f}",
                    "TVA déductible totale": f"{tot_ded:,.2f}",
                    "Solde": f"{tot_coll - tot_ded:,.2f}",
                    "Canva": str(canva_out),
                    "Détail": str(detail_out),
                }
            )
        except Exception as e:
            return ModuleResult(success=False, errors=[str(e)])
