"""
Module : Extraction IR
Wrapper pour extract_ir.py — Extraction avis de versement IR
Place extract_ir.py dans ce dossier pour activer.
"""
import sys, importlib
import importlib.util
from pathlib import Path
from modules.base_module import BaseModule, ModuleInput, ModuleResult


class ExtractionIR(BaseModule):

    name = "Extraction IR"
    description = (
        "Extraction des avis de versement IR (PDF natifs + scannés).\n"
        "Mois, montant IR, date de versement par fichier."
    )
    category = "Fiscalité"
    version = "1.0"
    help_text = (
        "ENTRÉE : PDF des avis de versement / accusés de dépôt IR\n"
        "Le mois est détecté depuis le nom du fichier OU le texte.\n\n"
        "SORTIE : Excel avec 3 colonnes :\n"
        "• Mois | Montant IR (MAD) | Date de versement\n\n"
        "FORMATS :\n"
        "• Bordereau-avis de versement (DGI)\n"
        "• Accusé de dépôt / réception"
    )
    detection_keywords = ["IR", "impôt sur le revenu", "avis de versement",
                          "retenue à la source", "bordereau"]
    detection_threshold = 0.4

    def get_required_inputs(self):
        return [
            ModuleInput(key="fichier_ir", label="Fichier(s) IR PDF",
                        input_type="file", extensions=[".pdf"],
                        required=True, multiple=True,
                        tooltip="Sélectionnez un ou plusieurs fichiers PDF IR."),
            ModuleInput(key="societe", label="Nom de la société",
                        input_type="text", default=""),
        ]

    def get_param_schema(self):
        return []  # Masqué pour l'utilisateur

    def validate(self, inputs):
        errors = []
        fichier = inputs.get("fichier_ir", "")
        if not fichier:
            errors.append("Sélectionnez au moins un fichier PDF IR.")
        if fichier:
            if isinstance(fichier, list):
                for path in fichier:
                    if not Path(path).exists():
                        errors.append(f"Fichier introuvable : {path}")
            elif not Path(fichier).exists():
                errors.append(f"Fichier introuvable : {fichier}")
        return (not errors, errors)

    def preview(self, inputs):
        return None

    def execute(self, inputs, output_dir, progress_callback=None):
        try:
            script_candidates = [
                "extract_ir.py",
                "main8.py",
            ]
            script = None
            for candidate in script_candidates:
                path = Path(__file__).parent / candidate
                if path.exists():
                    script = path
                    break

            if script is None:
                return ModuleResult(success=False,
                    message="Aucun script IR trouvé. Placez main8.py ou extract_ir.py dans modules/extraction_ir/")

            if progress_callback: progress_callback(5, "Chargement module IR...")

            spec = importlib.util.spec_from_file_location("modules.extraction_ir.ir_script", str(script))
            if spec is None or spec.loader is None:
                return ModuleResult(success=False,
                    errors=[f"Impossible de charger le script IR : {script.name}"])
            ir_mod = importlib.util.module_from_spec(spec)
            sys.modules[spec.name] = ir_mod
            spec.loader.exec_module(ir_mod)

            ir_mod.OUTPUT_DIR = Path(output_dir)

            fichiers = inputs.get("fichier_ir", "")
            pdf_paths = []
            if isinstance(fichiers, list):
                pdf_paths = [Path(p) for p in fichiers if Path(p).exists()]
            elif fichiers:
                pdf_paths = [Path(fichiers)]

            if not pdf_paths:
                return ModuleResult(success=False, message="Aucun PDF IR trouvé.")

            if progress_callback: progress_callback(10, "Traitement des fichiers IR...")
            declarations = []
            for i, pdf in enumerate(pdf_paths):
                if progress_callback:
                    pct = 10 + int(70 * i / len(pdf_paths))
                    progress_callback(pct, f"Traitement {pdf.name} ({i+1}/{len(pdf_paths)})...")
                declarations.append(ir_mod.process_ir_file(pdf))

            valides = [d for d in declarations if d.montant]
            if not valides:
                return ModuleResult(success=False, message="Aucun montant IR extrait.")

            if progress_callback: progress_callback(85, "Génération Excel...")
            societe = inputs.get("societe", "") or "SOCIETE"
            out = ir_mod.generate_excel(valides, societe, Path(output_dir))

            if progress_callback: progress_callback(100, "Terminé !")

            total = sum(d.montant for d in valides if d.montant)
            return ModuleResult(
                success=True, output_path=str(out),
                message=f"{len(valides)} avis IR extraits.",
                stats={
                    "Avis traités": len(declarations),
                    "Avec montant": len(valides),
                    "Total IR": f"{total:,.2f} MAD",
                }
            )
        except Exception as e:
            return ModuleResult(success=False, errors=[str(e)])
