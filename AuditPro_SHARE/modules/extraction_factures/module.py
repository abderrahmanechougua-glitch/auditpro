"""
Module : Extraction Factures
Wrapper pour factext_v2.py — OCR + extraction de factures marocaines
Place factext_v2.py dans ce dossier pour activer.
"""
import logging
import sys
import importlib.util
from pathlib import Path
import pandas as pd
from agent.skills_bridge import get_skills_bridge
from modules.base_module import BaseModule, ModuleInput, ModuleResult


logger = logging.getLogger(__name__)


class ExtractionFactures(BaseModule):

    _cached_script_path: str | None = None
    _cached_script_mtime_ns: int | None = None
    _cached_script_module = None

    name = "Extraction Factures"
    description = (
        "Extraction automatique (OCR) des factures fournisseurs PDF.\n"
        "N° facture, date, HT, TVA, TTC, fournisseur, ICE, IF...\n"
        "Contrôle TVA automatique + score de confiance."
    )
    category = "Comptabilité"
    version = "2.0"
    help_text = (
        "ENTRÉE : PDF de factures (natifs ou scannés)\n"
        "Formats supportés : AFRIQUIA, BEST PROFIL, EDIC, Green Line…\n\n"
        "SORTIE : Excel avec toutes les données extraites :\n"
        "• N° facture, date, fournisseur, description\n"
        "• HT, TVA (taux + montant), TTC\n"
        "• Contrôle TVA (HT+TVA=TTC, taux cohérent)\n"
        "• Montant en lettres (cross-check)\n"
        "• ICE, IF, Patente, RC, CNSS (optionnel)\n"
        "• Score de confiance par facture"
    )
    detection_keywords = ["facture", "fournisseur", "TVA", "HT", "TTC",
                          "montant", "ICE", "net à payer"]
    detection_threshold = 0.4

    def get_required_inputs(self):
        return [
            ModuleInput(key="fichier_factures", label="Fichier(s) facture PDF",
                        input_type="file", extensions=[".pdf"],
                        required=True, multiple=True,
                        tooltip="Sélectionnez un ou plusieurs fichiers PDF de factures."),
            ModuleInput(key="extraire_juridique", label="Extraire données juridiques ?",
                        input_type="combo", options=["Oui", "Non"],
                        default="Oui",
                        tooltip="ICE, IF, Patente, RC, CNSS"),
        ]

    def get_param_schema(self):
        return []  # Masqué pour l'utilisateur

    def validate(self, inputs):
        errors = []
        fichier = inputs.get("fichier_factures", "")
        if not fichier:
            errors.append("Sélectionnez au moins un fichier PDF de facture.")
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

    @classmethod
    def _load_script_module(cls, script: Path):
        """Charge le script OCR une seule fois tant qu'il n'a pas changé."""
        script = script.resolve()
        mtime_ns = script.stat().st_mtime_ns

        if (
            cls._cached_script_module is not None
            and cls._cached_script_path == str(script)
            and cls._cached_script_mtime_ns == mtime_ns
        ):
            return cls._cached_script_module

        module_name = "modules.extraction_factures.factures_script"
        spec = importlib.util.spec_from_file_location(module_name, str(script))
        if spec is None or spec.loader is None:
            raise RuntimeError(f"Impossible de charger le script factures : {script.name}")

        fact_mod = importlib.util.module_from_spec(spec)
        sys.modules[module_name] = fact_mod
        spec.loader.exec_module(fact_mod)

        cls._cached_script_module = fact_mod
        cls._cached_script_path = str(script)
        cls._cached_script_mtime_ns = mtime_ns
        return fact_mod

    def execute(self, inputs, output_dir, progress_callback=None):
        import traceback

        # Configuration d'un logger dédié au module (évite de reconfigurer le root logger)
        log_file = Path(output_dir) / "extraction_factures.log"
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
        module_logger = logging.getLogger("auditpro.extraction_factures")
        module_logger.setLevel(logging.INFO)
        if not any(isinstance(h, logging.FileHandler) and getattr(h, "baseFilename", "") == str(log_file) for h in module_logger.handlers):
            module_logger.addHandler(file_handler)

        try:
            bridge = get_skills_bridge()
            module_logger.info("Début extraction factures")
            module_logger.info("Dossier sortie: %s", output_dir)
            module_logger.info("Inputs: %s", inputs)

            script_candidates = [
                "factext_v2.py",
                "factextv19.py",
            ]
            script = None
            for candidate in script_candidates:
                path = Path(__file__).parent / candidate
                if path.exists():
                    script = path
                    module_logger.info("Script trouvé: %s", script)
                    break

            if script is None:
                error_msg = "Aucun script facture trouvé. Placez factextv19.py ou factext_v2.py dans modules/extraction_factures/"
                module_logger.error(error_msg)
                return ModuleResult(success=False, message=error_msg)

            if progress_callback:
                progress_callback(5, "Chargement module factures...")
            module_logger.info("Chargement du module factures...")

            try:
                fact_mod = self._load_script_module(script)
                module_logger.info("Module factures chargé avec succès")

            except Exception as e:
                error_msg = f"Erreur chargement module factures: {str(e)}"
                module_logger.error(error_msg)
                module_logger.error(traceback.format_exc())
                return ModuleResult(success=False, errors=[error_msg])

            # Vérification pdfplumber
            try:
                import pdfplumber
                module_logger.info("pdfplumber disponible")
            except ImportError:
                error_msg = "pdfplumber indisponible dans cet environnement."
                module_logger.error(error_msg)
                return ModuleResult(success=False, errors=[error_msg])

            fact_mod.OUTPUT_DIR = Path(output_dir)

            # OCR auto-détecté dans factextv19.py via core/ocr_paths.py.
            ocr_available = (
                getattr(fact_mod, "pytesseract", None) is not None
                and getattr(fact_mod, "convert_from_path", None) is not None
            )
            if not ocr_available:
                module_logger.info("OCR indisponible: traitement PDF natifs uniquement")

            extract_jur = inputs.get("extraire_juridique", "Oui") == "Oui"
            module_logger.info("Extraction données juridiques: %s", extract_jur)

            fichiers = inputs.get("fichier_factures", "")
            pdf_files = []
            if isinstance(fichiers, list):
                pdf_files = [Path(p) for p in fichiers if Path(p).exists()]
            elif fichiers:
                pdf_files = [Path(fichiers)]

            if not pdf_files:
                error_msg = "Aucun PDF trouvé."
                module_logger.error(error_msg)
                return ModuleResult(success=False, message=error_msg)

            module_logger.info("PDFs à traiter: %s", [str(p) for p in pdf_files])

            all_invoices = []
            for i, pdf in enumerate(pdf_files):
                try:
                    if progress_callback:
                        progress_callback(10 + int(70 * i / len(pdf_files)),
                                          f"{pdf.name} ({i+1}/{len(pdf_files)})...")

                    module_logger.info("Traitement PDF: %s", pdf.name)
                    invoices = fact_mod.process_pdf(pdf, extract_jur)

                    if invoices:
                        all_invoices.extend(invoices)
                        module_logger.info("Factures extraites de %s: %s", pdf.name, len(invoices))
                    else:
                        module_logger.warning("Aucune facture extraite de %s", pdf.name)
                        if ocr_available:
                            module_logger.info("OCR actif: tentative sur PDFs scannés gérée par le script")
                        else:
                            module_logger.info("OCR indisponible: PDF scannés peuvent ne pas être lisibles")

                except Exception as e:
                    error_msg = f"Erreur traitement {pdf.name}: {str(e)}"
                    module_logger.error(error_msg)
                    module_logger.error(traceback.format_exc())
                    # Continuer avec les autres PDFs
                    continue

            if not all_invoices:
                error_msg = "Aucune facture extraite. Vérifiez que les PDFs contiennent des factures valides."
                if not ocr_available:
                    error_msg += "\n\nOCR indisponible: les PDFs scannés peuvent ne pas être traités."
                module_logger.error(error_msg)
                return ModuleResult(success=False, message=error_msg)

            if progress_callback:
                progress_callback(90, "Génération Excel...")
            logging.info("Génération du fichier Excel...")

            try:
                import datetime
                ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                suffix = "Complet" if extract_jur else "Base"
                out = Path(output_dir) / f"Factures_{suffix}_{ts}.xlsx"
                fact_mod.generate_excel(all_invoices, out, extract_jur)
                module_logger.info("Fichier Excel généré: %s", out)

            except Exception as e:
                error_msg = f"Erreur génération Excel: {str(e)}"
                module_logger.error(error_msg)
                module_logger.error(traceback.format_exc())
                return ModuleResult(success=False, errors=[error_msg])

            if progress_callback:
                progress_callback(100, "Terminé !")

            ok_tva = sum(1 for i in all_invoices if getattr(i, 'tva_control', None) and
                        getattr(i.tva_control, 'statut', '') == "OK")

            xlsx_profile = bridge.profile_xlsx(str(out)) if Path(out).exists() else {}
            quality_warnings = []
            try:
                # Échantillon de validation pour limiter l'I/O sur gros exports.
                fact_df = pd.read_excel(out, nrows=3000)
                quality = bridge.validate_excel_data(fact_df)
                quality_warnings.extend(quality.get("warnings", [])[:4])
            except Exception as quality_error:
                quality_warnings.append(f"Profil qualité factures indisponible: {quality_error}")

            module_logger.info("Extraction terminée: %s factures, TVA OK: %s", len(all_invoices), ok_tva)

            return ModuleResult(
                success=True, output_path=str(out),
                message=f"{len(all_invoices)} facture(s) extraite(s). Consultez {log_file.name} pour les détails.",
                warnings=quality_warnings[:5],
                stats={
                    "Factures extraites": len(all_invoices),
                    "TVA OK": ok_tva,
                    "TVA Anomalie": sum(1 for i in all_invoices if getattr(i, 'tva_control', None) and
                                       "ANOMALIE" in getattr(i.tva_control, 'statut', '')),
                    "Score moyen": f"{sum(getattr(i, 'score_global', 0) for i in all_invoices)//max(1, len(all_invoices))}%",
                    "Classeur factures feuilles": xlsx_profile.get("sheet_count", 0),
                    "Lignes factures": xlsx_profile.get("total_rows", 0),
                    "Fichier log": str(log_file)
                }
            )

        except Exception as e:
            error_msg = f"Erreur inattendue: {str(e)}"
            module_logger.error(error_msg)
            module_logger.error(traceback.format_exc())
            return ModuleResult(success=False, errors=[error_msg])
