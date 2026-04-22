"""
Module : Extraction Factures
Wrapper pour factext_v2.py — OCR + extraction de factures marocaines
Place factext_v2.py dans ce dossier pour activer.
"""
import sys, importlib
import importlib.util
from pathlib import Path
from modules.base_module import BaseModule, ModuleInput, ModuleResult


class ExtractionFactures(BaseModule):

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

    def execute(self, inputs, output_dir, progress_callback=None):
        import logging
        import traceback

        # Configuration du logging
        log_file = Path(output_dir) / "extraction_factures.log"
        logging.basicConfig(
            filename=log_file,
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )

        try:
            logging.info("Début extraction factures")
            logging.info(f"Dossier sortie: {output_dir}")
            logging.info(f"Inputs: {inputs}")

            script_candidates = [
                "factext_v2.py",
                "factextv19.py",
            ]
            script = None
            for candidate in script_candidates:
                path = Path(__file__).parent / candidate
                if path.exists():
                    script = path
                    logging.info(f"Script trouvé: {script}")
                    break

            if script is None:
                error_msg = "Aucun script facture trouvé. Placez factextv19.py ou factext_v2.py dans modules/extraction_factures/"
                logging.error(error_msg)
                return ModuleResult(success=False, message=error_msg)

            if progress_callback:
                progress_callback(5, "Chargement module factures...")
            logging.info("Chargement du module factures...")

            try:
                spec = importlib.util.spec_from_file_location("modules.extraction_factures.factures_script", str(script))
                if spec is None or spec.loader is None:
                    error_msg = f"Impossible de charger le script factures : {script.name}"
                    logging.error(error_msg)
                    return ModuleResult(success=False, errors=[error_msg])

                fact_mod = importlib.util.module_from_spec(spec)
                sys.modules[spec.name] = fact_mod
                spec.loader.exec_module(fact_mod)
                logging.info("Module factures chargé avec succès")

            except Exception as e:
                error_msg = f"Erreur chargement module factures: {str(e)}"
                logging.error(error_msg)
                logging.error(traceback.format_exc())
                return ModuleResult(success=False, errors=[error_msg])

            # Vérification pdfplumber
            try:
                import pdfplumber
                logging.info("pdfplumber disponible")
            except ImportError:
                error_msg = "pdfplumber non installé. Installez avec: pip install pdfplumber"
                logging.error(error_msg)
                return ModuleResult(success=False, errors=[error_msg])

            tess = inputs.get("param_tesseract_path", r"C:\Program Files\Tesseract-OCR\tesseract.exe")
            popp = inputs.get("param_poppler_path", r"C:\poppler\poppler-25.12.0\Library\bin")

            # Vérification Tesseract
            if not Path(tess).exists():
                logging.warning(f"Tesseract introuvable: {tess}")
                logging.info("Utilisation en mode natif uniquement (pas d'OCR)")

            fact_mod.TESSERACT_PATH = tess
            fact_mod.POPPLER_PATH = popp
            fact_mod.OUTPUT_DIR = Path(output_dir)

            extract_jur = inputs.get("extraire_juridique", "Oui") == "Oui"
            logging.info(f"Extraction données juridiques: {extract_jur}")

            fichiers = inputs.get("fichier_factures", "")
            pdf_files = []
            if isinstance(fichiers, list):
                pdf_files = [Path(p) for p in fichiers if Path(p).exists()]
            elif fichiers:
                pdf_files = [Path(fichiers)]

            if not pdf_files:
                error_msg = "Aucun PDF trouvé."
                logging.error(error_msg)
                return ModuleResult(success=False, message=error_msg)

            logging.info(f"PDFs à traiter: {[str(p) for p in pdf_files]}")

            all_invoices = []
            for i, pdf in enumerate(pdf_files):
                try:
                    if progress_callback:
                        progress_callback(10 + int(70 * i / len(pdf_files)),
                                          f"{pdf.name} ({i+1}/{len(pdf_files)})...")

                    logging.info(f"Traitement PDF: {pdf.name}")
                    invoices = fact_mod.process_pdf(pdf, extract_jur)

                    if invoices:
                        all_invoices.extend(invoices)
                        logging.info(f"Factures extraites de {pdf.name}: {len(invoices)}")
                    else:
                        logging.warning(f"Aucune facture extraite de {pdf.name}")
                        # Proposition OCR si échec
                        if Path(tess).exists():
                            logging.info("Proposition d'OCR pour les PDFs scannés")
                        else:
                            logging.warning("Tesseract non configuré - impossible de traiter les PDFs scannés")

                except Exception as e:
                    error_msg = f"Erreur traitement {pdf.name}: {str(e)}"
                    logging.error(error_msg)
                    logging.error(traceback.format_exc())
                    # Continuer avec les autres PDFs
                    continue

            if not all_invoices:
                error_msg = "Aucune facture extraite. Vérifiez que les PDFs contiennent des factures valides."
                if not Path(tess).exists():
                    error_msg += "\n\nPour les PDFs scannés, installez Tesseract OCR."
                logging.error(error_msg)
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
                logging.info(f"Fichier Excel généré: {out}")

            except Exception as e:
                error_msg = f"Erreur génération Excel: {str(e)}"
                logging.error(error_msg)
                logging.error(traceback.format_exc())
                return ModuleResult(success=False, errors=[error_msg])

            if progress_callback:
                progress_callback(100, "Terminé !")

            ok_tva = sum(1 for i in all_invoices if getattr(i, 'tva_control', None) and
                        getattr(i.tva_control, 'statut', '') == "OK")

            logging.info(f"Extraction terminée: {len(all_invoices)} factures, TVA OK: {ok_tva}")

            return ModuleResult(
                success=True, output_path=str(out),
                message=f"{len(all_invoices)} facture(s) extraite(s). Consultez {log_file.name} pour les détails.",
                stats={
                    "Factures extraites": len(all_invoices),
                    "TVA OK": ok_tva,
                    "TVA Anomalie": sum(1 for i in all_invoices if getattr(i, 'tva_control', None) and
                                       "ANOMALIE" in getattr(i.tva_control, 'statut', '')),
                    "Score moyen": f"{sum(getattr(i, 'score_global', 0) for i in all_invoices)//max(1, len(all_invoices))}%",
                    "Fichier log": str(log_file)
                }
            )

        except Exception as e:
            error_msg = f"Erreur inattendue: {str(e)}"
            logging.error(error_msg)
            logging.error(traceback.format_exc())
            return ModuleResult(success=False, errors=[error_msg])
