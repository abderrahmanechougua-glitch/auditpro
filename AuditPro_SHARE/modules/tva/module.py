"""
Module : Centralisation TVA
Wrapper pour extract_tva_v3.py — Extraction déclarations TVA DGI Maroc
Place extract_tva_v3.py dans ce dossier pour activer.
"""
import sys
import importlib.util
import pandas as pd
from pathlib import Path
from agent.skills_bridge import get_skills_bridge
from modules.base_module import BaseModule, ModuleInput, ModuleResult


class CentralisationTVA(BaseModule):

    _cached_script_path: str | None = None
    _cached_script_mtime_ns: int | None = None
    _cached_script_module = None

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
        "OCR : Détection automatique interne (mode dégradé activé si OCR indisponible)."
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

    @classmethod
    def _load_script_module(cls, script: Path):
        script = script.resolve()
        mtime_ns = script.stat().st_mtime_ns

        if (
            cls._cached_script_module is not None
            and cls._cached_script_path == str(script)
            and cls._cached_script_mtime_ns == mtime_ns
        ):
            return cls._cached_script_module

        module_name = "modules.tva.tva_script"
        spec = importlib.util.spec_from_file_location(module_name, str(script))
        if spec is None or spec.loader is None:
            raise RuntimeError(f"Impossible de charger le script TVA : {script.name}")

        tva_mod = importlib.util.module_from_spec(spec)
        sys.modules[module_name] = tva_mod
        spec.loader.exec_module(tva_mod)

        cls._cached_script_module = tva_mod
        cls._cached_script_path = str(script)
        cls._cached_script_mtime_ns = mtime_ns
        return tva_mod

    def execute(self, inputs, output_dir, progress_callback=None):
        try:
            bridge = get_skills_bridge()
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

            if progress_callback: progress_callback(5, "Chargement du module TVA...")

            try:
                tva_mod = self._load_script_module(script)
            except Exception as e:
                return ModuleResult(success=False, errors=[str(e)])

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
            for i, pdf in enumerate(pdf_files):
                if progress_callback:
                    pct = 10 + int(70 * i / total)
                    progress_callback(pct, f"Traitement {pdf.name} ({i+1}/{total})...")
                try:
                    d = tva_mod.process_file(pdf)
                    if d: declarations.append(d)
                except Exception as e:
                    pass  # Logged dans le script

            if not declarations:
                return ModuleResult(success=False, message="Aucune déclaration extraite.")

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

            xlsx_warnings = []
            canva_profile = bridge.profile_xlsx(str(canva_out)) if canva_out.exists() else {}
            detail_profile = bridge.profile_xlsx(str(detail_out)) if detail_out.exists() else {}

            try:
                # Validation sur échantillon pour éviter une lecture disque complète sur gros exports.
                detail_df = pd.read_excel(detail_out, nrows=4000)
                quality = bridge.validate_excel_data(detail_df)
                xlsx_warnings.extend(quality.get("warnings", [])[:4])
            except Exception as quality_error:
                xlsx_warnings.append(f"Profil qualité TVA indisponible: {quality_error}")

            return ModuleResult(
                success=True,
                output_path=str(canva_out),
                message=f"TVA extraite pour {len(declarations)} mois.",
                warnings=xlsx_warnings[:5],
                stats={
                    "Mois traités": len(declarations),
                    "TVA collectée totale": f"{tot_coll:,.2f}",
                    "TVA déductible totale": f"{tot_ded:,.2f}",
                    "Solde": f"{tot_coll - tot_ded:,.2f}",
                    "Canva feuilles": canva_profile.get("sheet_count", 0),
                    "Détail feuilles": detail_profile.get("sheet_count", 0),
                    "Détail lignes": detail_profile.get("total_rows", 0),
                    "Canva": str(canva_out),
                    "Détail": str(detail_out),
                }
            )
        except Exception as e:
            return ModuleResult(success=False, errors=[str(e)])
