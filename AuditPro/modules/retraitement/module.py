"""
Module : Retraitement Comptable (BG / GL / AUX)
EN ATTENTE : Copier main.py (v7) dans ce dossier pour activer.
Script disponible dans votre conversation 'Retraitement automatisé de documents Excel'.
"""
import sys, importlib
import pandas as pd
from pathlib import Path
from modules.base_module import BaseModule, ModuleInput, ModuleResult


class RetraitementComptable(BaseModule):

    name = "Retraitement Comptable"
    description = (
        "Retraitement automatique BG / GL / Balances Auxiliaires.\n"
        "Sage 100, Dynamics 365/Navision, Dynamics AX.\n"
        "Détection auto du format + rapport de contrôle."
    )
    category = "Comptabilité"
    version = "8.0"
    help_text = (
        "FORMATS : Grand Livre, Balance Generale, Balance Auxiliaire\n"
        "SORTIE : fichier .xlsx avec feuilles donnees_nettoyees, resume, anomalies\n\n"
        "Detection auto + nettoyage universel + mapping intelligent"
    )
    detection_keywords = ["grand-livre", "balance des comptes", "balance générale",
                          "nominal", "sector", "C.j", "intitulé",
                          "solde initial", "mouvement débit"]
    detection_threshold = 0.3

    def get_required_inputs(self):
        return [
            ModuleInput(key="fichier_comptable", label="Fichier comptable (BG/GL/AUX)",
                        input_type="file", extensions=[".xlsx", ".xls", ".xlsm"]),
        ]

    def get_param_schema(self):
        return []  # Masqué pour l'utilisateur

    def validate(self, inputs):
        errors = []
        f = inputs.get("fichier_comptable", "")
        if not f: errors.append("Fichier comptable requis.")
        elif not Path(f).exists(): errors.append("Fichier introuvable.")
        return (not errors, errors)

    def preview(self, inputs):
        p = inputs.get("fichier_comptable", "")
        if p and Path(p).exists():
            try: return pd.read_excel(p, nrows=10)
            except: pass
        return None

    def execute(self, inputs, output_dir, progress_callback=None):
        try:
            script = Path(__file__).parent / "main.py"
            if not script.exists():
                return ModuleResult(success=False,
                    message="main.py (v7) introuvable.\n\n"
                            "Le script de retraitement doit être présent dans le dossier modules/retraitement/")

            if progress_callback:
                progress_callback(10, "Chargement du module retraitement...")

            # Importer le script
            import importlib.util
            spec = importlib.util.spec_from_file_location("retraitement_main", str(script))
            retraitement_mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(retraitement_mod)

            # Collecter les fichiers
            fichier = inputs.get("fichier_comptable", "")
            excel_files = []

            if isinstance(fichier, list):
                excel_files = [Path(p) for p in fichier if Path(p).exists()]
            elif fichier:
                if Path(fichier).exists():
                    excel_files = [Path(fichier)]

            if not excel_files:
                return ModuleResult(success=False, message="Aucun fichier Excel trouvé.")

            if progress_callback:
                progress_callback(30, f"Traitement de {len(excel_files)} fichier(s)...")

            # Traiter les fichiers
            results = retraitement_mod.process_files(excel_files, output_dir)

            if not results:
                return ModuleResult(success=False, message="Aucun retraitement effectué.")

            if progress_callback:
                progress_callback(100, "Terminé !")

            # Calculer les stats globales
            total_lignes = sum(r.get('total_lignes', 0) for r in results)
            total_retraitements = sum(r.get('retraitements_appliques', 0) for r in results)
            total_montant = sum(r.get('montant_total_retraitement', 0) for r in results)

            # Prendre le premier fichier de sortie comme principal
            output_path = results[0].get('output_file', '') if results else ''

            return ModuleResult(
                success=True,
                output_path=output_path,
                message=f"Retraitement effectué sur {len(excel_files)} fichier(s).",
                stats={
                    "Fichiers traités": len(excel_files),
                    "Total lignes": total_lignes,
                    "Retraitements appliqués": total_retraitements,
                    "Montant total retraitement": f"{total_montant:,.2f}",
                    "Fichier sortie": output_path,
                }
            )

        except Exception as e:
            return ModuleResult(success=False, errors=[str(e)])
