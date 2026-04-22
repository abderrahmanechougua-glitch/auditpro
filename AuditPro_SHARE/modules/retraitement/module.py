"""Module : Retraitement Comptable (BG / GL / AUX).

Wrapper UI aligné sur le moteur refactorisé v2.0.
Optionnellement, un lettrage automatique peut être lancé après retraitement.
"""

import datetime
from pathlib import Path

import pandas as pd

from modules.base_module import BaseModule, ModuleInput, ModuleResult
from modules.retraitement import process_file as process_retraitement_file
from modules.retraitement import process_gl as process_retraitement_gl


def _safe_join(values, sep=", "):
    """Assemble une sequence potentiellement heterogene en texte robuste."""
    if not values:
        return ""
    return sep.join(str(v) for v in values)


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
        "FORMATS : Sage 100 (cellules fusionnées), D365, AX\n"
        "SORTIE : Fichier normalisé + Rapport_Controle.xlsx\n\n"
        "Option : Lancer un lettrage automatique après retraitement (GL)."
    )
    detection_keywords = ["grand-livre", "balance des comptes", "balance générale",
                          "nominal", "sector", "C.j", "intitulé",
                          "solde initial", "mouvement débit"]
    detection_threshold = 0.3

    def get_required_inputs(self):
        return [
            ModuleInput(key="fichier_comptable", label="Fichier comptable (BG/GL/AUX)",
                        input_type="file", extensions=[".xlsx", ".xls", ".xlsm", ".xlsb", ".csv"]),
        ]

    def get_param_schema(self):
        return [
            {
                "key": "param_doc_type",
                "label": "Type de document",
                "type": "combo",
                "options": ["Auto-détection", "GL", "BG", "AUX"],
                "default": "Auto-détection",
                "tooltip": "Force le type si l'auto-détection comprend mal votre fichier."
            },
            {
                "key": "param_lancer_lettrage",
                "label": "Lancer lettrage après retraitement",
                "type": "combo",
                "options": ["Non", "Oui"],
                "default": "Non",
                "tooltip": "Si Oui et si type GL, lance automatiquement le module de lettrage."
            }
        ]

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
            if progress_callback:
                progress_callback(10, "Chargement du module retraitement...")

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

            doc_type_hint = inputs.get("param_doc_type", "Auto-détection")
            doc_type_for_engine = None if doc_type_hint == "Auto-détection" else doc_type_hint
            run_lettrage = inputs.get("param_lancer_lettrage", "Non") == "Oui"

            # Traiter les fichiers un par un pour fournir une progression visible.
            results = []
            total_files = len(excel_files)
            for idx, excel_file in enumerate(excel_files, start=1):
                if progress_callback:
                    # Réserver 30->90% au retraitement principal.
                    percent = 30 + int((idx - 1) * 60 / max(total_files, 1))
                    progress_callback(percent, f"Retraitement {idx}/{total_files}: {excel_file.name}")

                if doc_type_for_engine == "GL":
                    result = process_retraitement_gl(excel_file, output_dir=output_dir)
                else:
                    result = process_retraitement_file(
                        excel_file,
                        output_dir=output_dir,
                        doc_type_hint=doc_type_for_engine,
                    )
                results.append(result)

            if progress_callback:
                progress_callback(90, "Finalisation du retraitement...")

            if not results:
                return ModuleResult(success=False, message="Aucun retraitement effectué.")

            if progress_callback:
                progress_callback(100, "Terminé !")

            successful_results = [r for r in results if r.get("success")]
            failed_results = [r for r in results if not r.get("success")]
            if not successful_results:
                errors = [r.get("error", "Erreur inconnue") for r in failed_results]
                return ModuleResult(
                    success=False,
                    errors=errors or ["Aucun fichier n'a pu être retraité."],
                )

            total_lignes = sum(len(r.get("dataframe", pd.DataFrame())) for r in successful_results)
            total_issues = sum(len(r.get("validation_report", pd.DataFrame())) for r in successful_results)
            doc_types = sorted({str(r.get("doc_type", "Inconnu")) for r in successful_results})
            report_files = [r.get("report_file", "") for r in successful_results if r.get("report_file")]

            output_path = report_files[0] if report_files else ""
            warnings = []
            if failed_results:
                warnings.extend([f"Fichier ignoré: {str(r.get('error', 'erreur'))}" for r in failed_results[:3]])

            # Optionnel : lancer lettrage si document GL.
            lettrage_output = ""
            lettrage_message = ""
            if run_lettrage and output_path:
                first_type = successful_results[0].get("doc_type")
                if first_type == "GL":
                    if progress_callback:
                        progress_callback(92, "Lancement lettrage automatique...")
                    from modules.lettrage.module import LettrageGL
                    lettrage_module = LettrageGL()
                    lettrage_result = lettrage_module.execute(
                        {"grand_livre": output_path},
                        output_dir=output_dir,
                        progress_callback=progress_callback,
                    )
                    if lettrage_result.success and lettrage_result.output_path:
                        lettrage_output = lettrage_result.output_path
                        output_path = lettrage_output
                        lettrage_message = "Lettrage automatique effectué."
                    else:
                        msg = _safe_join(
                            lettrage_result.errors or [lettrage_result.message or "échec lettrage"],
                            sep=" ; ",
                        )
                        warnings.append(f"Lettrage non effectué: {msg}")
                else:
                    warnings.append("Lettrage automatique ignoré: type détecté différent de GL.")

            return ModuleResult(
                success=True,
                output_path=output_path,
                message=(
                    f"Retraitement effectué sur {len(successful_results)}/{len(excel_files)} fichier(s)."
                    + (f"\n{lettrage_message}" if lettrage_message else "")
                ),
                warnings=warnings[:8],
                stats={
                    "Fichiers traités": len(successful_results),
                    "Fichiers en échec": len(failed_results),
                    "Type demandé": doc_type_hint,
                    "Type(s) produit(s)": _safe_join(doc_types) if doc_types else "Aucun",
                    "Total lignes": total_lignes,
                    "Problèmes structurels": total_issues,
                    "Rapports générés": len(report_files),
                    "Lettrage auto": "Oui" if lettrage_output else "Non",
                    "Fichier sortie": output_path,
                }
            )

        except Exception as e:
            return ModuleResult(success=False, errors=[str(e)])
