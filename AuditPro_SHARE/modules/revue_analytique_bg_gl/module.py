"""
Revue Analytique BG-GL — module AuditPro.

Génère automatiquement une feuille de revue analytique par compte agrégé
4 chiffres depuis la feuille BG d'un classeur Excel.

Formules dynamiques, totalisation automatique, commentaire analytique standard.
"""
from pathlib import Path

import pandas as pd

from modules.base_module import BaseModule, ModuleInput, ModuleResult

from .bg_reader import BGReader, BGValidationError
from .account_grouper import AccountGrouper
from .formula_builder import FormulaBuilder
from .sheet_writer import SheetWriter
from .recap_sheet_writer import RecapitulativeSheetWriter
from .models import GenerationRunSummary


class RevueAnalytiqueBgGl(BaseModule):
    name = "Revue analytique"
    description = (
        "Génère les feuilles de revue analytique par compte agrégé 4 chiffres "
        "depuis la feuille BG, configurable par classes et exercice."
    )
    category = "Analytique"
    version = "2.0"
    help_text = (
        "ENTRÉE: Classeur Excel (.xlsx/.xls) contenant les colonnes obligatoires "
        "Compte, Intitule, Solde N, Solde N-1. "
        "Saisissez l'exercice (année) et sélectionnez les classes à traiter.\n"
        "Le module détecte automatiquement la feuille et la ligne d'en-tête.\n"
        "SORTIE: Même classeur enrichi d'une feuille analytique par compte 4 chiffres,\n"
        "avec:\n"
        "  • Ligne 1: 'En KMAD' (bold italic Arial 9)\n"
        "  • Ligne 2: En-têtes violet (Compte, Intitulé, Réf, dates de l'exercice)\n"
        "  • Lignes 3+: Données des sous-comptes\n"
        "  • Ligne total: Agrégé 4 chiffres (fond violet)\n"
        "  • Colonne Réf: vide à remplir"
    )
    detection_keywords = [
        "balance générale",
        "compte",
        "intitule",
        "solde n",
        "solde n-1",
        "revue analytique",
        "analytique",
    ]
    detection_threshold = 0.3

    def get_required_inputs(self) -> list[ModuleInput]:
        return [
            ModuleInput(
                key="fichier_bg",
                label="Classeur BG (Balance Générale)",
                input_type="file",
                extensions=[".xlsx", ".xls"],
                required=True,
                multiple=False,
            ),
            ModuleInput(
                key="exercice",
                label="Exercice (année)",
                input_type="text",
                required=True,
                default="2025",
                tooltip="Année de l'exercice (ex: 2025). Les colonnes de date seront 31/12/YYYY et 31/12/YYYY-1.",
            ),
            ModuleInput(
                key="classes",
                label="Classes à traiter (ex: 1,2,6,7 ou 123567)",
                input_type="text",
                required=False,
                default="67",
                tooltip="Saisissez les numéros de classe séparés par des virgules ou sans séparateur. Classes: 1=Immobilisations, 2=Stocks, 3=Tiers, 4=Trésorerie, 5=Capitaux, 6=Charges, 7=Produits. Laisser vide pour 6,7 par défaut.",
            ),
            ModuleInput(
                key="mode",
                label="Mode de génération",
                input_type="multiselect",
                options=[
                    "Feuilles individuelles par compte (4 chiffres)",
                    "Feuilles récapitulatives par poste (3 chiffres)",
                ],
                required=True,
                default=["Feuilles individuelles par compte (4 chiffres)", "Feuilles récapitulatives par poste (3 chiffres)"],
                tooltip="Choisissez les feuilles à générer: individuelles (détail par compte 4 chiffres) et/ou récapitulatives (regroupement par poste 3 chiffres).",
            ),
        ]

    def validate(self, inputs: dict) -> tuple[bool, list[str]]:
        errors = []
        fichier = inputs.get("fichier_bg")
        if not fichier:
            errors.append("Fichier BG requis.")
            return False, errors
        path = Path(fichier)
        if not path.exists():
            errors.append(f"Fichier introuvable : {fichier}")
            return False, errors
        if path.suffix.lower() not in {".xlsx", ".xls"}:
            errors.append(f"Format non supporté : {path.suffix}")
            return False, errors

        try:
            reader = BGReader(str(path))
            reader.validate_only()
        except BGValidationError as exc:
            errors.append(str(exc))
            return False, errors
        except Exception as exc:  # noqa: BLE001
            errors.append(f"Erreur lecture classeur : {exc}")
            return False, errors

        return True, []

    def preview(self, inputs: dict) -> pd.DataFrame | None:
        fichier = inputs.get("fichier_bg")
        if not fichier:
            return None
        try:
            reader = BGReader(str(fichier))
            rows = reader.read_rows()
        except Exception:  # noqa: BLE001
            return None
        grouper = AccountGrouper()
        aggregates = grouper.group(rows)
        data = [
            {"Compte 4": agg.code_4, "Intitulé": agg.label or "", "Sous-comptes": len(agg.sub_accounts)}
            for agg in aggregates
        ]
        return pd.DataFrame(data) if data else None

    def execute(
        self,
        inputs: dict,
        output_dir: str,
        progress_callback=None,
    ) -> ModuleResult:
        def _progress(pct: int, msg: str):
            if progress_callback:
                progress_callback(pct, msg)

        fichier = inputs.get("fichier_bg", "")
        exercice = inputs.get("exercice", "2025")
        classes_text = inputs.get("classes", "67").strip()
        mode = inputs.get("mode", [
            "Feuilles individuelles par compte (4 chiffres)",
            "Feuilles récapitulatives par poste (3 chiffres)",
        ])

        # Normaliser mode (peut être une liste ou une seule chaîne)
        if isinstance(mode, str):
            mode = [mode]

        # Valider qu'au moins un mode est sélectionné
        if not mode:
            return ModuleResult(
                success=False,
                message="Au moins un mode de génération doit être sélectionné.",
                errors=["Sélectionnez 'Feuilles individuelles', 'Feuilles récapitulatives' ou les deux."],
            )

        class_labels = {
            "1": "1 - Immobilisations",
            "2": "2 - Stocks",
            "3": "3 - Tiers",
            "4": "4 - Trésorerie",
            "5": "5 - Capitaux",
            "6": "6 - Charges",
            "7": "7 - Produits",
        }

        try:
            # Accepte format "1,2,6,7" ou "123567" (avec ou sans virgules)
            if "," in classes_text:
                class_nums = [s.strip() for s in classes_text.split(",") if s.strip()]
            else:
                class_nums = list(classes_text) if classes_text else ["6", "7"]

            classes_input = [class_labels.get(num, f"{num}") for num in class_nums if num in class_labels]
            if not classes_input:
                classes_input = ["6 - Charges", "7 - Produits"]
        except Exception:
            classes_input = ["6 - Charges", "7 - Produits"]

        summary = GenerationRunSummary(success=False)

        try:
            _progress(5, "Lecture et validation de la feuille BG…")
            reader = BGReader(fichier)
            rows = reader.read_rows()

            # Créer un nouveau workbook pour la sortie (évite les verrous d'entrée)
            from openpyxl import Workbook
            wb = Workbook()
            wb.remove(wb.active)  # Supprimer la feuille par défaut
            # Forcer Excel à recalculer les formules
            wb.calculation.calcMode = 'auto'

            summary.invalid_rows_count = reader.invalid_rows_count
            summary.invalid_rows_examples = reader.invalid_rows_examples
            if reader.warnings:
                summary.warnings.extend(reader.warnings)

            _progress(20, "Groupement des comptes agrégés…")
            grouper = AccountGrouper(classes=classes_input)
            aggregates = grouper.group(rows)
            summary.detected_aggregate_accounts_count = len(aggregates)

            _progress(30, f"{len(aggregates)} comptes agrégés détectés. Génération des feuilles…")

            # Déterminer quels modes générer
            generate_individual = "Feuilles individuelles par compte (4 chiffres)" in mode
            generate_recap = "Feuilles récapitulatives par poste (3 chiffres)" in mode

            sheets_generated = 0

            # === Génération des feuilles individuelles (4 chiffres) ===
            if generate_individual:
                _progress(35, "Génération des feuilles individuelles par compte (4 chiffres)…")

                formula_builder = FormulaBuilder(
                    bg_col_map=reader.col_map,
                    bg_last_row=reader.last_data_row,
                    has_ref_col=reader.has_ref_col,
                    bg_sheet_name=reader.sheet_name or "BG",
                )
                writer = SheetWriter(
                    wb,
                    formula_builder=formula_builder,
                    replace_existing=True,
                    bg_rows=rows,
                    exercice=exercice,
                )
                total = len(aggregates)
                for i, agg in enumerate(aggregates, 1):
                    writer.write_sheet(agg)
                    if i % 10 == 0 or i == total:
                        pct = 35 + int(25 * i / total)
                        _progress(pct, f"Feuilles individuelles générées : {i}/{total}")

                sheets_generated = total

            # === Génération des feuilles récapitulatives (3 chiffres) ===
            if generate_recap:
                _progress(65, "Génération des feuilles récapitulatives par poste (3 chiffres)…")

                recap_writer = RecapitulativeSheetWriter(wb, aggregates)
                recap_writer.generate_all()

                # Compter le nombre de feuilles récapitulatives (postes uniques)
                postes = set(agg.code_4[:3] for agg in aggregates)
                sheets_generated += len(postes)

                _progress(90, f"Feuilles récapitulatives générées : {len(postes)} poste(s)")

            summary.generated_sheets_count = sheets_generated

            # Créer un nouveau fichier avec suffixe pour éviter les conflits de permission
            input_name = Path(fichier).stem
            input_ext = Path(fichier).suffix
            output_filename = f"{input_name}_revue_analytique{input_ext}"
            output_dir_path = Path(output_dir)

            # Vérifier/créer le répertoire de sortie
            output_dir_path.mkdir(parents=True, exist_ok=True)

            output_path = output_dir_path / output_filename

            # Sauvegarder avec gestion d'erreur explicite
            try:
                wb.save(str(output_path))
            except PermissionError as pe:
                # Si permission denied, essayer de supprimer le fichier existant d'abord
                if output_path.exists():
                    try:
                        output_path.unlink()
                        wb.save(str(output_path))
                    except Exception as e2:
                        raise Exception(f"Impossible de créer/remplacer le fichier : {output_path}. Vérifiez que le fichier n'est pas ouvert dans Excel. Erreur: {e2}") from e2
                else:
                    raise Exception(f"Permission refusée pour écrire dans {output_path}. Vérifiez les permissions du répertoire.") from pe

            _progress(100, "Génération terminée.")

            summary.success = True
            msg = (
                f"{summary.generated_sheets_count} feuille(s) générée(s) pour "
                f"{summary.detected_aggregate_accounts_count} compte(s) agrégé(s)."
            )
            if summary.invalid_rows_count:
                msg += f" {summary.invalid_rows_count} ligne(s) ignorée(s) (compte invalide)."

            return ModuleResult(
                success=True,
                output_path=str(output_path),
                message=msg,
                stats={
                    "generated_sheets_count": summary.generated_sheets_count,
                    "detected_aggregate_accounts_count": summary.detected_aggregate_accounts_count,
                    "invalid_rows_count": summary.invalid_rows_count,
                },
                warnings=summary.warnings,
                errors=[],
            )

        except BGValidationError as exc:
            summary.errors.append(str(exc))
            return ModuleResult(
                success=False,
                message=str(exc),
                errors=[str(exc)],
            )
        except Exception as exc:  # noqa: BLE001
            summary.errors.append(str(exc))
            return ModuleResult(
                success=False,
                message=f"Erreur inattendue : {exc}",
                errors=[str(exc)],
            )
