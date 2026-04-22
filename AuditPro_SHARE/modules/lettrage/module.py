"""
Module : Lettrage Grand Livre
Lettrage automatique 1-1 / N-1 / 1-N des écritures comptables.
"""
import datetime
import pandas as pd
from pathlib import Path
from agent.skills_bridge import get_skills_bridge
from modules.base_module import BaseModule, ModuleInput, ModuleResult
from modules.lettrage.lettrage_engine import (
    SimpleLettrageEngine, analyse_comptes, export_analyse, auto_detect_columns
)


class LettrageGL(BaseModule):

    name = "Lettrage Grand Livre"
    description = (
        "Rapprochement automatique des écritures du Grand Livre.\n"
        "Algorithmes 1-1, N-1 et 1-N avec tolérance paramétrable.\n"
        "Génère le GL lettré + analyse des soldes par compte."
    )
    category = "Comptabilité"
    version = "2.0"
    help_text = (
        "ENTRÉE : Grand Livre auxiliaire (Excel)\n"
        "Colonnes requises : Compte, Débit, Crédit\n\n"
        "SORTIE :\n"
        "• Grand Livre lettré (colonne Code_lettre ajoutée)\n"
        "• Analyse des comptes (soldé / reste à payer / reste à encaisser)\n\n"
        "CONSEIL : Laissez les colonnes vides → détection automatique\n"
        "Si la détection échoue, saisissez le nom exact de la colonne.\n\n"
        "TOLÉRANCE : Écart max accepté pour lettrage (ex: 0.05 = 5 centimes)"
    )
    detection_keywords = ["débit", "crédit", "compte", "lettrage", "auxiliaire",
                          "tiers", "solde", "pièce", "journal", "grand livre"]
    detection_threshold = 0.35

    def get_required_inputs(self):
        return [
            ModuleInput(
                key="grand_livre",
                label="Grand Livre (Excel)",
                input_type="file",
                extensions=[".xlsx", ".xls", ".xlsm", ".csv"],
                tooltip="Fichier Excel du Grand Livre auxiliaire"
            ),
            ModuleInput(
                key="col_compte",
                label="Colonne Compte",
                input_type="text",
                required=False,
                default="",
                tooltip="Nom exact de la colonne Compte (vide = détection auto)"
            ),
            ModuleInput(
                key="col_debit",
                label="Colonne Débit",
                input_type="text",
                required=False,
                default="",
                tooltip="Nom exact de la colonne Débit (vide = détection auto)"
            ),
            ModuleInput(
                key="col_credit",
                label="Colonne Crédit",
                input_type="text",
                required=False,
                default="",
                tooltip="Nom exact de la colonne Crédit (vide = détection auto)"
            ),
            ModuleInput(
                key="col_journal",
                label="Colonne Journal (optionnel)",
                input_type="text",
                required=False,
                default="",
                tooltip="Colonne Journal (laisser vide si absente)"
            ),
        ]

    def get_param_schema(self):
        return [
            {
                "key": "param_tolerance",
                "label": "Tolérance de lettrage (€)",
                "type": "number",
                "default": 0.05,
                "min": 0.0,
                "max": 10.0,
                "step": 0.01,
                "tooltip": "Écart maximum accepté pour le lettrage automatique (ex: 0.05 = 5 centimes)"
            },
            {
                "key": "param_classes",
                "label": "Classes comptables",
                "type": "text",
                "default": "34567",
                "tooltip": "Classes de comptes à traiter (ex: 34567 pour tiers)"
            },
            {
                "key": "param_exclure_od",
                "label": "Exclure comptes OD",
                "type": "combo",
                "options": ["Oui", "Non"],
                "default": "Non",
                "tooltip": "Exclure les comptes d'ordre (OD) du lettrage"
            },
            {
                "key": "param_code_lettre_col",
                "label": "Colonne Code_lettre",
                "type": "text",
                "default": "Code_lettre",
                "tooltip": "Nom de la colonne pour les codes de lettrage"
            }
        ]

    def validate(self, inputs):
        errors = []
        p = inputs.get("grand_livre", "")
        if not p:
            errors.append("Le fichier Grand Livre est requis.")
        elif not Path(p).exists():
            errors.append(f"Fichier introuvable : {p}")
        return (not errors, errors)

    @staticmethod
    def _read_input_dataframe(file_path: str, nrows: int | None = None) -> pd.DataFrame:
        """Read GL input file with extension-aware engines and clear dependency errors."""
        ext = Path(file_path).suffix.lower()

        if ext == ".csv":
            return pd.read_csv(file_path, nrows=nrows)

        if ext == ".xlsb":
            try:
                import pyxlsb  # noqa: F401  # Ensures dependency is available for pandas engine.
            except Exception as dep_error:
                raise RuntimeError(
                    "Format .xlsb détecté mais le package 'pyxlsb' est absent. "
                    "Installez-le (pip install pyxlsb) ou convertissez le fichier en .xlsx."
                ) from dep_error
            return pd.read_excel(file_path, nrows=nrows, engine="pyxlsb")

        # For .xlsx/.xlsm and most standard Excel files.
        return pd.read_excel(file_path, nrows=nrows)

    def preview(self, inputs):
        p = inputs.get("grand_livre", "")
        if not p or not Path(p).exists():
            return None
        try:
            return self._read_input_dataframe(p, nrows=10)
        except Exception:
            return None

    def execute(self, inputs, output_dir, progress_callback=None):
        try:
            bridge = get_skills_bridge()
            # ── 1. Chargement du fichier ──────────────────────────────
            p = inputs["grand_livre"]
            if progress_callback:
                progress_callback(5, "Chargement du Grand Livre...")

            try:
                df = self._read_input_dataframe(p)
            except Exception as e:
                return ModuleResult(success=False,
                    errors=[f"Erreur lecture fichier : {e}"])

            if df.empty:
                return ModuleResult(success=False,
                    message="Le fichier est vide.")

            # ── 2. Résolution des colonnes ────────────────────────────
            if progress_callback:
                progress_callback(10, "Détection des colonnes...")

            col_compte = inputs.get("col_compte", "").strip()
            col_debit  = inputs.get("col_debit", "").strip()
            col_credit = inputs.get("col_credit", "").strip()
            col_journal = inputs.get("col_journal", "").strip()

            # Auto-détection si non spécifié
            auto_cpt, auto_deb, auto_cre, auto_jnl = auto_detect_columns(df)

            def _find_optional_col(candidates):
                lower = {str(c).lower().strip(): c for c in df.columns}
                for name in candidates:
                    if name in lower:
                        return lower[name]
                for c in df.columns:
                    cl = str(c).lower()
                    if any(k in cl for k in candidates):
                        return c
                return None

            auto_date = _find_optional_col(["date", "date piece", "date ecriture", "date d'ecriture", "date comptable"])
            auto_piece = _find_optional_col(["piece", "pièce", "num piece", "numero piece", "référence", "reference", "ref"])
            auto_libelle = _find_optional_col(["libellé", "libelle", "description", "intitulé", "intitule", "motif"])

            col_compte  = col_compte  or auto_cpt
            col_debit   = col_debit   or auto_deb
            col_credit  = col_credit  or auto_cre
            col_journal = col_journal or auto_jnl

            # Validation des colonnes
            missing = []
            for col, label in [(col_compte, "Compte"), (col_debit, "Débit"),
                                (col_credit, "Crédit")]:
                if not col:
                    missing.append(label)
                elif col not in df.columns:
                    missing.append(f"{label} ('{col}' introuvable)")
            if missing:
                cols_dispo = ", ".join(df.columns.tolist()[:20])
                return ModuleResult(success=False,
                    errors=[
                        f"Colonnes non trouvées : {', '.join(missing)}",
                        f"Colonnes disponibles : {cols_dispo}"
                    ])

            # ── 3. Construction du mapping ────────────────────────────
            try:
                tolerance = float(
                    inputs.get("param_tolerance", "0.05").replace(",", ".")
                )
            except Exception:
                tolerance = 0.05

            classes_str = inputs.get("param_classes", "34567").strip()
            classes = list(classes_str) if classes_str else list("34567")

            exclure_od = inputs.get("param_exclure_od", "Non") == "Oui"
            code_lettre_col = inputs.get("param_code_lettre_col", "Code_lettre").strip() \
                              or "Code_lettre"

            mapping = {
                "compte":       col_compte,
                "debit":        col_debit,
                "credit":       col_credit,
                "journal":      col_journal if col_journal and col_journal in df.columns else None,
                "date":         auto_date,
                "piece":        auto_piece,
                "libelle":      auto_libelle,
                "exclure_od":   exclure_od,
                "code_lettre":  code_lettre_col,
                "classes":      classes,
                "tolerance":    tolerance,
            }

            audit_report = bridge.audit_accounting_data(
                df,
                key_columns={
                    "account": col_compte,
                    "debit": col_debit,
                    "credit": col_credit,
                    "date": auto_date,
                    "piece": auto_piece,
                }
            )

            # ── 4. Exécution du lettrage ──────────────────────────────
            if progress_callback:
                progress_callback(15, f"Démarrage lettrage ({len(df)} lignes)...")

            engine = SimpleLettrageEngine(df, mapping)
            df_lettré, stats, total_lettered, elig, todo = engine.run(
                progress_callback=progress_callback
            )

            # ── 5. Analyse des comptes ────────────────────────────────
            if progress_callback:
                progress_callback(88, "Analyse des soldes par compte...")

            summary = analyse_comptes(df_lettré, col_compte, col_debit, col_credit)

            # ── 6. Export ─────────────────────────────────────────────
            if progress_callback:
                progress_callback(93, "Export des fichiers...")

            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            out_dir = Path(output_dir)
            out_dir.mkdir(parents=True, exist_ok=True)

            # Grand livre lettré
            stem = Path(p).stem
            out_gl = out_dir / f"{stem}_lettré_{ts}.xlsx"
            df_lettré.to_excel(str(out_gl), index=False)

            # Analyse
            analyse_paths = export_analyse(summary, out_dir / "analyse")

            if progress_callback:
                progress_callback(100, "Terminé !")

            # Statistiques résumées
            n_soldes = int((summary["etat"] == "Soldé").sum())
            n_rap    = int((summary["etat"] == "Reste à payer").sum())
            n_rae    = int((summary["etat"] == "Reste à encaisser").sum())

            return ModuleResult(
                success=True,
                output_path=str(out_gl),
                message=(
                    f"Colonnes utilisées : Compte='{col_compte}', "
                    f"Débit='{col_debit}', Crédit='{col_credit}'\n"
                    f"Tolérance : ±{tolerance:.4f}"
                ),
                warnings=audit_report.get("findings", [])[:5] + audit_report.get("warnings", [])[:5],
                stats={
                    "Lignes traitées":       len(df),
                    "Doublons exacts":       audit_report.get("stats", {}).get("duplicate_rows", 0),
                    "Doublons potentiels":   audit_report.get("stats", {}).get("duplicate_entries", 0),
                    "Ecart débit/crédit":    f"{audit_report.get('stats', {}).get('balance_gap', 0.0):,.2f}",
                    "Montants atypiques":    audit_report.get("stats", {}).get("outlier_rows", 0),
                    "Lignes éligibles":      int(elig),
                    "Lignes lettrées":       total_lettered,
                    "Lettrage 1-1":          stats["1-1"],
                    "Lettrage N-1":          stats["N-1"],
                    "Lettrage 1-N":          stats["1-N"],
                    "Comptes soldés":        n_soldes,
                    "Reste à payer":         n_rap,
                    "Reste à encaisser":     n_rae,
                    "Grand Livre lettré":    str(out_gl),
                }
            )

        except Exception as e:
            import traceback
            return ModuleResult(
                success=False,
                errors=[str(e), traceback.format_exc()[-500:]]
            )
