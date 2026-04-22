from pathlib import Path
from modules.base_module import BaseModule, ModuleInput, ModuleResult
from .reconciliation import run_reconciliation


class ReconciliationBGLiasse(BaseModule):
    name = "Réconciliation BG-Liasse"
    description = (
        "Rapprochement Balance Générale vs Liasse Fiscale avec calcul des écarts, "
        "mapping des comptes et génération d'un rapport Excel formaté."
    )
    category = "Comptabilité"
    version = "1.0"
    help_text = (
        "ENTRÉES: BG Excel (.xlsx/.xls) et Liasse (Excel ou PDF).\n"
        "SORTIE: Rapport Excel avec tableau de réconciliation, indicateurs et statuts d'écarts."
    )
    detection_keywords = [
        "balance générale",
        "num compte",
        "solde débit",
        "solde crédit",
        "liasse",
        "fiscal",
    ]
    detection_threshold = 0.25

    @staticmethod
    def _as_list(value):
        if value is None:
            return []
        if isinstance(value, list):
            return value
        return [value]

    def get_required_inputs(self):
        return [
            ModuleInput(
                key="fichier_bg",
                label="Balance Générale (1 ou plusieurs fichiers)",
                input_type="file",
                extensions=[".xlsx", ".xls"],
                required=True,
                multiple=True,
            ),
            ModuleInput(
                key="fichier_liasse",
                label="Liasse Fiscale (1 ou plusieurs fichiers)",
                input_type="file",
                extensions=[".xlsx", ".xls", ".pdf"],
                required=True,
                multiple=True,
            ),
        ]

    def validate(self, inputs):
        errors = []

        bg_files = inputs.get("fichier_bg")
        liasse_files = inputs.get("fichier_liasse")

        if not bg_files:
            errors.append("Fichier BG requis.")
        if not liasse_files:
            errors.append("Fichier Liasse requis.")

        for path in self._as_list(bg_files):
            p = Path(path)
            if not p.exists():
                errors.append(f"Fichier BG introuvable : {path}")
            elif p.suffix.lower() not in {".xlsx", ".xls"}:
                errors.append(f"Format BG non supporté : {path}")

        for path in self._as_list(liasse_files):
            p = Path(path)
            if not p.exists():
                errors.append(f"Fichier Liasse introuvable : {path}")
            elif p.suffix.lower() not in {".xlsx", ".xls", ".pdf"}:
                errors.append(f"Format Liasse non supporté : {path}")

        return (not errors, errors)

    def preview(self, inputs):
        return None

    def execute(self, inputs, output_dir, progress_callback=None):
        try:
            bg_files = self._as_list(inputs.get("fichier_bg"))
            liasse_files = self._as_list(inputs.get("fichier_liasse"))

            if not bg_files or not liasse_files:
                return ModuleResult(success=False, message="Fichiers BG et Liasse requis.")

            pairs = []
            if len(bg_files) == len(liasse_files):
                pairs = list(zip(bg_files, liasse_files))
            elif len(bg_files) == 1:
                pairs = [(bg_files[0], lf) for lf in liasse_files]
            elif len(liasse_files) == 1:
                pairs = [(bg, liasse_files[0]) for bg in bg_files]
            else:
                return ModuleResult(
                    success=False,
                    message="Nombre de fichiers BG et Liasse incompatible pour traitement batch.",
                )

            reports = []
            summaries = []
            total_pairs = len(pairs)
            for idx, (bg_path, liasse_path) in enumerate(pairs, start=1):
                if progress_callback:
                    pct = int(100 * (idx - 1) / max(total_pairs, 1))
                    progress_callback(pct, f"Réconciliation {idx}/{total_pairs}...")

                report_path, summary = run_reconciliation(bg_path, liasse_path, output_dir)
                reports.append(report_path)
                summary["bg_file"] = str(bg_path)
                summary["liasse_file"] = str(liasse_path)
                summaries.append(summary)

            if progress_callback:
                progress_callback(100, "Réconciliation terminée")

            global_stats = {
                "traitements": len(summaries),
                "rapports": reports,
                "ok": int(sum(s["ok"] for s in summaries)),
                "ecarts_moyens": int(sum(s["medium"] for s in summaries)),
                "ecarts_significatifs": int(sum(s["large"] for s in summaries)),
                "total_ecart_absolu": float(sum(s["total_abs_ecart"] for s in summaries)),
                "details": summaries,
            }

            return ModuleResult(
                success=True,
                output_path=reports[0] if len(reports) == 1 else output_dir,
                message=f"{len(reports)} rapport(s) généré(s).",
                stats=global_stats,
                warnings=[] if len(reports) == 1 else ["Traitement batch exécuté."],
            )
        except Exception as exc:
            return ModuleResult(success=False, message=str(exc), errors=[str(exc)])
