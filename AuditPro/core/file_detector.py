"""
Détection intelligente du type de fichier.
Analyse les colonnes et noms d'onglets pour proposer le bon module.
"""
import pandas as pd
from pathlib import Path


class FileDetector:
    """Analyse un fichier Excel et retourne le module le plus pertinent."""

    def detect(self, file_path: str, modules: dict) -> list[dict]:
        """
        Analyse un fichier et score chaque module.
        
        Retourne une liste triée :
        [{"module": instance, "score": 0.85, "matched_keywords": [...], "columns": [...]}]
        """
        path = Path(file_path)

        if path.suffix.lower() == ".csv":
            df = pd.read_csv(file_path, nrows=50, encoding="utf-8", errors="replace")
            sheets = {"CSV": df}
        else:
            try:
                xls = pd.ExcelFile(file_path)
                sheets = {}
                for sheet in xls.sheet_names[:5]:  # Max 5 onglets
                    try:
                        sheets[sheet] = pd.read_excel(xls, sheet_name=sheet, nrows=50)
                    except Exception:
                        pass
            except Exception as e:
                return [{"module": None, "score": 0, "error": str(e)}]

        # Collecter tous les noms de colonnes (normalisés)
        all_columns = set()
        all_sheet_names = set()
        for sheet_name, df in sheets.items():
            all_sheet_names.add(sheet_name.lower().strip())
            for col in df.columns:
                all_columns.add(str(col).lower().strip())

        # Scanner les premières lignes pour mots-clés cachés dans les données
        all_values = set()
        for df in sheets.values():
            for col in df.columns:
                for val in df[col].dropna().head(10):
                    all_values.add(str(val).lower().strip())

        # Scorer chaque module
        results = []
        searchable = all_columns | all_sheet_names | all_values

        for mod in modules.values():
            keywords = getattr(mod, 'detection_keywords', [])
            threshold = getattr(mod, 'detection_threshold', 0.5)

            if not keywords:
                continue

            matched = []
            for kw in keywords:
                kw_lower = kw.lower()
                if any(kw_lower in item for item in searchable):
                    matched.append(kw)

            score = len(matched) / len(keywords) if keywords else 0

            if score >= threshold:
                results.append({
                    "module": mod,
                    "score": round(score, 2),
                    "matched_keywords": matched,
                    "columns": list(all_columns),
                    "sheets": list(all_sheet_names),
                })

        # Trier par score décroissant
        results.sort(key=lambda x: x["score"], reverse=True)
        return results

    def get_file_info(self, file_path: str) -> dict:
        """Retourne des métadonnées basiques sur un fichier."""
        path = Path(file_path)
        info = {
            "name": path.name,
            "size_kb": round(path.stat().st_size / 1024, 1),
            "extension": path.suffix.lower(),
            "sheets": [],
            "total_rows": 0,
        }

        if path.suffix.lower() in (".xlsx", ".xls", ".xlsm"):
            try:
                xls = pd.ExcelFile(file_path)
                info["sheets"] = xls.sheet_names
                for sheet in xls.sheet_names[:3]:
                    df = pd.read_excel(xls, sheet_name=sheet)
                    info["total_rows"] += len(df)
            except Exception:
                pass

        return info
