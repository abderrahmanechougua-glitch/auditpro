"""
Moteur de lettrage comptable — sans code interactif.
Adapté pour intégration dans AuditPro (GUI).
"""
import pandas as pd
from pathlib import Path
import warnings
warnings.filterwarnings("ignore")


# ============================================================================
# MOTEUR DE LETTRAGE (1-1 / N-1 / 1-N)
# ============================================================================

class SimpleLettrageEngine:
    def __init__(self, df, mapping):
        self.df = df.copy()
        self.mapping = mapping
        self.next_code = "A"
        self.existing_codes = set()

    def _get_next_code(self):
        """Génère le prochain code de lettrage avec gestion des suffixes."""
        # Si continuer_codes est activé, analyser les codes existants
        if self.mapping.get("continuer_codes", False) and not self.existing_codes:
            code_col = self.mapping.get("code_lettre") or "Code_lettre"
            if code_col in self.df.columns:
                existing = self.df[code_col].astype(str).str.strip()
                existing = existing[existing.str.match(r'^[A-Z]+\d*$') & existing.ne("")]
                if not existing.empty:
                    # Extraire les bases (lettres) et trouver le max
                    bases = existing.str.extract(r'^([A-Z]+)\d*$')[0].dropna()
                    if not bases.empty:
                        # Prendre la base la plus utilisée ou la dernière en ordre alpha
                        base_counts = bases.value_counts()
                        self.current_base = base_counts.index[0]

                        # Trouver le suffixe max pour cette base
                        mask = existing.str.match(f'^{self.current_base}\\d+$')
                        suffixes = existing[mask].str.extract(f'^{self.current_base}(\\d+)$')[0].astype(int)
                        self.next_suffix = suffixes.max() + 1 if not suffixes.empty else 1
                    else:
                        self.current_base = "A"
                        self.next_suffix = 1
                else:
                    self.current_base = "A"
                    self.next_suffix = 1
            else:
                self.current_base = "A"
                self.next_suffix = 1
        elif not hasattr(self, 'current_base'):
            self.current_base = "A"
            self.next_suffix = 1

        # Générer le code
        code = f"{self.current_base}{self.next_suffix}"
        self.next_suffix += 1

        # Si on dépasse 99, passer à la base suivante
        if self.next_suffix > 99:
            self.current_base = chr(ord(self.current_base) + 1)
            self.next_suffix = 1

        return code

    def _find_combo(self, arr, target, tol=0.01, max_items=None):
        """Trouve une combinaison d'écritures qui équilibrent la cible."""
        if max_items is None:
            max_items = int(self.mapping.get("max_combinaisons", 15))

        vals = [round(float(x), 2) for x in arr]
        target = round(float(target), 2)
        if len(vals) > max_items:
            return None
        dp = {0.0: []}
        for i, a in enumerate(vals):
            for s, combo in list(dp.items()):
                s2 = round(s + a, 2)
                if s2 <= target + tol and s2 not in dp:
                    dp[s2] = combo + [i]
            if target in dp:
                return dp[target]
        return None

    def _prepare(self):
        df = self.df.copy()

        df["_compte"] = df[self.mapping["compte"]].astype(str).str.extract(r"(\d+)")[0]
        df["_classe"] = df["_compte"].str[0]

        # Utiliser les nouvelles classes à lettrer
        classes_str = self.mapping.get("classes_lettrer", "4,5")
        allowed_classes = set(classes_str.replace(" ", "").split(","))
        df["_valid"] = df["_classe"].isin(allowed_classes) & df["_compte"].str.len().ge(4)

        df["_debit"]  = pd.to_numeric(df[self.mapping["debit"]], errors="coerce").fillna(0)
        df["_credit"] = pd.to_numeric(df[self.mapping["credit"]], errors="coerce").fillna(0)

        df["_exclu"] = False
        # Exclure journaux OD si demandé
        if self.mapping.get("exclure_od", "Non") == "Oui" and self.mapping.get("journal"):
            df["_journal"] = df[self.mapping["journal"]].astype(str).str.upper()
            df["_exclu"] = df["_journal"].eq("OD")

        code_col = self.mapping.get("code_lettre") or "Code_lettre"
        self.mapping["code_lettre"] = code_col
        if code_col not in df.columns:
            df[code_col] = ""

        df["_deja"] = df[code_col].astype(str).str.strip().ne("")
        self.tolerance = float(self.mapping.get("tolerance", 0.05))

        # Ajouter colonne pour codes avec suffixes
        df["_code_base"] = df[code_col].astype(str).str.extract(r"^([A-Z]+)\d*$")[0]
        df["_partiellement_lettre"] = df["_code_base"].notna() & (df[code_col] != df["_code_base"])

        return df

    def run(self, progress_callback=None):
        df = self._prepare()

        total = len(df)
        elig = (df["_valid"] & ~df["_exclu"]).sum()
        todo = (df["_valid"] & ~df["_exclu"] & ~df["_deja"]).sum()

        code_col = self.mapping["code_lettre"]
        comptes = df.loc[df["_valid"] & ~df["_exclu"], "_compte"].dropna().unique()

        stats = {"1-1": 0, "N-1": 0, "1-N": 0}
        total_lettered = 0
        n_comptes = len(comptes)

        for k, cpt in enumerate(comptes, 1):
            if progress_callback and k % 50 == 0:
                pct = 10 + int(75 * k / n_comptes)
                progress_callback(pct, f"Compte {k}/{n_comptes} — {cpt}")

            grp = df[(df["_compte"] == cpt) & ~df["_deja"]]
            if grp.empty:
                continue

            deb = grp[grp["_debit"] > 0]
            cre = grp[grp["_credit"] > 0]
            if deb.empty and cre.empty:
                continue

            used_d = []
            used_c = []

            # 1-1
            for d_idx, drow in deb.iterrows():
                eq = cre[(abs(cre["_credit"] - drow["_debit"]) <= self.tolerance) & (~cre.index.isin(used_c))]
                if not eq.empty:
                    c_idx = eq.index[0]
                    code = self._get_next_code()
                    df.loc[d_idx, code_col] = code
                    df.loc[c_idx, code_col] = code
                    used_d.append(d_idx)
                    used_c.append(c_idx)
                    stats["1-1"] += 1
                    total_lettered += 2

            rem_deb = deb.drop(used_d, errors="ignore")
            rem_cre = cre.drop(used_c, errors="ignore")

            # N-1
            for c_idx in list(rem_cre.index):
                if c_idx not in rem_cre.index:
                    continue
                credit_val = float(rem_cre.at[c_idx, "_credit"])
                if rem_deb.empty:
                    break
                arr  = rem_deb["_debit"].tolist()
                idxs = rem_deb.index.tolist()
                combo = self._find_combo(arr, credit_val, tol=self.tolerance, max_items=int(self.mapping.get("max_combinaisons", 15)))
                if combo:
                    sel = [idxs[i] for i in combo]
                    code = self._get_next_code()
                    for x in sel + [c_idx]:
                        df.loc[x, code_col] = code
                    stats["N-1"] += 1
                    total_lettered += len(sel) + 1
                    rem_deb = rem_deb.drop(sel, errors="ignore")
                    rem_cre = rem_cre.drop([c_idx], errors="ignore")

            # 1-N
            for d_idx in list(rem_deb.index):
                if d_idx not in rem_deb.index:
                    continue
                debit_val = float(rem_deb.at[d_idx, "_debit"])
                if rem_cre.empty:
                    break
                arr  = rem_cre["_credit"].tolist()
                idxs = rem_cre.index.tolist()
                combo = self._find_combo(arr, debit_val, tol=self.tolerance, max_items=int(self.mapping.get("max_combinaisons", 15)))
                if combo:
                    sel = [idxs[i] for i in combo]
                    code = self._get_next_code()
                    for x in [d_idx] + sel:
                        df.loc[x, code_col] = code
                    stats["1-N"] += 1
                    total_lettered += len(sel) + 1
                    rem_cre = rem_cre.drop(sel, errors="ignore")

        # Nettoyage colonnes temporaires
        temp_cols = [c for c in df.columns if isinstance(c, str) and c.startswith("_")]
        df = df.drop(columns=temp_cols, errors="ignore")

        return df, stats, total_lettered, elig, todo


# ============================================================================
# ANALYSE DES COMPTES
# ============================================================================

def analyse_comptes(df, compte_col, debit_col, credit_col):
    work = df.copy()
    work["_compte_num"] = work[compte_col].astype(str).str.extract(r"(\d+)")[0]
    work["_debit_n"]  = pd.to_numeric(work[debit_col], errors="coerce").fillna(0)
    work["_credit_n"] = pd.to_numeric(work[credit_col], errors="coerce").fillna(0)

    summary = work.groupby("_compte_num").agg(
        total_debit=("_debit_n", "sum"),
        total_credit=("_credit_n", "sum"),
        lignes=("_compte_num", "count")
    ).reset_index()

    summary["solde"] = summary["total_debit"] - summary["total_credit"]
    summary["etat"] = summary["solde"].apply(
        lambda x: "Soldé" if abs(x) < 0.01
        else ("Reste à payer" if x > 0 else "Reste à encaisser")
    )
    return summary


def export_analyse(summary, folder):
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)

    paths = {}
    for fname, df_out in [
        ("analyse_detaillee.xlsx", summary),
        ("comptes_solde.xlsx", summary[summary["etat"] == "Soldé"]),
        ("comptes_non_solde.xlsx", summary[summary["etat"] != "Soldé"]),
    ]:
        p = folder / fname
        df_out.to_excel(p, index=False)
        paths[fname] = str(p)

    return paths


# ============================================================================
# DÉTECTION AUTO DES COLONNES
# ============================================================================

def auto_detect_columns(df):
    """Tente de détecter automatiquement les colonnes compte, débit, crédit."""
    cols_lower = {c.lower().strip(): c for c in df.columns}

    compte = None
    for name in ["compte", "comptes", "n° compte", "num compte", "numero compte",
                  "code compte", "compte gl", "compte auxiliaire", "n°compte"]:
        if name in cols_lower:
            compte = cols_lower[name]
            break

    debit = None
    for name in ["debit", "débit", "dt", "montant d", "montant débit",
                  "mouvement débiteur", "débiteur", "d"]:
        if name in cols_lower:
            debit = cols_lower[name]
            break

    credit = None
    for name in ["credit", "crédit", "ct", "montant c", "montant crédit",
                  "mouvement créditeur", "créditeur", "c"]:
        if name in cols_lower:
            credit = cols_lower[name]
            break

    journal = None
    for name in ["journal", "code journal", "jnl", "journal code"]:
        if name in cols_lower:
            journal = cols_lower[name]
            break

    return compte, debit, credit, journal
