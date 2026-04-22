"""
Moteur de lettrage comptable — Production-grade.
Refactorisé pour AuditPro (GUI) — audit, traçabilité, performance.

Architecture:
  - CodeGenerator       : séquence Excel-style (A→Z→AA→AB→…), sans collision
  - validate_mapping    : validation précoce du mapping avant exécution
  - LettrageResult      : dataclass résultat complet (df + stats + trace + non lettrés)
  - extract_ref_token   : extraction de référence configurable et sans faux positifs
  - DataPreparator      : préparation vectorisée, colonnes préfixées _AP_ anti-collision
  - find_combo          : recherche combinatoire bornée (greedy + DP limité)
  - MatchingEngine      : moteur de règles modulaire avec journal de traçabilité
  - SimpleLettrageEngine: façade publique, API rétro-compatible
"""
from __future__ import annotations

import logging
import re
import warnings
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

import pandas as pd

warnings.filterwarnings("ignore")
logger = logging.getLogger(__name__)

# Préfixe exclusif pour toutes les colonnes temporaires — évite les collisions
TMP = "_AP_"


# ============================================================================
# CODE GENERATOR — Excel-style (A→Z→AA→AB→…, scalable to infinity)
# ============================================================================

def _n_to_excel_code(n: int) -> str:
    """Entier 1-based → code Excel (1→A, 26→Z, 27→AA, 703→AAA…)."""
    result: list[str] = []
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result.append(chr(65 + rem))
    return "".join(reversed(result))


def _excel_code_to_n(code: str) -> int:
    """Code Excel → entier 1-based."""
    n = 0
    for ch in code.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


class CodeGenerator:
    """
    Générateur de codes lettrage Excel-style, garanti unique et sans limite.
    Thread-safe dans un contexte mono-thread (pandas/GUI).
    """

    def __init__(self, start: int = 1, existing: Optional[set] = None):
        self._counter = start
        self._used: set[str] = set(existing or [])

    @classmethod
    def from_existing_codes(cls, codes: pd.Series) -> "CodeGenerator":
        """Continue après le code existant le plus élevé."""
        valid = set(
            codes.astype(str)
            .str.strip()
            .pipe(lambda s: s[s.str.match(r"^[A-Z]+$")])
        )
        if not valid:
            return cls(start=1)
        max_n = max(_excel_code_to_n(c) for c in valid)
        return cls(start=max_n + 1, existing=valid)

    def next(self) -> str:
        """Retourne le prochain code unique."""
        while True:
            code = _n_to_excel_code(self._counter)
            self._counter += 1
            if code not in self._used:
                self._used.add(code)
                return code


# ============================================================================
# MAPPING VALIDATION
# ============================================================================

_REQUIRED_MAPPING_KEYS = {"compte", "debit", "credit"}


def validate_mapping(mapping: dict, df_columns: list) -> list[str]:
    """
    Valide le mapping avant exécution.
    Retourne une liste de messages d'erreur (vide = OK).
    Ne modifie pas le mapping.
    """
    errors: list[str] = []
    col_set = set(df_columns)

    for key in _REQUIRED_MAPPING_KEYS:
        if key not in mapping:
            errors.append(f"Clé obligatoire manquante dans mapping: '{key}'")
        elif mapping[key] not in col_set:
            errors.append(
                f"Colonne '{mapping[key]}' (mapping['{key}']) introuvable. "
                f"Colonnes disponibles: {sorted(col_set)}"
            )

    # code_lettre is an output column — created by the engine when absent, never validated
    for opt_key in ("libelle", "piece", "journal", "date"):
        val = mapping.get(opt_key)
        if val and val not in col_set:
            errors.append(
                f"Colonne optionnelle '{val}' (mapping['{opt_key}']) introuvable."
            )

    return errors


# ============================================================================
# RESULT DATACLASS
# ============================================================================

@dataclass
class LettrageResult:
    """Résultat complet d'un cycle de lettrage."""
    df: pd.DataFrame             # Grand livre lettré (colonnes temporaires supprimées)
    stats: dict                  # {"1-1": n, "N-1": n, "1-N": n, "ref": n, "transfer": n}
    total_lettered: int          # Nouvelles lignes lettrées durant ce cycle
    eligible_lines: int          # Lignes éligibles au lettrage
    todo_lines: int              # Lignes éligibles non encore lettrées avant exécution
    trace: pd.DataFrame          # Journal d'audit de chaque décision de lettrage
    unmatched: pd.DataFrame      # Lignes éligibles restées sans code


# ============================================================================
# REFERENCE EXTRACTION
# ============================================================================

# Patterns par ordre de spécificité — surchargeables via mapping["ref_patterns"]
_DEFAULT_REF_PATTERNS: list[str] = [
    r"\b[A-Z]{1,6}-\d{3,8}(?:-[A-Z0-9]{1,8})?(?:/\d{1,6})?\b",
    r"\b[A-Z0-9]{2,6}/\d{3,8}\b",
    r"\b[A-Z]{2,6}\d{5,}\b",
]

_REF_BANNED: frozenset = frozenset({
    "VIR", "VIREMENT", "VERSEMENT", "ESPECE", "CHEQUE", "CHQ",
    "OPERATION", "DEBIT", "CREDIT", "FACTURE", "REGLEMENT", "AVOIR",
    "SOLDE", "REPORT", "RAN", "TVA", "TTC", "HT",
})


def extract_ref_token(val: object, patterns: Optional[list] = None) -> str:
    """
    Extrait un token de référence structuré depuis un libellé ou n° de pièce.
    Retourne une chaîne vide si rien de fiable n'est trouvé.
    Patterns configurables pour s'adapter aux conventions de l'entreprise.
    """
    txt = re.sub(r"\s+", " ", str(val or "").upper().strip())
    if not txt or txt in ("NAN", "NONE", "NAT", ""):
        return ""

    for pat in (patterns or _DEFAULT_REF_PATTERNS):
        for tok in re.findall(pat, txt):
            t = tok.strip()
            if len(t) >= 5 and t not in _REF_BANNED:
                return t
    return ""


# ============================================================================
# DATA PREPARATION
# ============================================================================

def _parse_classes(classes_val) -> set:
    """Parse classes_lettrer: chaîne "34567", "3,4,5", ou liste."""
    if isinstance(classes_val, (list, tuple, set)):
        return {str(x).strip() for x in classes_val if str(x).strip()}
    s = str(classes_val).replace(" ", "")
    if "," in s:
        result = {x for x in s.split(",") if x}
    else:
        result = {x for x in s if x}
    return result or set("34567")


class DataPreparator:
    """
    Prépare le DataFrame brut pour le moteur de matching.
    Toutes les colonnes temporaires utilisent le préfixe TMP (_AP_).
    Ne modifie jamais le mapping passé en paramètre.
    """

    def __init__(self, mapping: dict, tolerance: float = 0.05):
        self.mapping = dict(mapping)   # copie défensive
        self.tolerance = tolerance

    def prepare(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        m = self.mapping

        # Compte et classe comptable
        df[TMP + "compte"] = (
            df[m["compte"]].astype(str).str.extract(r"(\d+)")[0].fillna("")
        )
        df[TMP + "classe"] = df[TMP + "compte"].str[:1].fillna("")

        # Éligibilité au lettrage
        allowed_classes = _parse_classes(
            m.get("classes") or m.get("classes_lettrer") or "34567"
        )
        df[TMP + "valid"] = (
            df[TMP + "classe"].isin(allowed_classes)
            & df[TMP + "compte"].str.len().ge(4)
        )

        # Montants (NaN → 0, pas de chaînes "nan")
        df[TMP + "debit"]  = pd.to_numeric(df[m["debit"]],  errors="coerce").fillna(0.0)
        df[TMP + "credit"] = pd.to_numeric(df[m["credit"]], errors="coerce").fillna(0.0)

        # Exclusion journaux OD
        df[TMP + "exclu"] = False
        excl_flag = m.get("exclure_od", False)
        if isinstance(excl_flag, str):
            excl_flag = excl_flag.strip().lower() in {"oui", "true", "1", "yes"}
        if excl_flag and m.get("journal") and m["journal"] in df.columns:
            df[TMP + "exclu"] = df[m["journal"]].astype(str).str.upper().eq("OD")

        # Colonne code lettrage — normalisation des NaN et valeurs vides
        code_col = m.get("code_lettre") or "Code_lettre"
        if code_col not in df.columns:
            df[code_col] = ""
        df[code_col] = (
            df[code_col]
            .astype(str)
            .str.strip()
            .replace({"nan": "", "None": "", "NaN": "", "none": "", "NaT": ""})
        )
        df[TMP + "already_lettered"] = df[code_col].ne("")

        # Clé de référence (pièce prioritaire sur libellé)
        ref_patterns = m.get("ref_patterns")
        df[TMP + "ref_key"] = self._build_ref_keys(df, ref_patterns)

        # Date optionnelle
        date_col = m.get("date")
        if date_col and date_col in df.columns:
            df[TMP + "date"] = pd.to_datetime(df[date_col], errors="coerce")
        else:
            df[TMP + "date"] = pd.NaT

        return df

    def _build_ref_keys(self, df: pd.DataFrame, patterns: Optional[list]) -> pd.Series:
        """Extraction vectorisée des clés de référence (pièce > libellé)."""
        m = self.mapping
        piece_col = m.get("piece")
        lib_col   = m.get("libelle")

        series = pd.Series("", index=df.index, dtype=str)

        if piece_col and piece_col in df.columns:
            series = df[piece_col].map(lambda v: extract_ref_token(v, patterns))

        if lib_col and lib_col in df.columns:
            lib_tokens = df[lib_col].map(lambda v: extract_ref_token(v, patterns))
            # Libellé seulement là où la pièce n'a rien donné
            series = series.where(series.ne(""), lib_tokens)

        return series


# ============================================================================
# COMBINATION SEARCH — Greedy pre-filter + Bounded DP
# ============================================================================

def find_combo(
    values: list,
    target: float,
    tolerance: float = 0.01,
    max_items: int = 15,
) -> Optional[list]:
    """
    Trouve les indices de `values` dont la somme est égale à `target` (±tolerance).
    Stratégie: vérification rapide en O(n), puis DP borné sur items pré-filtrés.
    Retourne None si aucune combinaison trouvée ou si len(values) > max_items.
    """
    if not values or len(values) > max_items:
        return None

    vals = [round(float(v), 2) for v in values]
    target_r = round(float(target), 2)
    tol = round(float(tolerance), 4)

    # Vérification rapide : item unique
    for i, v in enumerate(vals):
        if abs(v - target_r) <= tol:
            return [i]

    # Pré-filtre : ne garder que les items ≤ target + tol (tri décroissant)
    candidates = [(i, v) for i, v in enumerate(vals) if 0 < v <= target_r + tol]
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[1], reverse=True)

    orig_idxs  = [i for i, _ in candidates[:max_items]]
    filt_vals  = [v for _, v in candidates[:max_items]]

    # DP borné avec terminaison anticipée
    dp: dict[float, list] = {0.0: []}
    for pos, a in enumerate(filt_vals):
        for s, combo in list(dp.items()):
            s2 = round(s + a, 2)
            if s2 <= target_r + tol and s2 not in dp:
                dp[s2] = combo + [orig_idxs[pos]]
        # Terminaison anticipée
        for s, combo in dp.items():
            if abs(s - target_r) <= tol and combo:
                return combo

    for s, combo in dp.items():
        if abs(s - target_r) <= tol and combo:
            return combo

    return None


# ============================================================================
# MATCHING ENGINE
# ============================================================================

class MatchingEngine:
    """
    Applique les règles de lettrage en séquence et accumule un journal de traçabilité.

    Règles (dans l'ordre):
      1. Propagation de codes existants par référence commune
      2. Nouveaux codes sur groupes référence débit/crédit équilibrés
      3. Appariement virements/transferts (montant opposé + date proche)
      4. Complétion de groupes partiellement lettrés (résiduel exact)
      5. Lettrage 1-1 par compte (vectorisé via merge)
      6. Lettrage N-1 par compte (DP borné)
      7. Lettrage 1-N par compte (DP borné)
    """

    def __init__(self, mapping: dict, code_gen: CodeGenerator, tolerance: float = 0.05):
        self.mapping    = dict(mapping)
        self.code_gen   = code_gen
        self.tolerance  = tolerance
        self.code_col   = mapping.get("code_lettre") or "Code_lettre"
        self.strict_mode: bool = bool(mapping.get("strict_mode", False))
        self.max_combo: int    = int(mapping.get("max_combinaisons", 15))
        self._trace: list[dict] = []
        self._stats: dict[str, int] = {
            "1-1": 0, "N-1": 0, "1-N": 0, "ref": 0, "transfer": 0
        }

    # ------------------------------------------------------------------ #
    # Entry point
    # ------------------------------------------------------------------ #

    def run(
        self,
        df: pd.DataFrame,
        progress_callback: Optional[Callable] = None,
    ) -> pd.DataFrame:
        """Applique toutes les règles et retourne le DataFrame enrichi."""
        cc = self.code_col

        def _p(pct: int, msg: str) -> None:
            if progress_callback:
                progress_callback(pct, msg)

        _p(5,  "Pré-lettrage: propagation références...")
        self._rule_ref_propagation(df, cc)

        _p(10, "Pré-lettrage: nouveaux codes référence...")
        self._rule_ref_new_codes(df, cc)

        _p(15, "Pré-lettrage: virements/transferts...")
        self._rule_transfer_pairing(df, cc)

        _p(20, "Complétion groupes résiduels...")
        self._rule_residual_completion(df, cc)

        # Actualiser le flag après pré-lettrage
        df[TMP + "already_lettered"] = df[cc].ne("")

        comptes = (
            df.loc[df[TMP + "valid"] & ~df[TMP + "exclu"], TMP + "compte"]
            .dropna()
            .unique()
        )
        n = len(comptes)
        for k, cpt in enumerate(comptes, 1):
            if k % 50 == 0:
                _p(20 + int(70 * k / n), f"Compte {k}/{n} — {cpt}")
            self._match_account(df, cc, cpt)

        _p(95, "Finalisation...")
        return df

    # ------------------------------------------------------------------ #
    # Rule 1: Reference propagation
    # ------------------------------------------------------------------ #

    def _rule_ref_propagation(self, df: pd.DataFrame, cc: str) -> None:
        """Propage un code lettrage existant à toutes les lignes de même référence."""
        base_mask = df[TMP + "ref_key"].ne("") & df[TMP + "valid"] & ~df[TMP + "exclu"]
        ref_with_code = df[base_mask & df[cc].ne("")]
        if ref_with_code.empty:
            return

        # Code dominant par référence (mode)
        ref_to_code = (
            ref_with_code
            .groupby(TMP + "ref_key")[cc]
            .agg(lambda s: s.mode().iloc[0] if not s.mode().empty else s.iloc[0])
        )

        unlettered_mask = base_mask & df[cc].eq("")
        update = df.loc[unlettered_mask, TMP + "ref_key"].map(ref_to_code).dropna()
        if not update.empty:
            df.loc[update.index, cc] = update.values
            for idx, code in update.items():
                self._log(idx, code, "ref_propagation", "ref", 0.85, df)
            self._stats["ref"] += len(update)

    # ------------------------------------------------------------------ #
    # Rule 2: New codes from balanced reference groups
    # ------------------------------------------------------------------ #

    def _rule_ref_new_codes(self, df: pd.DataFrame, cc: str) -> None:
        """Crée un code pour les groupes référence débit/crédit équilibrés."""
        mask = (
            df[TMP + "ref_key"].ne("")
            & df[TMP + "valid"]
            & ~df[TMP + "exclu"]
            & df[cc].eq("")
        )
        ref_df = df[mask]
        if ref_df.empty:
            return

        agg = ref_df.groupby(TMP + "ref_key").agg(
            dsum=(TMP + "debit",  "sum"),
            csum=(TMP + "credit", "sum"),
            has_d=(TMP + "debit",  lambda x: (x > 0).any()),
            has_c=(TMP + "credit", lambda x: (x > 0).any()),
        )
        min_tol = max(self.tolerance, 0.02)
        balanced = agg[agg["has_d"] & agg["has_c"] & (agg["dsum"] - agg["csum"]).abs().le(min_tol)]

        for ref in balanced.index:
            idxs = df.index[mask & (df[TMP + "ref_key"] == ref)]
            code = self.code_gen.next()
            df.loc[idxs, cc] = code
            for idx in idxs:
                self._log(idx, code, "ref_new", "ref", 0.90, df)
            self._stats["ref"] += len(idxs)

    # ------------------------------------------------------------------ #
    # Rule 3: Transfer pairing
    # ------------------------------------------------------------------ #

    def _rule_transfer_pairing(self, df: pd.DataFrame, cc: str) -> None:
        """Rapprochement par montant opposé pour écritures de virement/transfert."""
        lib_col = self.mapping.get("libelle")
        if not lib_col or lib_col not in df.columns:
            return

        kw = r"VIR(?:EMENT)?|VERSEMENT|VERST|TRANSFERT"
        mask = (
            df[TMP + "valid"]
            & ~df[TMP + "exclu"]
            & df[cc].eq("")
            & df[lib_col].astype(str).str.upper().str.contains(kw, regex=True, na=False)
        )
        unl = df[mask]
        if unl.empty:
            return

        deb = unl[unl[TMP + "debit"] > 0]
        cre = unl[unl[TMP + "credit"] > 0]
        if deb.empty or cre.empty:
            return

        used_cre: set = set()
        for d_idx, drow in deb.iterrows():
            amt = round(float(drow[TMP + "debit"]), 2)
            cands = cre[
                (abs(cre[TMP + "credit"].round(2) - amt) <= self.tolerance)
                & (~cre.index.isin(used_cre))
            ]
            if cands.empty:
                continue

            # Tri par proximité de date si disponible
            d_date = drow.get(TMP + "date")
            if pd.notna(d_date) and TMP + "date" in df.columns:
                try:
                    deltas = (cands[TMP + "date"] - d_date).abs()
                    best_idx = deltas.idxmin()
                except Exception:
                    best_idx = cands.index[0]
            else:
                best_idx = cands.index[0]

            code = self.code_gen.next()
            df.loc[d_idx, cc]   = code
            df.loc[best_idx, cc] = code
            used_cre.add(best_idx)
            self._log(d_idx,    code, "transfer", "transfer", 0.80, df)
            self._log(best_idx, code, "transfer", "transfer", 0.80, df)
            self._stats["transfer"] += 2

    # ------------------------------------------------------------------ #
    # Rule 4: Residual completion
    # ------------------------------------------------------------------ #

    def _rule_residual_completion(self, df: pd.DataFrame, cc: str) -> None:
        """Complète les groupes déjà lettrés avec un résiduel exact non lettré."""
        lettered = df[df[cc].ne("")]
        if lettered.empty:
            return

        net_by_code = lettered.groupby(cc).apply(
            lambda g: round(float(g[TMP + "debit"].sum() - g[TMP + "credit"].sum()), 2)
        )
        imbalanced = net_by_code[net_by_code.abs() > self.tolerance]
        if imbalanced.empty:
            return

        unlet_mask = df[TMP + "valid"] & ~df[TMP + "exclu"] & df[cc].eq("")
        unlettered = df[unlet_mask].copy()

        for code, net in imbalanced.items():
            residual = abs(round(net, 2))
            if net > 0:
                cand = unlettered[
                    abs(unlettered[TMP + "credit"].round(2) - residual) <= self.tolerance
                ]
            else:
                cand = unlettered[
                    abs(unlettered[TMP + "debit"].round(2) - residual) <= self.tolerance
                ]
            if cand.empty:
                continue

            grp = lettered[lettered[cc] == code]
            comptes = set(grp[TMP + "compte"].dropna().astype(str))
            classes = set(grp[TMP + "classe"].dropna().astype(str))

            same_cpt = cand[cand[TMP + "compte"].astype(str).isin(comptes)]
            if not same_cpt.empty:
                pick, confidence = same_cpt.index[0], 0.88
            else:
                same_cls = cand[cand[TMP + "classe"].astype(str).isin(classes)]
                if not same_cls.empty:
                    pick, confidence = same_cls.index[0], 0.70
                else:
                    if self.strict_mode:
                        continue
                    pick, confidence = cand.index[0], 0.55

            df.loc[pick, cc] = code
            # Retirer de la vue locale pour éviter le double usage
            unlettered = unlettered.drop(pick, errors="ignore")
            self._log(pick, code, "residual_completion", "residual", confidence, df)

    # ------------------------------------------------------------------ #
    # Rule 5-7: Per-account matching (1-1 vectorized, N-1 / 1-N DP)
    # ------------------------------------------------------------------ #

    def _match_account(self, df: pd.DataFrame, cc: str, cpt: str) -> None:
        """Lettrage 1-1, N-1 et 1-N pour un compte donné."""
        base_mask = (
            (df[TMP + "compte"] == cpt)
            & df[TMP + "valid"]
            & ~df[TMP + "exclu"]
            & ~df[TMP + "already_lettered"]
            & df[cc].eq("")
        )
        grp = df[base_mask]
        if grp.empty:
            return

        deb = grp[grp[TMP + "debit"] > 0].copy()
        cre = grp[grp[TMP + "credit"] > 0].copy()
        if deb.empty or cre.empty:
            return

        used_d: set = set()
        used_c: set = set()

        # --- 1-1 : merge vectorisé sur montant arrondi ---
        deb["_r"] = deb[TMP + "debit"].round(2)
        cre["_r"] = cre[TMP + "credit"].round(2)

        merged = (
            deb[["_r"]].reset_index()
            .merge(cre[["_r"]].reset_index(), on="_r", suffixes=("_d", "_c"))
        )
        for _, row in merged.iterrows():
            d_idx = row["index_d"]
            c_idx = row["index_c"]
            if d_idx in used_d or c_idx in used_c:
                continue
            # Vérification de tolérance sur valeurs d'origine (pas arrondi seul)
            if abs(float(df.at[d_idx, TMP + "debit"]) - float(df.at[c_idx, TMP + "credit"])) > self.tolerance:
                continue
            code = self.code_gen.next()
            df.loc[d_idx, cc] = code
            df.loc[c_idx, cc] = code
            used_d.add(d_idx)
            used_c.add(c_idx)
            self._log(d_idx, code, "1-1", "1-1", 0.95, df)
            self._log(c_idx, code, "1-1", "1-1", 0.95, df)
            self._stats["1-1"] += 1

        rem_deb = deb.drop(list(used_d), errors="ignore")
        rem_cre = cre.drop(list(used_c), errors="ignore")

        # --- N-1 ---
        for c_idx in list(rem_cre.index):
            if c_idx not in rem_cre.index or rem_deb.empty:
                continue
            credit_val = float(rem_cre.at[c_idx, TMP + "credit"])
            combo = find_combo(
                rem_deb[TMP + "debit"].tolist(),
                credit_val,
                tolerance=self.tolerance,
                max_items=self.max_combo,
            )
            if combo:
                sel = [rem_deb.index[i] for i in combo]
                code = self.code_gen.next()
                for x in sel + [c_idx]:
                    df.loc[x, cc] = code
                    self._log(x, code, "N-1", "N-1", 0.88, df)
                self._stats["N-1"] += 1
                rem_deb = rem_deb.drop(sel, errors="ignore")
                rem_cre = rem_cre.drop([c_idx], errors="ignore")

        # --- 1-N ---
        for d_idx in list(rem_deb.index):
            if d_idx not in rem_deb.index or rem_cre.empty:
                continue
            debit_val = float(rem_deb.at[d_idx, TMP + "debit"])
            combo = find_combo(
                rem_cre[TMP + "credit"].tolist(),
                debit_val,
                tolerance=self.tolerance,
                max_items=self.max_combo,
            )
            if combo:
                sel = [rem_cre.index[i] for i in combo]
                code = self.code_gen.next()
                for x in [d_idx] + sel:
                    df.loc[x, cc] = code
                    self._log(x, code, "1-N", "1-N", 0.88, df)
                self._stats["1-N"] += 1
                rem_cre = rem_cre.drop(sel, errors="ignore")

    # ------------------------------------------------------------------ #
    # Trace logging
    # ------------------------------------------------------------------ #

    def _log(
        self,
        idx: int,
        code: str,
        rule: str,
        category: str,
        confidence: float,
        df: pd.DataFrame,
    ) -> None:
        m = self.mapping
        entry: dict = {
            "ligne":       int(idx),
            "code_lettre": code,
            "regle":       rule,
            "categorie":   category,
            "confiance":   round(confidence, 2),
        }
        for key in ("compte", "debit", "credit"):
            col = m.get(key)
            if col and col in df.columns and idx in df.index:
                entry[key] = df.at[idx, col]
        lib_col = m.get("libelle")
        if lib_col and lib_col in df.columns and idx in df.index:
            raw = str(df.at[idx, lib_col])
            entry["libelle"] = raw[:80] if raw not in ("nan", "None") else ""
        date_col = m.get("date")
        if date_col and date_col in df.columns and idx in df.index:
            entry["date"] = df.at[idx, date_col]
        self._trace.append(entry)

    def get_trace_df(self) -> pd.DataFrame:
        if not self._trace:
            return pd.DataFrame(
                columns=["ligne", "code_lettre", "regle", "categorie", "confiance"]
            )
        return pd.DataFrame(self._trace)

    def get_stats(self) -> dict:
        return dict(self._stats)


# ============================================================================
# PUBLIC FACADE — backward-compatible API
# ============================================================================

class SimpleLettrageEngine:
    """
    Façade publique — préserve l'API historique tout en déléguant aux composants
    modulaires.  Le tuple de retour de `run()` est identique à l'original:
        (df, stats, total_lettered, eligible, todo)

    Pour accéder au résultat complet (trace, non lettrés), utiliser `run_full()`.
    """

    def __init__(self, df: pd.DataFrame, mapping: dict):
        self.df      = df
        self.mapping = dict(mapping)   # copie défensive — jamais modifié en place

    # Rétro-compatibilité
    def run(self, progress_callback: Optional[Callable] = None) -> tuple:
        """Retourne (df, stats, total_lettered, eligible_lines, todo_lines)."""
        result = self.run_full(progress_callback)
        return (
            result.df,
            result.stats,
            result.total_lettered,
            result.eligible_lines,
            result.todo_lines,
        )

    def run_full(self, progress_callback: Optional[Callable] = None) -> LettrageResult:
        """Résultat complet avec trace d'audit et lignes non lettrées."""
        m         = self.mapping
        tolerance = float(m.get("tolerance", 0.05))
        code_col  = m.get("code_lettre") or "Code_lettre"

        # Validation du mapping
        errors = validate_mapping(m, list(self.df.columns))
        if errors:
            raise ValueError("Mapping invalide:\n" + "\n".join(errors))

        # Préparation
        prep = DataPreparator(mapping=m, tolerance=tolerance)
        df   = prep.prepare(self.df)

        # Décompte initial
        elig_mask = df[TMP + "valid"] & ~df[TMP + "exclu"]
        eligible  = int(elig_mask.sum())
        todo      = int((elig_mask & ~df[TMP + "already_lettered"]).sum())

        # Générateur de codes
        if m.get("continuer_codes", False):
            code_gen = CodeGenerator.from_existing_codes(df[code_col])
        else:
            code_gen = CodeGenerator(start=1)

        # Exécution
        engine = MatchingEngine(mapping=m, code_gen=code_gen, tolerance=tolerance)
        df     = engine.run(df, progress_callback=progress_callback)

        # Nouvelles lignes lettrées durant ce cycle
        newly_lettered = int(
            (elig_mask & df[code_col].ne("") & ~df[TMP + "already_lettered"]).sum()
        )

        # Lignes non lettrées (éligibles sans code)
        unmatched_mask = elig_mask & df[code_col].eq("")
        unmatched = df[unmatched_mask].drop(
            columns=[c for c in df.columns if c.startswith(TMP)], errors="ignore"
        )

        # Nettoyage colonnes temporaires
        temp_cols = [c for c in df.columns if c.startswith(TMP)]
        df_clean  = df.drop(columns=temp_cols, errors="ignore")

        return LettrageResult(
            df=df_clean,
            stats=engine.get_stats(),
            total_lettered=newly_lettered,
            eligible_lines=eligible,
            todo_lines=todo,
            trace=engine.get_trace_df(),
            unmatched=unmatched,
        )


# ============================================================================
# ANALYSE DES COMPTES
# ============================================================================

def analyse_comptes(
    df: pd.DataFrame,
    compte_col: str,
    debit_col: str,
    credit_col: str,
) -> pd.DataFrame:
    work = df.copy()
    work["_compte_num"] = work[compte_col].astype(str).str.extract(r"(\d+)")[0]
    work["_debit_n"]    = pd.to_numeric(work[debit_col],  errors="coerce").fillna(0)
    work["_credit_n"]   = pd.to_numeric(work[credit_col], errors="coerce").fillna(0)

    summary = (
        work.groupby("_compte_num")
        .agg(
            total_debit=("_debit_n",  "sum"),
            total_credit=("_credit_n", "sum"),
            lignes=("_compte_num",     "count"),
        )
        .reset_index()
        .rename(columns={"_compte_num": "compte"})
    )
    summary["solde"] = summary["total_debit"] - summary["total_credit"]
    summary["etat"]  = summary["solde"].apply(
        lambda x: "Soldé" if abs(x) < 0.01
        else ("Reste à payer" if x > 0 else "Reste à encaisser")
    )
    return summary


def export_analyse(summary: pd.DataFrame, folder) -> dict:
    out = Path(folder)
    out.mkdir(parents=True, exist_ok=True)

    paths: dict[str, str] = {}
    for fname, df_out in [
        ("analyse_detaillee.xlsx",  summary),
        ("comptes_solde.xlsx",      summary[summary["etat"] == "Soldé"]),
        ("comptes_non_solde.xlsx",  summary[summary["etat"] != "Soldé"]),
    ]:
        p = out / fname
        df_out.to_excel(p, index=False)
        paths[fname] = str(p)

    return paths


# ============================================================================
# COLUMN AUTO-DETECTION
# ============================================================================

def auto_detect_columns(df: pd.DataFrame) -> tuple:
    """Détecte automatiquement les colonnes compte, débit, crédit, journal."""
    cols_lower = {c.lower().strip(): c for c in df.columns}

    def _find(names: list) -> Optional[str]:
        for name in names:
            if name in cols_lower:
                return cols_lower[name]
        return None

    compte = _find([
        "compte", "comptes", "n° compte", "num compte", "numero compte",
        "code compte", "compte gl", "compte auxiliaire", "n°compte",
    ])
    debit = _find([
        "debit", "débit", "dt", "montant d", "montant débit",
        "mouvement débiteur", "débiteur", "d",
    ])
    credit = _find([
        "credit", "crédit", "ct", "montant c", "montant crédit",
        "mouvement créditeur", "créditeur", "c",
    ])
    journal = _find([
        "journal", "code journal", "jnl", "journal code",
    ])

    return compte, debit, credit, journal
