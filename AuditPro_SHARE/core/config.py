"""
Configuration globale de l'application AuditPro.
"""
import os
from pathlib import Path

# ── Chemins ──────────────────────────────────────────────
APP_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = APP_DIR / "data"
RESOURCES_DIR = APP_DIR / "resources"
MODULES_DIR = APP_DIR / "modules"

PROFILES_FILE = DATA_DIR / "profiles.json"
HISTORY_FILE = DATA_DIR / "history.json"

# ── Application ──────────────────────────────────────────
APP_NAME = "AuditPro"
APP_VERSION = "1.0.0"
APP_SUBTITLE = "Assistant d'audit intelligent"
ORGANISATION = ""

# ── Paramètres UI ────────────────────────────────────────
MAX_HISTORY = 20
PREVIEW_ROWS = 10
SUPPORTED_EXTENSIONS = [".xlsx", ".xls", ".xlsm", ".csv"]

# ── Fonctions optionnelles ───────────────────────────────
# Désactivé par défaut pour garder une UI légère.
ENABLE_AI = os.environ.get("AUDITPRO_ENABLE_AI", "0").strip().lower() in {"1", "true", "yes", "on"}

# ── Couleurs thème (baseline Figma enterprise) ───────
COLORS = {
    "primary":      "#B882EE",
    "primary_dark":  "#C86FD0",
    "primary_light": "#F2A5F2",
    "accent":        "#FDF4FF",
    "success":       "#10B981",
    "warning":       "#F59E0B",
    "danger":        "#EF4444",
    "bg_main":       "#FAFBFC",
    "bg_panel":      "#FFFFFF",
    "bg_sidebar":    "#FFFFFF",
    "text_primary":  "#1F2937",
    "text_secondary":"#6B7280",
    "text_on_dark":  "#1F2937",
    "border":        "#E5E7EB",
}

# ── Création auto des dossiers ───────────────────────────
DATA_DIR.mkdir(exist_ok=True)
