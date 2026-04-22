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

# ── Couleurs thème (violet foncé) ─────────────────────
COLORS = {
    "primary":      "#4B286D",
    "primary_dark":  "#3A1F54",
    "primary_light": "#6B4D8A",
    "accent":        "#E8336D",   # Rose accent
    "success":       "#28A745",
    "warning":       "#FFC107",
    "danger":        "#DC3545",
    "bg_main":       "#F5F5F7",
    "bg_panel":      "#FFFFFF",
    "bg_sidebar":    "#2D2040",
    "text_primary":  "#1A1A2E",
    "text_secondary":"#6C757D",
    "text_on_dark":  "#EEEAF3",
    "border":        "#DEE2E6",
}

# ── Création auto des dossiers ───────────────────────────
DATA_DIR.mkdir(exist_ok=True)
