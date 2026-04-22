"""
ocr_paths.py — Détection automatique des chemins Tesseract et Poppler.

Recherche dans l'ordre :
  1. Dossier bundlé  : <racine_app>/vendor/tesseract/  et  <racine_app>/vendor/poppler/bin/
  2. Variable d'env  : TESSERACT_PATH  /  POPPLER_PATH
  3. Emplacements    Windows courants
  4. PATH système    (shutil.which)

Si Tesseract est introuvable, retourne None — l'app doit travailler en mode
pdfplumber uniquement (PDF natifs) et ignorer l'OCR sans crasher.
"""

import glob
import os
import shutil
import sys
from pathlib import Path

# ------------------------------------------------------------------ #
# Racine de l'application  (robust whether frozen via PyInstaller or not)
# ------------------------------------------------------------------ #
if getattr(sys, "frozen", False):
    _APP_ROOT = Path(sys.executable).parent
else:
    _APP_ROOT = Path(__file__).resolve().parent.parent   # core/ → AuditPro_SHARE/


def _find_tesseract() -> str | None:
    """Retourne le chemin complet vers tesseract.exe, ou None."""

    # 1. Bundlé dans l'app
    candidates = [
        _APP_ROOT / "vendor" / "tesseract" / "tesseract.exe",
        _APP_ROOT / "tesseract" / "tesseract.exe",
    ]
    for c in candidates:
        if c.exists():
            return str(c)

    # 2. Variable d'environnement
    env_val = os.environ.get("TESSERACT_PATH", "")
    if env_val and Path(env_val).exists():
        return env_val

    # 3. Emplacements Windows standard
    standard = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    for p in standard:
        if Path(p).exists():
            return p

    # 4. PATH système
    which = shutil.which("tesseract")
    if which:
        return which

    return None


def _find_poppler() -> str | None:
    """Retourne le dossier bin/ de Poppler, ou None."""

    # 1. Bundlé dans l'app
    candidates = [
        _APP_ROOT / "vendor" / "poppler" / "bin",
        _APP_ROOT / "vendor" / "poppler" / "Library" / "bin",
        _APP_ROOT / "poppler" / "bin",
    ]
    for c in candidates:
        if (c / "pdftoppm.exe").exists() or (c / "pdftoppm").exists():
            return str(c)

    # 2. Variable d'environnement
    env_val = os.environ.get("POPPLER_PATH", "")
    if env_val and (Path(env_val) / "pdftoppm.exe").exists():
        return env_val

    # 3. Emplacements Windows courants (toute version)
    patterns = [
        r"C:\poppler\*\Library\bin",
        r"C:\poppler\*\bin",
        r"C:\Program Files\poppler*\bin",
    ]
    for pat in patterns:
        matches = glob.glob(pat)
        if matches:
            return matches[0]

    # 4. PATH système  (pdftoppm doit être accessible)
    which = shutil.which("pdftoppm")
    if which:
        return str(Path(which).parent)

    return None


# ------------------------------------------------------------------ #
# Chemins résolus (calculés une seule fois au chargement du module)
# ------------------------------------------------------------------ #
TESSERACT_PATH: str | None = _find_tesseract()
POPPLER_PATH:   str | None = _find_poppler()

# Ajouter le dossier Tesseract au PATH si trouvé
if TESSERACT_PATH:
    tess_dir = str(Path(TESSERACT_PATH).parent)
    path_entries = os.environ.get("PATH", "").split(os.pathsep)
    if tess_dir not in path_entries:
        os.environ["PATH"] = os.environ["PATH"] + os.pathsep + tess_dir

# ------------------------------------------------------------------ #
# Configuration automatique de pytesseract si disponible
# ------------------------------------------------------------------ #
