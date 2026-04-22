#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
setup_vendor.py — Prépare les binaires Tesseract et Poppler dans vendor/
pour que l'app AuditPro soit autonome (aucune installation requise).

Usage :
    python setup_vendor.py

Stratégie :
  1. Copie depuis l'installation système si trouvée (rapide, ~30 sec)
  2. Sinon télécharge depuis GitHub (nécessite connexion internet)
"""

import os
import sys
import shutil
import glob
import zipfile
import urllib.request
import urllib.error
from pathlib import Path

# ------------------------------------------------------------------ #
APP_ROOT = Path(__file__).parent
VENDOR   = APP_ROOT / "vendor"
TESS_DST = VENDOR / "tesseract"
POPP_DST = VENDOR / "poppler" / "bin"
# ------------------------------------------------------------------ #

def log(msg, ok=True):
    icon = "✅" if ok else "❌"
    print(f"  {icon}  {msg}")

def step(msg):
    print(f"\n{'─'*55}")
    print(f"  {msg}")
    print(f"{'─'*55}")

# ================================================================== #
# TESSERACT
# ================================================================== #

TESSERACT_SYSTEM_PATHS = [
    r"C:\Program Files\Tesseract-OCR",
    r"C:\Program Files (x86)\Tesseract-OCR",
]

# Pas de zip portable officiel pour Tesseract — copie depuis le système uniquement.
# Installation manuelle si absent : https://github.com/UB-Mannheim/tesseract/wiki


def find_tesseract_system() -> Path | None:
    for p in TESSERACT_SYSTEM_PATHS:
        exe = Path(p) / "tesseract.exe"
        if exe.exists():
            return Path(p)
    which = shutil.which("tesseract")
    if which:
        return Path(which).parent
    return None


def copy_tesseract(src: Path):
    """Copie le dossier Tesseract système vers vendor/tesseract/ (copie atomique)."""
    tmp = TESS_DST.parent / "_tess_tmp"
    if tmp.exists():
        shutil.rmtree(tmp)
    print(f"     Copie de {src} → {TESS_DST}  (peut prendre 1-2 min)...")
    shutil.copytree(src, tmp, dirs_exist_ok=False)
    if TESS_DST.exists():
        shutil.rmtree(TESS_DST)
    tmp.rename(TESS_DST)
    log(f"Tesseract copié ({sum(1 for _ in TESS_DST.rglob('*'))} fichiers)")


def setup_tesseract():
    step("TESSERACT OCR")
    if (TESS_DST / "tesseract.exe").exists():
        log("Déjà présent dans vendor/tesseract/ — rien à faire")
        return True

    src = find_tesseract_system()
    if src:
        log(f"Trouvé sur le système : {src}")
        copy_tesseract(src)
        return True

    log("Tesseract introuvable sur ce système", ok=False)
    print("\n  Téléchargement automatique non supporté (pas de zip portable officiel).")
    print("  → Installez-le depuis : https://github.com/UB-Mannheim/tesseract/wiki")
    print(f"  → Puis relancez ce script (la copie se fera automatiquement)")
    return False


# ================================================================== #
# POPPLER
# ================================================================== #

POPPLER_SYSTEM_GLOBS = [
    r"C:\poppler\*\Library\bin",
    r"C:\poppler\*\bin",
    r"C:\Program Files\poppler*\bin",
    r"C:\Program Files (x86)\poppler*\bin",
]

POPPLER_ZIP_URL = (
    "https://github.com/oschwartz10612/poppler-windows/releases/download/"
    "v24.08.0-0/Release-24.08.0-0.zip"
)


def find_poppler_system() -> Path | None:
    for pattern in POPPLER_SYSTEM_GLOBS:
        matches = glob.glob(pattern)
        for m in matches:
            if (Path(m) / "pdftoppm.exe").exists():
                return Path(m)
    which = shutil.which("pdftoppm")
    if which:
        return Path(which).parent
    return None


def copy_poppler(src: Path):
    """Copie uniquement le dossier bin/ de Poppler vers vendor/poppler/bin/ (copie atomique)."""
    POPP_DST.parent.mkdir(parents=True, exist_ok=True)
    tmp = POPP_DST.parent / "_poppler_tmp_bin"
    if tmp.exists():
        shutil.rmtree(tmp)
    print(f"     Copie de {src} → {POPP_DST}...")
    shutil.copytree(src, tmp)
    if POPP_DST.exists():
        shutil.rmtree(POPP_DST)
    tmp.rename(POPP_DST)
    log(f"Poppler copié ({sum(1 for _ in POPP_DST.rglob('*'))} fichiers)")


def download_poppler():
    """Télécharge et extrait Poppler depuis GitHub."""
    zip_path = VENDOR / "_poppler_download.zip"
    VENDOR.mkdir(parents=True, exist_ok=True)

    print(f"     Téléchargement Poppler ({POPPLER_ZIP_URL})...")
    try:
        def progress(count, block_size, total):
            pct = min(int(count * block_size * 100 / total), 100)
            print(f"\r     Progression : {pct}%   ", end="", flush=True)
        urllib.request.urlretrieve(POPPLER_ZIP_URL, zip_path, reporthook=progress)
        print()
    except urllib.error.URLError as e:
        log(f"Téléchargement échoué : {e}", ok=False)
        return False

    print("     Extraction...")
    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(VENDOR / "_poppler_tmp")
    zip_path.unlink(missing_ok=True)

    # Trouver le dossier Library/bin ou bin dans l'archive extraite
    tmp = VENDOR / "_poppler_tmp"
    bin_dirs = list(tmp.rglob("pdftoppm.exe"))
    if not bin_dirs:
        log("pdftoppm.exe introuvable dans l'archive", ok=False)
        shutil.rmtree(tmp, ignore_errors=True)
        return False

    bin_dir = bin_dirs[0].parent
    copy_poppler(bin_dir)
    shutil.rmtree(tmp, ignore_errors=True)
    return True


def setup_poppler():
    step("POPPLER (pdf2image)")
    if (POPP_DST / "pdftoppm.exe").exists():
        log("Déjà présent dans vendor/poppler/bin/ — rien à faire")
        return True

    src = find_poppler_system()
    if src:
        log(f"Trouvé sur le système : {src}")
        copy_poppler(src)
        return True

    log("Poppler introuvable sur ce système — tentative de téléchargement...", ok=True)
    return download_poppler()


# ================================================================== #
# VÉRIFICATION FINALE
# ================================================================== #

def verify():
    step("VÉRIFICATION")
    tess_ok = (TESS_DST / "tesseract.exe").exists()
    popp_ok = (POPP_DST / "pdftoppm.exe").exists()

    log(f"Tesseract  : {TESS_DST / 'tesseract.exe'}", ok=tess_ok)
    log(f"Poppler    : {POPP_DST / 'pdftoppm.exe'}", ok=popp_ok)

    # Vérifier tessdata (langue française)
    tessdata = TESS_DST / "tessdata"
    fra_ok = (tessdata / "fra.traineddata").exists() if tessdata.exists() else False
    eng_ok = (tessdata / "eng.traineddata").exists() if tessdata.exists() else False
    log(f"Tessdata FR: {tessdata / 'fra.traineddata'}", ok=fra_ok)
    log(f"Tessdata EN: {tessdata / 'eng.traineddata'}", ok=eng_ok)

    if tess_ok and popp_ok and fra_ok and eng_ok:
        print("\n  🎉  vendor/ est prêt — l'app est entièrement autonome !")
        print(f"       Partagez le dossier AuditPro_SHARE complet.")
    else:
        print("\n  ⚠️  Certains composants manquent. L'OCR sera limité.")
        if not fra_ok and tess_ok:
            print("  → Tessdata manquant : copiez les fichiers *.traineddata")
            print(f"    depuis C:\\Program Files\\Tesseract-OCR\\tessdata\\")
            print(f"    vers    {tessdata}\\")


# ================================================================== #
# MAIN
# ================================================================== #

if __name__ == "__main__":
    print("\n" + "="*55)
    print("  AuditPro — Préparation des binaires OCR (vendor/)")
    print("="*55)

    VENDOR.mkdir(parents=True, exist_ok=True)

    tess_ok = setup_tesseract()
    popp_ok = setup_poppler()
    verify()
