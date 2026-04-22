# -*- mode: python ; coding: utf-8 -*-
"""
Spec PyInstaller pour AuditPro v1.1
Build : pyinstaller AuditPro.spec
"""
from pathlib import Path
import sys

ROOT = Path(SPECPATH)

block_cipher = None

a = Analysis(
    [str(ROOT / 'main.py')],
    pathex=[str(ROOT)],
    binaries=[],
    datas=[
        # Ressources visuelles
        (str(ROOT / 'resources'), 'resources'),
        # Tous les modules avec leurs scripts
        (str(ROOT / 'modules' / 'tva'),                  'modules/tva'),
        (str(ROOT / 'modules' / 'cnss'),                 'modules/cnss'),
        (str(ROOT / 'modules' / 'lettrage'),             'modules/lettrage'),
        (str(ROOT / 'modules' / 'retraitement'),         'modules/retraitement'),
        (str(ROOT / 'modules' / 'circularisation'),      'modules/circularisation'),
        (str(ROOT / 'modules' / 'srm_generator'),        'modules/srm_generator'),
        (str(ROOT / 'modules' / 'extraction_factures'),  'modules/extraction_factures'),
        (str(ROOT / 'modules' / 'extraction_ir'),        'modules/extraction_ir'),
        # Core
        (str(ROOT / 'core'),   'core'),
        (str(ROOT / 'ui'),     'ui'),
    ],
    hiddenimports=[
        # PyQt6
        'PyQt6', 'PyQt6.QtWidgets', 'PyQt6.QtCore', 'PyQt6.QtGui',
        'PyQt6.sip',
        # Data
        'pandas', 'numpy', 'openpyxl', 'openpyxl.styles',
        'openpyxl.utils', 'openpyxl.chart',
        # PDF / OCR
        'pdfplumber', 'pdfminer', 'pdfminer.high_level',
        'pdf2image', 'pytesseract',
        'pypdf', 'pypdf._reader',
        # Word
        'docx', 'docx.shared', 'docx.enum.text',
        # Email
        'email', 'email.mime', 'email.mime.multipart',
        'email.mime.text', 'email.mime.application',
        # Stdlib
        'importlib', 'importlib.util', 'pathlib',
        'json', 'datetime', 'io', 'contextlib',
        # Modules AuditPro
        'modules', 'modules.base_module',
        'modules.tva', 'modules.tva.module',
        'modules.cnss', 'modules.cnss.module',
        'modules.lettrage', 'modules.lettrage.module',
        'modules.lettrage.lettrage_engine',
        'modules.retraitement', 'modules.retraitement.module',
        'modules.circularisation', 'modules.circularisation.module',
        'modules.srm_generator', 'modules.srm_generator.module',
        'modules.extraction_factures', 'modules.extraction_factures.module',
        'modules.extraction_ir', 'modules.extraction_ir.module',
        'core', 'core.config', 'core.module_registry',
        'core.file_detector', 'core.profiles', 'core.history', 'core.worker',
        'ui', 'ui.main_window', 'ui.workspace', 'ui.module_panel',
        'ui.assistant_panel', 'ui.preview_table', 'ui.profile_dialog',
        'ui.styles',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter', 'matplotlib', 'scipy', 'sklearn',
        'IPython', 'jupyter', 'notebook', 'pytest',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='AuditPro',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,           # Pas de fenêtre console
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(ROOT / 'resources' / 'AuditPro.ico'),
    version_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='AuditPro',
)
