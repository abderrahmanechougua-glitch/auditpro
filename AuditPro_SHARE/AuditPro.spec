# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['c:\\Users\\Abderrahmane.CHOUGUA\\Downloads\\AuditPro_v1.1_1\\AuditPro_SHARE\\main.py'],
    pathex=[],
    binaries=[],
    datas=[('c:\\Users\\Abderrahmane.CHOUGUA\\Downloads\\AuditPro_v1.1_1\\AuditPro_SHARE\\resources', 'resources'), ('c:\\Users\\Abderrahmane.CHOUGUA\\Downloads\\AuditPro_v1.1_1\\AuditPro_SHARE\\modules', 'modules'), ('c:\\Users\\Abderrahmane.CHOUGUA\\Downloads\\AuditPro_v1.1_1\\AuditPro_SHARE\\core', 'core'), ('c:\\Users\\Abderrahmane.CHOUGUA\\Downloads\\AuditPro_v1.1_1\\AuditPro_SHARE\\ui', 'ui'), ('c:\\Users\\Abderrahmane.CHOUGUA\\Downloads\\AuditPro_v1.1_1\\AuditPro_SHARE\\vendor', 'vendor')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

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
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['c:\\Users\\Abderrahmane.CHOUGUA\\Downloads\\AuditPro_v1.1_1\\AuditPro_SHARE\\resources\\AuditPro.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='AuditPro',
)
