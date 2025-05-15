# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['marshrutka.py'],
    pathex=[],
    binaries=[],
    datas=[('specialists.json', '.'), ('plavka.xlsx', '.'), ('marshrutka.xlsx', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    module_collection_mode={'PySide6': 'pyz+py'},
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Электронная_маршрутная_карта',
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
    icon=['images\\app.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Электронная_маршрутная_карта',
)
