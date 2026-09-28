# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


PROJECT_ROOT = Path(SPECPATH)
DATA_FILES = [
    (str(path), ".")
    for path in PROJECT_ROOT.glob("app_icon*.png")
]
DATA_FILES.append((str(PROJECT_ROOT / "app_icon.ico"), "."))

a = Analysis(
    [str(PROJECT_ROOT / "main.py")],
    pathex=[str(PROJECT_ROOT)],
    binaries=[],
    datas=DATA_FILES,
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
    name="Pokazaniya",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    icon=str(PROJECT_ROOT / "app_icon.ico"),
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="Pokazaniya",
)
