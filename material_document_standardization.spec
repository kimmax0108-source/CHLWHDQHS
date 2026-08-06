# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
from PyInstaller.utils.hooks import collect_submodules

ROOT = Path(SPECPATH)
ASSET_FILES = [
    (str(path), "assets")
    for path in (ROOT / "assets").iterdir()
    if path.is_file()
]

hiddenimports = (
    collect_submodules("material_document_app")
    + collect_submodules("purchase_request_app")
    + collect_submodules("expense_statement_app")
    + collect_submodules("material_claim_manager")
)

a = Analysis(
    [str(ROOT / "main.py")],
    pathex=[str(ROOT / "src")],
    binaries=[],
    datas=[
        (str(ROOT / "templates" / "purchase_request_3sheet_template.xlsx"), "templates"),
        (str(ROOT / "templates" / "expense_statement_template.xlsx"), "templates"),
        (str(ROOT / "presets" / "default_presets.json"), "presets"),
        (str(ROOT / "presets" / "expense_default_project.json"), "presets"),
    ] + ASSET_FILES,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

a.datas = [entry for entry in a.datas if "translations" not in entry[0].lower()]
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="material_document_standardization",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(ROOT / "assets" / "app.ico"),
)
