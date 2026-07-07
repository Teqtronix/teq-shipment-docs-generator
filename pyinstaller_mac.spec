# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules

hiddenimports = (
    collect_submodules("tkinter")
    + collect_submodules("tkinterdnd2")
    + collect_submodules("customtkinter")
    + collect_submodules("pypdfium2")
    + collect_submodules("reportlab")
)

datas = [
    ("Logo.ico", "."),
    ("templates/Template (AUS CI).xlsx", "."),
    ("templates/Template (AUS PL).xlsx", "."),
    ("templates/Template (EU PL 50).xlsx", "."),
    ("templates/Template (EU PL 55).xlsx", "."),
    ("templates/Template (EU CI 50).xlsx", "."),
    ("templates/Template (EU CI 55).xlsx", "."),
    ("templates/Template (TJX CA CI).xlsx", "."),
    ("templates/Template (TJX CA PL).xlsx", "."),
    ("templates/TEMPLATE_LABELS_____________TJX AUS.docx", "."),
    ("templates/TEMPLATE_LABELS_____________TJX EU.docx", "."),
    ("templates/TJX CA FOB Labels template.docx", "."),
    ("templates/Template (Indigo Inner Ctn).docx", "."),
]

a = Analysis(
    ["tjx_aus_eu_processor.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
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
    name="Teq Shipment Docs Generator",
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
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="Teq Shipment Docs Generator",
)

app = BUNDLE(
    coll,
    name="Teq Shipment Docs Generator.app",
    icon=None,
    bundle_identifier="com.teqtronix.shipment-docs-generator",
)
