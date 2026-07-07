# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

py_base = Path(sys.base_prefix)
py_dlls = py_base / "DLLs"
tcl_dir = py_base / "tcl"
tcl_data_dir = tcl_dir / "tcl8.6"
tk_data_dir = tcl_dir / "tk8.6"
tcl_module_dir = tcl_dir / "tcl8"

extra_binaries = []
for dll_name in ("_tkinter.pyd", "tcl86t.dll", "tk86t.dll"):
    dll_path = py_dlls / dll_name
    if dll_path.exists():
        extra_binaries.append((str(dll_path), "."))

extra_datas = []
if tcl_data_dir.exists():
    extra_datas.append((str(tcl_data_dir), "_tcl_data"))
if tk_data_dir.exists():
    extra_datas.append((str(tk_data_dir), "_tk_data"))
if tcl_module_dir.exists():
    extra_datas.append((str(tcl_module_dir), "tcl8"))
extra_datas += collect_data_files("tkinter")
extra_datas.append((str(py_base / "Lib" / "tkinter"), "tkinter"))

tk_hiddenimports = (
    ['tkinter', '_tkinter']
    + collect_submodules("tkinter")
    + collect_submodules("pdfplumber")
    + collect_submodules("pdfminer")
    + collect_submodules("pypdfium2")
    + collect_submodules("reportlab")
)

a = Analysis(
    ['tjx_aus_eu_processor.py'],
    pathex=[],
    binaries=extra_binaries,
    datas=[
        ('Logo.ico', '.'),
        ('templates/Template (AUS CI).xlsx', '.'),
        ('templates/Template (AUS PL).xlsx', '.'),
        ('templates/Template (EU PL 50).xlsx', '.'),
        ('templates/Template (EU PL 55).xlsx', '.'),
        ('templates/Template (EU CI 50).xlsx', '.'),
        ('templates/Template (EU CI 55).xlsx', '.'),
        ('templates/Template (TJX CA CI).xlsx', '.'),
        ('templates/Template (TJX CA PL).xlsx', '.'),
        ('templates/TEMPLATE_LABELS_____________TJX AUS.docx', '.'),
        ('templates/TEMPLATE_LABELS_____________TJX EU.docx', '.'),
        ('templates/TJX CA FOB Labels template.docx', '.'),
        ('templates/Template (Indigo Inner Ctn).docx', '.'),
    ] + extra_datas,
    hiddenimports=tk_hiddenimports,
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
    name='Teq Shipment Docs Generator',
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
    icon=['Logo.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Teq Shipment Docs Generator',
)
