# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec for ExtraDataCleaner
# Built by:  build_exe.bat  (or  pyinstaller ExtraDataCleaner.spec)

import sys
import os
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs

block_cipher = None

# ── Collect tkinter TCL/TK data and DLLs (required on Windows) ───────────────
datas    = []
binaries = []

try:
    datas += collect_data_files('tkinter')
except Exception:
    pass

# TCL/TK runtime library folders (tcl8.6, tk8.6, etc.)
tcl_dir = Path(sys.base_prefix) / 'tcl'
if tcl_dir.exists():
    for item in tcl_dir.iterdir():
        if item.is_dir():
            datas.append((str(item), item.name))

# _tkinter.pyd + tcl86t.dll + tk86t.dll from the DLLs folder
dlls_dir = Path(sys.base_prefix) / 'DLLs'
if dlls_dir.exists():
    for dll in dlls_dir.glob('*.dll'):
        if dll.stem.lower().startswith(('tcl', 'tk')):
            binaries.append((str(dll), '.'))
    tkpyd = dlls_dir / '_tkinter.pyd'
    if tkpyd.exists():
        binaries.append((str(tkpyd), '.'))

# ── Bundle the icon ────────────────────────────────────────────────────────────
icon_path = Path('icon.ico')
if icon_path.exists():
    datas.append((str(icon_path), '.'))

# ── Analysis ──────────────────────────────────────────────────────────────────
a = Analysis(
    ['cleaner.py'],
    pathex=[str(Path('.').resolve())],
    binaries=binaries,
    datas=datas,
    hiddenimports=[
        'gui',
        'core',
        '_tkinter',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.font',
        'pandas',
        'pandas._libs.tslibs.base',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'numpy',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'chardet',
        'dateutil',
        'dateutil.parser',
        'et_xmlfile',
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'scipy', 'IPython', 'jupyter',
        'PIL', 'cv2', 'PyQt5', 'wx', 'gi',
        'sqlalchemy', 'psycopg2', 'boto3', 'botocore',
    ],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ExtraDataCleaner',
    debug=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # no black console window
    icon='icon.ico' if Path('icon.ico').exists() else None,
)
