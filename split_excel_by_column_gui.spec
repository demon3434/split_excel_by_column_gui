# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from pathlib import Path

conda_bin = Path(sys.base_prefix) / 'Library' / 'bin'
if conda_bin.exists():
    os.environ['PATH'] = str(conda_bin) + os.pathsep + os.environ.get('PATH', '')

extra_binaries = []
if conda_bin.exists():
    for dll_name in [
        'tcl86t.dll',
        'tk86t.dll',
        'libcrypto-3-x64.dll',
        'libssl-3-x64.dll',
        'libexpat.dll',
        'ffi.dll',
        'liblzma.dll',
        'LIBBZ2.dll',
        'libmpdec-4.dll',
        'vcruntime140.dll',
        'vcruntime140_1.dll',
        'msvcp140.dll',
    ]:
        p = conda_bin / dll_name
        if not p.exists():
            p = Path(sys.base_prefix) / dll_name
        if p.exists():
            extra_binaries.append((str(p), '.'))

a = Analysis(
    ['split_excel_by_column_gui.py'],
    pathex=[str(conda_bin)] if conda_bin.exists() else [],
    binaries=extra_binaries,
    datas=[('logo2.png', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pandas', 'numpy', 'matplotlib', 'scipy', 'numba', 'pyarrow', 'tables', 'sqlalchemy', 'qtpy', 'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'IPython', 'jupyter', 'notebook', 'PIL', 'lxml', 'pygame', 'torch', 'tensorflow'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Excel按列拆分',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['logo_icon_hd.ico'],
)
