# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['Excel按列号拆分(GUI).py'],
    pathex=[],
    binaries=[],
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
