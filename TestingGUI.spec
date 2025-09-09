# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('resources', 'resources')],
    hiddenimports=['matplotlib.backends.backend_tkagg', 'PIL._tkinter_finder', 'pkg_resources.py2_warn', 'matplotlib.pyplot', 'matplotlib.widgets', 'matplotlib.cm', 'matplotlib.figure', 'matplotlib.patches', 'pandas.core.arrays', 'pandas.core.groupby', 'pandas.io.formats.format', 'pandas.io.common', 'pandas._libs', 'numpy.core._methods', 'numpy.lib.format'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter.test', 'test'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='TestingGUI',
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
    icon=['resources\\ccell_icon.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='TestingGUI',
)
