# -*- mode: python ; coding: utf-8 -*-

# PyInstaller spec file for Standardized Testing GUI
# This creates a professional Windows executable

import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Collect all data files from resources directory
datas = []
if os.path.exists('resources'):
    for root, dirs, files in os.walk('resources'):
        for file in files:
            file_path = os.path.join(root, file)
            rel_path = os.path.relpath(file_path, '.')
            datas.append((file_path, os.path.dirname(rel_path)))

# Collect any additional data files
datas.extend([
    ('requirements.txt', '.'),
    # Add any other data files your app needs
])

# Hidden imports for packages that PyInstaller might miss
hiddenimports = [
    'sklearn.utils._cython_blas',
    'sklearn.neighbors.typedefs',
    'sklearn.neighbors.quad_tree',
    'sklearn.tree._utils',
    'openpyxl.cell._writer',
    'PIL._tkinter_finder',
    'cv2',
    'matplotlib.backends.backend_tkagg',
    'numpy.random.common',
    'numpy.random.bounded_integers',
    'numpy.random.entropy',
]

# Add any submodules that need to be included
hiddenimports.extend(collect_submodules('sklearn'))
hiddenimports.extend(collect_submodules('scipy'))

a = Analysis(
    ['main.py'],                    # Your main entry point
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'jupyter',
        'IPython',
        'notebook',
        'pytest',
        'sphinx',
        'setuptools',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='TestingGUI',                    # Your executable name
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,                        # Set to False for GUI app
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='resources/icon.ico' if os.path.exists('resources/icon.ico') else None,
    version='version_info.txt' if os.path.exists('version_info.txt') else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='TestingGUI',
)