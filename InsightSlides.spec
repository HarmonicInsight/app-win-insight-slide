# -*- mode: python ; coding: utf-8 -*-
"""
InsightSlides PyInstaller spec (recommended: onedir, GUI)
- Fixes TOC normalization error by integrating collect_data_files('pptx') via Analysis(datas=...)
- Designed for stable packaging first; switch to onefile after runtime verification.
"""

from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# Collect package data for python-pptx (package name: pptx)
pptx_datas = collect_data_files('pptx')

a = Analysis(
    ['InsightSlides.py'],
    pathex=[],
    binaries=[],
    datas=pptx_datas,
    hiddenimports=[
        # python-pptx common paths
        'pptx',
        'pptx.util',
        'pptx.dml.color',
        'pptx.enum.shapes',
        'pptx.enum.text',

        # your stack
        'openpyxl',
        'openpyxl.styles',
        'PIL',
        'PIL._tkinter_finder',
        'tksheet',
        'tkinter',
        'tkinter.ttk',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # keep exe smaller (remove if you actually use them)
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'pytest',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
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
    name='InsightSlides',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI app
    disable_windowed_traceback=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='InsightSlides',
)
