# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['doc2md_windows_drop.py'],
    pathex=['C:\\Users\\U743539\\Downloads\\doc2md_asciidoc_pandoc'],
    binaries=[],
    datas=[],
    hiddenimports=['doc2md_asciidoc_pandoc', 'fitz', 'pymupdf', 'pymupdf4llm', 'pytesseract', 'PIL', 'PIL.Image', 'PIL.ImageEnhance', 'pandas', 'openpyxl', 'xlrd', 'mammoth', 'markdownify', 'tabulate'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['IPython', 'jupyter', 'notebook', 'matplotlib', 'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'scipy', 'dask', 'distributed', 'pyarrow', 'numba', 'llvmlite', 'pygame', 'pytest', 'pandas.tests', 'pymupdf.layout', 'torch', 'tensorflow', 'keras', 'onnxruntime'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='doc2md-drop',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
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
    name='doc2md-drop',
)
