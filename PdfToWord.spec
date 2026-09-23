# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

a = Analysis(
    ['gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('pdf_to_word_converter.py', '.'),
        ('icon.ico', '.'),
        ('icon.png', '.'),
    ] + collect_data_files('rapidocr_onnxruntime')
      + collect_data_files('tkinterdnd2'),
    hiddenimports=['pdf2docx', 'fitz', 'docx', 'cv2',
                   'rapidocr_onnxruntime', 'tkinterdnd2']
        + collect_submodules('rapidocr_onnxruntime')
        + collect_submodules('tkinterdnd2'),
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PdfToWord',
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
    icon='icon.ico',
    version='version.txt',
)
