# my_app.spec

import sys
from os import path
site_packages = next(p for p in sys.path if 'site-packages' in p)

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[('CM_logo.png', '.'),
        (path.join(site_packages,"docx","parts"), 
                "docx/parts"),
        (path.join(site_packages,"docx","templates"), 
                "docx/templates"),
        (path.join(site_packages,"reportlab"), 
                "reportlab"),
        (path.join(site_packages, "pymupdf"),
                "pymupdf")],
    hiddenimports=['docx'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='pdf_to_word_converter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # Set to False to hide the console window
    # icon='app_icon.ico'  # Optional: Add an icon file for the executable
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='pdf_to_word_converter'
)
