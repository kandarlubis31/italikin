# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for Italikin
Build command: pyinstaller italikin.spec
"""

block_cipher = None

a = Analysis(
    ['italikin.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'customtkinter',
        'docx',
        'docx.document',
        'docx.text',
        'docx.text.run',
        'docx.enum.text',
        'docx.shared',
        'PIL',
        'PIL._tkinter_finder',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter.test',
        'unittest',
        'pytest',
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
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
    [],
    exclude_binaries=True,
    name='Italikin',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # No console window for GUI app
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add icon file here: icon='assets/icon.ico'
    version='file_version_info.txt',  # Version info for Windows
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Italikin',
)
