# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller Spec-Datei – DRK Dienstkleidung App
# Erstellen mit: pyinstaller DRKDienstkleidung.spec
# Das EXE landet in: App/dist/DRKDienstkleidung/
#
# WICHTIG: Die EXE muss im App/-Ordner ausgeführt werden oder
# mit --distpath . direkt nach App/ gebaut werden.

import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Alle PySide6-Plugins (Fonts, Plattform etc.) einbinden
pyside6_datas = collect_data_files('PySide6')

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        # PySide6 Daten (Plugins, Qt-Translations etc.)
        *pyside6_datas,
        # python-docx Vorlagendaten
        *collect_data_files('docx'),
    ],
    hiddenimports=[
        # PySide6 Module
        'PySide6.QtCore',
        'PySide6.QtGui',
        'PySide6.QtWidgets',
        'PySide6.QtPrintSupport',
        # Datenbankmodule
        'sqlite3',
        # Dokument-Bibliotheken
        'docx',
        'docx.oxml',
        'docx.oxml.ns',
        'docx.shared',
        'docx.enum.text',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        # App-Module
        'database',
        'config',
        'utils',
        'styles',
        'pathutils',
        'main_window',
        'modules.dashboard',
        'modules.bestand',
        'modules.ausgabe',
        'modules.mitarbeiter',
        'modules.verlauf',
        'modules.bestellung',
        'modules.word_protokoll',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'numpy',
        'scipy',
        'PIL',
        'IPython',
        'jupyter',
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
    name='DRKDienstkleidung',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # kein Konsolenfenster
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
