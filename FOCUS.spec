# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all

pyside6_datas, pyside6_binaries, pyside6_hiddenimports = collect_all('PySide6')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=pyside6_binaries,
    datas=[('assets', 'assets'), ('common/i18n', 'common/i18n')] + pyside6_datas,
    hiddenimports=['PySide6.QtSvg', 'PySide6.QtPdf', 'PySide6.QtPdfWidgets', 'modules.instructor_module', 'modules.co_analysis_module', 'modules.po_analysis_module', 'modules.help_module', 'modules.about_module'] + pyside6_hiddenimports,
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
    [],
    exclude_binaries=True,
    name='focus',
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
    version='version.txt',
    icon=['assets\\kare-logo.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='focus',
)
