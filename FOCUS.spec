# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all

_qt_datas, _qt_binaries, _qt_hiddenimports = [], [], []
for _mod in [
    'PySide6.QtCore',
    'PySide6.QtGui',
    'PySide6.QtWidgets',
    'PySide6.QtNetwork',
    'PySide6.QtSvg',
    'PySide6.QtPdf',
    'PySide6.QtPdfWidgets',
]:
    _d, _b, _h = collect_all(_mod)
    _qt_datas += _d
    _qt_binaries += _b
    _qt_hiddenimports += _h

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=_qt_binaries,
    datas=[('assets', 'assets'), ('common/i18n', 'common/i18n')] + _qt_datas,
    hiddenimports=['modules.instructor_module', 'modules.co_analysis_module', 'modules.po_analysis_module', 'modules.help_module', 'modules.about_module'] + _qt_hiddenimports,
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
    upx=False,
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
    upx=False,
    upx_exclude=[],
    name='focus',
)
