# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_dynamic_libs, collect_data_files

# collect_dynamic_libs scans PySide6's __path__ (the package directory) and
# returns all Qt6*.dll files — the ones that C extension .pyd modules cannot
# self-report via collect_all on individual submodules.
# collect_data_files picks up the Qt platform/image plugins (qwindows.dll etc.)
# that Qt requires at runtime to initialise a window.
_qt_binaries = collect_dynamic_libs('PySide6')
_qt_datas    = collect_data_files('PySide6', subdir='Qt/plugins')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=_qt_binaries,
    datas=[('assets', 'assets'), ('common/i18n', 'common/i18n')] + _qt_datas,
    hiddenimports=[
        'PySide6.QtCore',
        'PySide6.QtGui',
        'PySide6.QtWidgets',
        'PySide6.QtNetwork',
        'PySide6.QtSvg',
        'PySide6.QtPdf',
        'PySide6.QtPdfWidgets',
        'modules.instructor_module',
        'modules.co_analysis_module',
        'modules.po_analysis_module',
        'modules.help_module',
        'modules.about_module',
    ],
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
