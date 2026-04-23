# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_dynamic_libs, collect_data_files

# Qt DLLs live in site-packages/PySide6/ (155 of them).
# collect_dynamic_libs places them in the PySide6/ subdirectory of the bundle.
# The runtime hook (installer/rthook_pyside6.py) calls os.add_dll_directory()
# so Python 3.8+ can find them when loading QtCore.pyd and friends.
_qt_binaries = collect_dynamic_libs('PySide6')

# Platform and image-format plugins.  Actual path in this PySide6 layout is
# PySide6/plugins/ (not PySide6/Qt/plugins/).
_qt_datas = (
    collect_data_files('PySide6', subdir='plugins/platforms')
    + collect_data_files('PySide6', subdir='plugins/imageformats')
    + collect_data_files('PySide6', subdir='plugins/iconengines')
    + collect_data_files('PySide6', subdir='plugins/styles')
    + collect_data_files('PySide6', subdir='plugins/tls')
)

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
    runtime_hooks=['installer/rthook_pyside6.py'],
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
