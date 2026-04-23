# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules

# Dynamic loading roots discovered in codebase:
# - common/module_plugins.py -> lazy_module_class("modules.*", ...)
# - domain/template_strategy_router.py -> importlib.import_module("domain.template_versions.*")
_DYNAMIC_IMPORT_ROOTS = (
    'modules',
    'domain.template_versions',
)
_dynamic_hiddenimports = []
for _root in _DYNAMIC_IMPORT_ROOTS:
    _dynamic_hiddenimports.extend(collect_submodules(_root))

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('assets', 'assets'), ('common/i18n', 'common/i18n')],
    hiddenimports=[
        'PySide6.QtCore',
        'PySide6.QtGui',
        'PySide6.QtWidgets',
        'PySide6.QtNetwork',
        'PySide6.QtSvg',
        'PySide6.QtPdf',
        'PySide6.QtPdfWidgets',
        *_dynamic_hiddenimports,
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
