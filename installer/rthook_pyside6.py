import os
import sys

# Python 3.8+ no longer searches subdirectories when loading DLLs for .pyd
# extension modules. Qt DLLs live in the PySide6 subfolder of the bundle, so
# they must be explicitly added to the DLL search path before any PySide6
# submodule is imported.
if sys.platform == "win32":
    _base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    for _sub in ("", "PySide6"):
        _d = os.path.join(_base, _sub)
        if os.path.isdir(_d):
            os.add_dll_directory(_d)
