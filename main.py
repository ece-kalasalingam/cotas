import os
import sys
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
import qdarktheme


from scripts.utils import resource_path
from main_window import MainWindow


def main() -> int:
    os.environ["QT_ADAPTIVE_STRUCTURE_SENSITIVITY"] = "1"
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    app = QApplication(sys.argv)
    app.setOrganizationName("COTAS")
    app.setApplicationName("COTAS")

    # Theme
    qdarktheme.setup_theme("auto")

    # Icon (works in onedir build)
    app.setWindowIcon(QIcon(resource_path("assets/kare-logo.ico")))

    window = MainWindow()
    window.show()

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())