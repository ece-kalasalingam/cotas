import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
import qdarktheme

from core.resources import resource_path
from main_window import MainWindow


def main() -> int:
    app = QApplication(sys.argv)

    # Theme
    qdarktheme.setup_theme("auto")

    # Icon (works in onedir build)
    app.setWindowIcon(QIcon(resource_path("assets/kare-logo.ico")))

    window = MainWindow()
    window.show()

    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())