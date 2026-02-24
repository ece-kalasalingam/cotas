from PySide6.QtWidgets import QLabel, QGraphicsOpacityEffect, QWidget
from PySide6.QtCore import Qt, QPropertyAnimation, QTimer, QPoint

class ToastNotification(QLabel):
    def __init__(self, parent, message, type="info", duration=6000):
        super().__init__(parent)
        
        colors = {
            "success": "#2ecc71",
            "warning": "#f39c12",
            "error": "#e74c3c",
            "info": "#3498db"
        }
        
        bg_color = colors.get(type, "#333333")
        self.setText(message)
        self.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.setFixedSize(380, 55) 
        
        # Sharp corners (no border-radius)
        self.setStyleSheet(f"""
            QLabel {{
                background-color: {bg_color};
                color: white;
                border: none;
                font-size: 11pt;
                font-family: 'Segoe UI', sans-serif;
                padding-left: 15px;
                padding-right: 15px;
            }}
        """)

        self.opacity_effect = QGraphicsOpacityEffect(self)
        self.setGraphicsEffect(self.opacity_effect)
        
        self.update_position()
        self.fade_in()
        QTimer.singleShot(duration, self.fade_out)

    def update_position(self):
        p = self.parent()
        if isinstance(p, QWidget):
            p_rect = p.rect()
            margin = 20
            
            # X = Total Width - Toast Width - Margin
            x = p_rect.width() - self.width() - margin
            
            # Y = Total Height - Toast Height - Margin
            y = p_rect.height() - self.height() - margin
            
            self.move(QPoint(x, y))

    def fade_in(self):
        self.anim = QPropertyAnimation(self.opacity_effect, b"opacity")
        self.anim.setDuration(300)
        self.anim.setStartValue(0)
        self.anim.setEndValue(1)
        self.anim.start()
        self.show()

    def fade_out(self):
        self.anim = QPropertyAnimation(self.opacity_effect, b"opacity")
        self.anim.setDuration(800)
        self.anim.setStartValue(1)
        self.anim.setEndValue(0)
        self.anim.finished.connect(self.deleteLater)
        self.anim.start()