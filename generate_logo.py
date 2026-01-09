from PySide6.QtGui import QPainter, QColor, QPen, QFont, QPixmap
from PySide6.QtCore import Qt, QRect

def create_logo(path):
    pixmap = QPixmap(256, 256)
    pixmap.fill(Qt.transparent)
    
    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing)
    
    # 画背景圆
    painter.setBrush(QColor("#2563eb"))
    painter.setPen(Qt.NoPen)
    painter.drawEllipse(10, 10, 236, 236)
    
    # 画字母 C
    painter.setPen(QPen(Qt.white, 30))
    font = QFont("Arial", 120, QFont.Bold)
    painter.setFont(font)
    painter.drawText(QRect(0, 0, 256, 256), Qt.AlignCenter, "C")
    
    painter.end()
    pixmap.save(path)

if __name__ == "__main__":
    import sys
    from PySide6.QtWidgets import QApplication
    app = QApplication(sys.argv)
    create_logo("ui/logo.png")
