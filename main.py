import sys
from PySide6.QtWidgets import QApplication
from main_window import MainWindow

# 이 프로그램의 유일한 시작점
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())