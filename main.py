import sys
import os
import ctypes
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QFontDatabase, QFont, QPalette, QColor
from main_window import MainWindow

# --- Windows에서 작업 표시줄 아이콘을 명확하게 설정하기 위한 코드 ---
# AppUserModelID를 설정하여 Windows가 앱을 고유하게 식별하도록 함
myappid = 'my-org.excel-updater.2-0' # arbitrary string
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
# ---------------------------------------------------------

def resource_path(relative_path):
    """ 리소스에 대한 절대 경로를 반환합니다. (개발 및 PyInstaller 환경 모두에서 작동) """ 
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller는 임시 폴더를 만들고 _MEIPASS에 경로를 저장합니다.
        base_path = sys._MEIPASS
    else:
        # 개발 환경에서는 스크립트 파일이 위치한 폴더를 기준으로 합니다.
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)


def apply_light_palette(app: QApplication):
    """Force a light palette so Windows dark mode does not bleed into the app."""
    app.setStyle("Fusion")
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(248, 249, 250))
    palette.setColor(QPalette.WindowText, QColor(33, 37, 41))
    palette.setColor(QPalette.Base, QColor(255, 255, 255))
    palette.setColor(QPalette.AlternateBase, QColor(240, 240, 240))
    palette.setColor(QPalette.Text, QColor(33, 37, 41))
    palette.setColor(QPalette.Button, QColor(248, 249, 250))
    palette.setColor(QPalette.ButtonText, QColor(33, 37, 41))
    palette.setColor(QPalette.Highlight, QColor(51, 102, 204))
    palette.setColor(QPalette.HighlightedText, QColor(255, 255, 255))
    app.setPalette(palette)

# 이 프로그램의 유일한 시작점
if __name__ == '__main__':
    app = QApplication(sys.argv)
    apply_light_palette(app)
    icon_path = resource_path("icon.ico")
    window = MainWindow(icon_path=icon_path)
    window.show()
    sys.exit(app.exec())
