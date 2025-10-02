import sys
import os
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QFontDatabase, QFont
from main_window import MainWindow

# 이 프로그램의 유일한 시작점
if __name__ == '__main__':
    app = QApplication(sys.argv)

    # --- Pretendard 폰트 적용 ---
    font_path = os.path.join("fonts", "Pretendard-Regular.woff2")
    if os.path.exists(font_path):
        font_id = QFontDatabase.addApplicationFont(font_path)
        if font_id != -1:
            font_family = QFontDatabase.applicationFontFamilies(font_id)[0]
            app.setFont(QFont(font_family))
        else:
            print(f"경고: 폰트 파일을 불러오는 데 실패했습니다: {font_path}")
    else:
        print(f"경고: 폰트 파일이 존재하지 않습니다: {font_path}")
    # --------------------------

    window = MainWindow()
    window.show()
    sys.exit(app.exec())