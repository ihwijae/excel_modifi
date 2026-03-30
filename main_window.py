import sys
from PySide6.QtWidgets import QMainWindow, QTabWidget, QMessageBox, QWidget
from PySide6.QtGui import QIcon
import easyocr

from business_status_tab import BusinessStatusTab
from credit_rating_tab import CreditRatingTab
from smpp_dialog import SmppLookupDialog


class MainWindow(QMainWindow):
    def __init__(self, icon_path=None):
        super().__init__()
        self.setWindowTitle("엑셀업데이트 도구 v2.0 (베타)")
        self.setGeometry(100, 100, 1220, 820)
        if icon_path:
            self.setWindowIcon(QIcon(icon_path))

        try:
            self.reader = easyocr.Reader(["ko", "en"], gpu=False)
        except Exception as exc:  # pylint: disable=broad-except
            QMessageBox.critical(self, "EasyOCR 초기화 실패", f"EasyOCR 로딩 중 오류가 발생했습니다:\n{exc}")
            sys.exit()

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        self.business_tab = BusinessStatusTab(self.reader)
        self.tabs.addTab(self.business_tab, "1. 경영상태")

        self.credit_tab = CreditRatingTab(self.reader)
        self.tabs.addTab(self.credit_tab, "2. 신용평가 데이터")

        self.smpp_placeholder = QWidget()
        self.smpp_tab_index = self.tabs.addTab(self.smpp_placeholder, "3. 중소/여성기업 조회")
        self.last_tab_index = 0
        self.tabs.currentChanged.connect(self.on_tab_changed)

    def on_tab_changed(self, index: int):
        if index == self.smpp_tab_index:
            self.tabs.blockSignals(True)
            self.tabs.setCurrentIndex(self.last_tab_index)
            self.tabs.blockSignals(False)
            self.open_smpp_dialog()
        else:
            if self.tabs.widget(index) is self.credit_tab:
                self.credit_tab.on_tab_activated()
            self.last_tab_index = index

    def open_smpp_dialog(self):
        excel_paths = self.business_tab.get_excel_paths()
        default_excel = self.business_tab.get_selected_excel_path()
        dialog = SmppLookupDialog(
            excel_paths=excel_paths,
            default_excel_path=default_excel,
            parent=self,
        )
        dialog.exec()
