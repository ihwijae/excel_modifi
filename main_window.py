import sys
from PySide6.QtWidgets import QMainWindow, QTabWidget, QMessageBox
import easyocr
from business_status_tab import BusinessStatusTab
from credit_rating_tab import CreditRatingTab



class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("데이터 자동 업데이트 프로그램 v2.0 (탭 기능)")
        self.setGeometry(100, 100, 1220, 820)

        # EasyOCR 리더는 프로그램 시작 시 한 번만 생성하여 모든 탭에서 공유
        try:
            self.reader = easyocr.Reader(['ko', 'en'], gpu=False)
        except Exception as e:
            QMessageBox.critical(self, "EasyOCR 로드 오류", f"EasyOCR 초기화 중 치명적 오류 발생: {e}")
            sys.exit()

        # 탭 위젯 생성
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # 1. 경영상태 분석 탭 추가 (기존 OcrUpdaterWindow를 사용)
        self.business_tab = BusinessStatusTab(self.reader) # reader를 전달
        self.tabs.addTab(self.business_tab, "📄 경영상태 분석")

        # 2. 신용평가 업데이트 탭 추가
        self.credit_tab = CreditRatingTab(self.reader)
        self.tabs.addTab(self.credit_tab, "✨ 신용평가 일괄 업데이트")