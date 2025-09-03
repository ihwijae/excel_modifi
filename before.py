# [main.py 파일 전체를 이 ко드로 통째로 교체하세요]
import sys, os, re, shutil
from PySide6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel,
                               QLineEdit, QPushButton, QMessageBox, QFileDialog, QGroupBox, QScrollArea,
                               QTableWidget, QTableWidgetItem, QHeaderView, QComboBox, QTabWidget, QMainWindow)
from PySide6.QtCore import Qt, Signal, QThread, QRect, QPoint
from PySide6.QtGui import QPixmap, QPainter, QPen, QGuiApplication, QImage

import easyocr
from pdf2image import convert_from_path
import numpy as np
from PIL import Image

try:
    import ocr_logic
    import config
    import ocr_utils
except ImportError as e:
    QMessageBox.critical(None, "파일 누락 오류", f"필수 파일을 찾을 수 없습니다: {e}\nmain.py, ocr_logic.py, config.py, ocr_utils.py 파일이 모두 같은 폴더에 있는지 확인해주세요.")
    sys.exit()



# --- 이미지 위에 사각형을 그릴 수 있는 커스텀 라벨 ---
class ImageLabel(QLabel):
    roi_selected = Signal(QRect)
    def __init__(self, parent=None):
        super().__init__(parent); self.begin, self.end, self.selecting = QPoint(), QPoint(), False
    def mousePressEvent(self, event):
        if self.selecting and event.button() == Qt.MouseButton.LeftButton: self.begin, self.end, _ = event.pos(), event.pos(), self.update()
    def mouseMoveEvent(self, event):
        if self.selecting: self.end, _ = event.pos(), self.update()
    def mouseReleaseEvent(self, event):
        if self.selecting and event.button() == Qt.MouseButton.LeftButton: self.selecting, self.roi_selected.emit(QRect(self.begin, self.end).normalized()), self.update()
    def paintEvent(self, event):
        super().paintEvent(event)
        if self.selecting: painter = QPainter(self); painter.setPen(QPen(Qt.red, 2, Qt.SolidLine)); painter.drawRect(QRect(self.begin, self.end).normalized())

# --- Ctrl+휠 줌, Shift+휠 스크롤 기능을 위한 커스텀 스크롤 영역 ---
class ZoomableScrollArea(QScrollArea):
    def __init__(self, main_window, parent=None):
        super().__init__(parent); self.main_window = main_window; self.setWidgetResizable(True); self.setAlignment(Qt.AlignCenter)
    def wheelEvent(self, event):
        modifiers = QGuiApplication.keyboardModifiers()
        if modifiers == Qt.KeyboardModifier.ControlModifier:
            angle = event.angleDelta().y(); self.main_window.zoom_image(1.2 if angle > 0 else 0.8); event.accept()
        elif modifiers == Qt.KeyboardModifier.ShiftModifier:
            delta = event.angleDelta().y(); h_bar = self.horizontalScrollBar(); h_bar.setValue(h_bar.value() - delta); event.accept()
        else: super().wheelEvent(event)

# --- 지정된 영역만 OCR 분석하는 스레드 ---
class RoiOcrWorker(QThread):
    progress = Signal(str, str); finished = Signal(str)
    def __init__(self, reader, image_qimage, fields_to_process):
        super().__init__(); self.reader, self.image_qimage, self.fields_to_process = reader, image_qimage, fields_to_process
    def run(self):
        try:
            pil_image = Image.fromqpixmap(self.image_qimage)
            for field, data in self.fields_to_process.items():
                rect = data.get('roi')
                if not rect: self.progress.emit(field, "[지정 안됨]"); continue
                cropped_pil = pil_image.crop((rect.x(), rect.y(), rect.x() + rect.width(), rect.y() + rect.height()))
                preprocessed_img = ocr_utils.preprocess_image_for_ocr(cropped_pil)
                result = self.reader.readtext(preprocessed_img, detail=0, paragraph=True)
                text = " ".join(result) if result else ""
                self.progress.emit(field, text.strip())
            self.finished.emit("모든 영역 분석 완료!")
        except Exception as e:
            self.finished.emit(f"분석 중 오류 발생: {e}")


# [핵심] 바로 이 위치에, 아래의 새로운 클래스를 통째로 추가하세요
# --- 연말 색상 업데이트를 위한 스레드 클래스 ---
class ColorUpdateWorker(QThread):
    finished = Signal(str)
    def __init__(self, excel_path):
        super().__init__()
        self.excel_path = excel_path
    def run(self):
        # ocr_logic.py에 있는 색상 변경 함수를 호출하고, 결과 메시지를 받음
        result_message = ocr_logic.batch_update_colors(self.excel_path)
        self.finished.emit(result_message)



# [핵심] "신용평가 일괄 업데이트" 탭 (새로운 기능이 들어갈 공간)
class CreditRatingTab(QWidget):
    def __init__(self, reader):
        super().__init__()
        self.reader = reader
        main_layout = QVBoxLayout(self)
        label = QLabel("여기에 '신용평가 일괄 업데이트' 기능이 구현될 예정입니다.")
        label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(label)




# --- 메인 윈도우 클래스 (최종 완성본) ---
class BusinessStatusTab(QWidget):
    def __init__(self, reader):
        super().__init__()
        self.poppler_path = r'C:\poppler\poppler-24.02.0\bin'

        self.reader = reader
        # self.setWindowTitle("협력업체 데이터 관리 프로그램 v2.0")
        # self.setGeometry(100, 100, 1200, 850)
        
        self.original_pixmap = None
        self.scale_factor = 1.0
        self.fields_to_extract = {key:{} for key in config.COLUMN_MAP.keys()}
        self.current_field_to_set = None
        # [핵심] '변경 전' 데이터에서 가져온 업체명을 저장할 변수
        self.current_company_name = None 
        
        self.setup_ui()
        self.connect_signals()
        # try:
        #     self.reader = easyocr.Reader(['ko', 'en'], gpu=False)
        # except Exception as e:
        #     QMessageBox.critical(self, "EasyOCR 로드 오류", f"EasyOCR 초기화 중 오류가 발생했습니다: {e}")

    # [setup_ui 함수를 이 ко드로 통째로 교체하세요]
    def setup_ui(self):
        main_layout = QHBoxLayout(self)
        left_panel = QGroupBox("1. PDF/이미지 뷰어")
        left_layout = QVBoxLayout(left_panel)
        self.image_label = ImageLabel(self)
        self.scroll_area = ZoomableScrollArea(self)
        self.scroll_area.setWidget(self.image_label)
        zoom_layout = QHBoxLayout()
        self.zoom_in_button, self.zoom_out_button, self.zoom_fit_button = QPushButton("➕"), QPushButton("➖"), QPushButton("🔲")
        self.zoom_label = QLabel("100%")
        zoom_layout.addStretch(1); zoom_layout.addWidget(self.zoom_out_button); zoom_layout.addWidget(self.zoom_in_button)
        zoom_layout.addWidget(self.zoom_fit_button); zoom_layout.addWidget(self.zoom_label); zoom_layout.addStretch(1)
        left_layout.addWidget(self.scroll_area, 1); left_layout.addLayout(zoom_layout)
        
        right_panel = QWidget(); right_layout = QVBoxLayout(right_panel); right_panel.setFixedWidth(450)
        file_box = QGroupBox("파일 선택")
        file_layout = QHBoxLayout(file_box)
        self.file_path_entry = QLineEdit(); self.file_path_entry.setReadOnly(True)
        self.file_select_button = QPushButton("📁 파일 열기")
        file_layout.addWidget(self.file_path_entry); file_layout.addWidget(self.file_select_button)
        
        excel_box = QGroupBox("엑셀 정보")
        excel_box = QGroupBox("2. 업데이트할 엑셀 정보")
        excel_layout = QGridLayout(excel_box)
        self.excel_file_path_entry = QLineEdit(); self.excel_file_path_entry.setReadOnly(True)
        self.excel_select_button = QPushButton("📂 엑셀 선택")
        excel_layout.addWidget(QLabel("엑셀 파일:"), 0, 0); excel_layout.addWidget(self.excel_file_path_entry, 0, 1); excel_layout.addWidget(self.excel_select_button, 0, 2)

        self.color_update_button = QPushButton("🎨 연말 색상 업데이트")
        excel_layout.addWidget(QLabel("엑셀 파일:"), 0, 0)
        excel_layout.addWidget(self.excel_file_path_entry, 0, 1)
        excel_layout.addWidget(self.excel_select_button, 0, 2)
        excel_layout.addWidget(self.color_update_button, 1, 0, 1, 3) # 버튼을 아래쪽에 추가
        
        roi_box = QGroupBox("데이터 영역 지정")
        roi_layout = QGridLayout(roi_box)
        for row, field in enumerate(self.fields_to_extract.keys()):
            lbl, btn, entry = QLabel(f"{field}:"), QPushButton("지정"), QLineEdit(); btn.setProperty("field_name", field)
            roi_layout.addWidget(lbl,row,0); roi_layout.addWidget(btn,row,1); roi_layout.addWidget(entry,row,2)
            self.fields_to_extract[field].update({"roi":None, "entry":entry, "button":btn}); btn.clicked.connect(self.prepare_to_set_roi)
        
        preview_box = QGroupBox("4. 변경 전/후 미리보기")
        preview_layout = QHBoxLayout(preview_box)
        self.before_table = self.create_preview_table()
        self.after_table = self.create_preview_table()
        before_vbox = QVBoxLayout(); before_vbox.addWidget(QLabel("<b>변경 전 (엑셀 원본)</b>")); before_vbox.addWidget(self.before_table)
        after_vbox = QVBoxLayout(); after_vbox.addWidget(QLabel("<b>변경 후 (OCR 결과)</b>")); after_vbox.addWidget(self.after_table)
        preview_layout.addLayout(before_vbox); preview_layout.addLayout(after_vbox)

        # [핵심] 5. 보관 폴더 및 '자료 종류' 설정 UI 추가
        archive_box = QGroupBox("5. 처리 완료 파일 보관 및 이름 설정")
        archive_layout = QGridLayout(archive_box)
        self.archive_path_entry = QLineEdit()
        self.archive_path_entry.setPlaceholderText("PDF/이미지를 옮길 폴더를 선택하세요.")
        self.archive_select_button = QPushButton("📂 폴더 선택")
        
        # 자료 종류 선택 드롭다운 메뉴 추가
        self.file_type_combo = QComboBox()
        self.file_type_combo.addItems(["-- 자료 종류 선택 --", "전기경영상태", "통신경영상태", "소방경영상태"])
        
        archive_layout.addWidget(QLabel("보관 폴더:"), 0, 0)
        archive_layout.addWidget(self.archive_path_entry, 0, 1)
        archive_layout.addWidget(self.archive_select_button, 0, 2)
        archive_layout.addWidget(QLabel("자료 종류:"), 1, 0)
        archive_layout.addWidget(self.file_type_combo, 1, 1, 1, 2)

        action_box = QGroupBox("6. 실행")
        action_layout = QVBoxLayout(action_box)
        self.run_ocr_button, self.compare_button, self.save_button = QPushButton("1. 지정 영역 분석"), QPushButton("2. 원본 데이터 비교"), QPushButton("3. 확정 및 엑셀 저장")
        self.save_button.setEnabled(False)
        self.save_button.setStyleSheet("font-weight: bold; background-color: #A93226;")
        action_layout.addWidget(self.run_ocr_button)
        action_layout.addWidget(self.compare_button)
        action_layout.addWidget(self.save_button)
        
        right_layout.addWidget(file_box)
        right_layout.addWidget(excel_box)
        right_layout.addWidget(roi_box)
        right_layout.addWidget(preview_box, 1)
        right_layout.addWidget(archive_box)
        right_layout.addWidget(action_box)
        
        main_layout.addWidget(left_panel, 1)
        main_layout.addWidget(right_panel)

    def create_preview_table(self):
        table = QTableWidget()
        table.setColumnCount(2); table.setRowCount(len(self.fields_to_extract))
        table.verticalHeader().setVisible(False); table.horizontalHeader().setVisible(False)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents); table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        for row, field_name in enumerate(self.fields_to_extract.keys()):
            item = QTableWidgetItem(field_name); item.setFlags(item.flags() & ~Qt.ItemIsEditable); table.setItem(row, 0, item)
        return table

    def connect_signals(self):
        self.file_select_button.clicked.connect(self.open_file)
        self.excel_select_button.clicked.connect(self.select_excel_file)
        self.archive_select_button.clicked.connect(self.select_archive_folder)
        self.image_label.roi_selected.connect(self.on_roi_selected)
        self.zoom_in_button.clicked.connect(lambda: self.zoom_image(1.2))
        self.zoom_out_button.clicked.connect(lambda: self.zoom_image(0.8))
        self.zoom_fit_button.clicked.connect(self.fit_to_window)
        self.run_ocr_button.clicked.connect(self.run_roi_ocr)
        self.compare_button.clicked.connect(self.compare_data)
        self.save_button.clicked.connect(self.save_data_to_excel)
        self.color_update_button.clicked.connect(self.start_color_update)

    def open_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "파일 선택", "", "PDF 및 이미지 파일 (*.pdf *.png *.jpg *.jpeg)")
        if file_path:
            self.file_path_entry.setText(file_path)
            for field in self.fields_to_extract.values():
                if field.get('button'): field['button'].setText("지정"); field['button'].setStyleSheet("")
                if field.get('entry'): field['entry'].clear()
                field['roi'] = None
            try:
                # [핵심] PDF와 이미지 파일 처리 로직을 명확하게 분리
                if file_path.lower().endswith('.pdf'):
                    images = convert_from_path(file_path, poppler_path=self.poppler_path, dpi=300, first_page=1, last_page=1)
                    if images:
                        img_pil = images[0]
                        # Pillow 이미지를 QPixmap으로 변환하기 위해 임시 저장
                        img_pil.save("temp_page.png")
                        self.original_pixmap = QPixmap("temp_page.png")
                    else:
                        QMessageBox.warning(self, "PDF 오류", "PDF 파일을 이미지로 변환할 수 없습니다.")
                        return
                else: # 이미지 파일일 경우
                    self.original_pixmap = QPixmap(file_path)
                
                if self.original_pixmap.isNull():
                    QMessageBox.critical(self, "파일 열기 오류", "이미지 파일을 불러올 수 없습니다. 파일이 손상되었거나 지원하지 않는 형식일 수 있습니다.")
                    self.original_pixmap = None
                    return

                self.scale_factor = 1.0
                self.fit_to_window()

            except Exception as e:
                QMessageBox.critical(self, "파일 열기 오류", f"파일을 여는 중 오류가 발생했습니다:\n{e}")

    def select_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "업데이트할 엑셀 파일 선택", "", "Excel 파일 (*.xlsx *.xls)")
        if file_path: self.excel_file_path_entry.setText(file_path)

    def select_archive_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "보관할 폴더 선택")
        if folder_path: self.archive_path_entry.setText(folder_path)

    def zoom_image(self, factor):
        if self.original_pixmap:
            self.scale_factor *= factor
            new_width = int(self.original_pixmap.width() * self.scale_factor)
            scaled_pixmap = self.original_pixmap.scaledToWidth(new_width, Qt.SmoothTransformation)
            self.image_label.setPixmap(scaled_pixmap)
            self.zoom_label.setText(f"{int(self.scale_factor * 100)}%")

    def fit_to_window(self):
        if self.original_pixmap:
            scaled_pixmap = self.original_pixmap.scaled(self.scroll_area.viewport().size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.image_label.setPixmap(scaled_pixmap)
            if self.original_pixmap.width() > 0: self.scale_factor = self.image_label.pixmap().width() / self.original_pixmap.width()
            self.zoom_label.setText(f"{int(self.scale_factor * 100)}%")

    def prepare_to_set_roi(self):
        sender = self.sender(); self.current_field_to_set = sender.property("field_name")
        self.image_label.selecting = True; self.setCursor(Qt.CrossCursor); self.image_label.setCursor(Qt.CrossCursor)

    def on_roi_selected(self, rect):
        if self.current_field_to_set:
            original_rect = QRect(int(rect.x()/self.scale_factor), int(rect.y()/self.scale_factor), int(rect.width()/self.scale_factor), int(rect.height()/self.scale_factor))
            self.fields_to_extract[self.current_field_to_set]['roi'] = original_rect
            button = self.fields_to_extract[self.current_field_to_set]['button']
            button.setText(f"지정됨({original_rect.x()},{original_rect.y()})"); button.setStyleSheet("background-color: #2ECC71;")
        self.image_label.selecting = False; self.current_field_to_set = None
        self.setCursor(Qt.ArrowCursor); self.image_label.setCursor(Qt.ArrowCursor)
        
    def run_roi_ocr(self):
        if not self.original_pixmap: QMessageBox.warning(self, "오류", "먼저 분석할 이미지를 열어주세요."); return
        fields_to_process = {k: v for k, v in self.fields_to_extract.items() if v.get('roi')}
        if not fields_to_process: QMessageBox.warning(self, "오류", "하나 이상의 영역을 먼저 지정해주세요."); return
        self.run_ocr_button.setEnabled(False); self.run_ocr_button.setText("분석 중...")
        original_qimage = self.original_pixmap.toImage()
        self.worker = RoiOcrWorker(self.reader, original_qimage, fields_to_process)
        self.worker.progress.connect(self.update_ocr_result); self.worker.finished.connect(self.on_ocr_finished); self.worker.start()

    def update_ocr_result(self, field_name, text):
        cleaned_text = text
        if field_name == '사업자등록번호': cleaned_text = ocr_utils.clean_biz_number(text)
        elif '실적' in field_name or '시평액' in field_name: cleaned_text = ocr_utils.clean_ocr_number(text)
        elif '비율' in field_name: cleaned_text = "".join(re.findall(r'[\d.]', text))
        self.fields_to_extract[field_name]['entry'].setText(cleaned_text)

    def on_ocr_finished(self, message):
        self.run_ocr_button.setEnabled(True); self.run_ocr_button.setText("1. 지정 영역 분석")
        QMessageBox.information(self, "분석 완료", message)

    # [compare_data 함수를 이 ко드로 통째로 교체하세요]
    def compare_data(self):
        excel_path = self.excel_file_path_entry.text()
        biz_no = self.fields_to_extract['사업자등록번호']['entry'].text().strip()
        if not (excel_path and biz_no):
            QMessageBox.warning(self, "오류", "엑셀 파일과 사업자등록번호가 모두 필요합니다.")
            return
            
        before_data, error = ocr_logic.find_company_data(excel_path, biz_no)
        if error:
            QMessageBox.critical(self, "조회 오류", error)
            return
        
        # [핵심] 파일명에 사용할 원본 업체명을 self 변수에 저장
        self.current_company_name = before_data.get('상호')
            
        self.populate_preview_table(self.before_table, before_data, is_after=False)
        after_data = {k: v['entry'].text() for k, v in self.fields_to_extract.items()}
        self.populate_preview_table(self.after_table, after_data, is_after=True)
        
        self.save_button.setEnabled(True)
        QMessageBox.information(self, "비교 완료", "내용을 확인하고 '3. 확정 및 엑셀 저장' 버튼을 누르세요.")

        # [save_data_to_excel 함수를 이 ко드로 통째로 교체하세요]
    def save_data_to_excel(self):
        source_file_path = self.file_path_entry.text()
        excel_path = self.excel_file_path_entry.text()
        archive_folder = self.archive_path_entry.text()
        file_type = self.file_type_combo.currentText()
        biz_no = self.fields_to_extract['사업자등록번호']['entry'].text().strip()

        # --- 1. 모든 정보가 준비되었는지 확인 (안전장치) ---
        if not (excel_path and biz_no and source_file_path):
            QMessageBox.warning(self, "정보 부족", "PDF/이미지, 엑셀, 사업자등록번호가 모두 필요합니다."); return
        if not archive_folder:
            QMessageBox.warning(self, "경로 지정 필요", "5번 항목에서 '보관 폴더'를 먼저 지정해주세요."); return
        if file_type == "-- 자료 종류 선택 --":
            QMessageBox.warning(self, "종류 선택 필요", "5번 항목에서 '자료 종류'를 먼저 선택해주세요."); return
        if not self.current_company_name:
            QMessageBox.warning(self, "업체명 오류", "'2. 원본 데이터 비교'를 먼저 실행하여 원본 업체명을 불러와주세요."); return

        # --- 2. 엑셀 업데이트 실행 ---
        update_data = {k: v['entry'].text() for k, v in self.fields_to_extract.items()}
        updated_log, error = ocr_logic.update_company_data(excel_path, biz_no, update_data)
        
        if error:
            QMessageBox.critical(self, "엑셀 업데이트 오류", error)
            return

        # --- 3. 엑셀 업데이트 성공 시, 파일 이동 및 이름 변경 실행 ---
        try:
            # 새로운 파일명 생성: (주)이이이주식회사_전기경영상태.pdf
            original_filename_with_ext = os.path.basename(source_file_path)
            _, file_extension = os.path.splitext(original_filename_with_ext)
            new_filename = f"{self.current_company_name}_{file_type}{file_extension}"
            
            destination_path = os.path.join(archive_folder, new_filename)
            
            # 중복 파일 처리
            count = 1
            while os.path.exists(destination_path):
                name, ext = os.path.splitext(new_filename)
                destination_path = os.path.join(archive_folder, f"{name} ({count}){ext}")
                count += 1
                
            shutil.move(source_file_path, destination_path)
            
            final_message = f"성공적으로 업데이트했습니다.\n\n<b>[수정된 항목]</b>\n{', '.join(updated_log)}\n\n"
            final_message += f"<b>[파일 이동]</b>\n'{original_filename_with_ext}' 파일을\n'{new_filename}'으로 변경하여 저장했습니다."
            
            QMessageBox.information(self, "모든 작업 완료", final_message)
            
            # [핵심] 모든 작업이 끝난 후, UI를 깔끔하게 초기화
            self.reset_ui_for_next_file()

        except Exception as e:
            QMessageBox.critical(self, "파일 이동 오류", f"엑셀 업데이트는 성공했지만, 원본 파일을 이동하는 중 오류가 발생했습니다:\n{e}")
        
        self.save_button.setEnabled(False)

    def populate_preview_table(self, table, data, is_after=False):
        if not table: return
        for row, key in enumerate(self.fields_to_extract.keys()):
            value = data.get(key, "")
            display_text = ""
            if value is not None and value != "":
                try:
                    if '비율' in key:
                        numeric_value = float(str(value).replace('%', ''))
                        display_text = f"{numeric_value * 100 if not is_after else numeric_value:.2f}%"
                    elif '실적' in key or '시평액' in key:
                        numeric_value = int(float(str(value).replace(',', '')))
                        display_text = f"{numeric_value * 1000 if is_after else numeric_value:,}"
                    else:
                        display_text = str(value)
                except (ValueError, TypeError):
                    display_text = str(value)
            
            value_item = QTableWidgetItem(display_text)
            value_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            table.setItem(row, 1, value_item)
        table.resizeRowsToContents()

    def reset_ui_for_next_file(self):
        """모든 작업 완료 후, 다음 파일을 처리하기 위해 UI를 초기화합니다."""
        # 1. 파일 경로들 초기화
        self.file_path_entry.clear()
        
        # 2. 이미지 뷰어 초기화
        self.original_pixmap = None
        self.image_label.clear()
        
        # 3. 오른쪽 패널 초기화
        for field in self.fields_to_extract.values():
            if field.get('button'):
                field['button'].setText("지정")
                field['button'].setStyleSheet("")
            if field.get('entry'):
                field['entry'].clear()
            field['roi'] = None
            
        # 4. 미리보기 테이블 초기화
        self.before_table.clearContents()
        self.after_table.clearContents()
        
        # 5. [핵심] 자료 종류 드롭다운 메뉴 초기화
        self.file_type_combo.setCurrentIndex(0)
        
        # 6. 저장 버튼 비활성화
        self.save_button.setEnabled(False)

    def start_color_update(self):
            excel_path = self.excel_file_path_entry.text()
            if not excel_path:
                QMessageBox.warning(self, "파일 선택 오류", "먼저 색상을 업데이트할 엑셀 파일을 선택해주세요.")
                return

            reply = QMessageBox.question(self, "연말 색상 업데이트 확인",
                                        f"'{os.path.basename(excel_path)}' 파일의 모든 데이터 상태 색상을 갱신하시겠습니까?\n\n"
                                        "- 초록색 -> 파란색\n"
                                        "- 파란색 -> 색 없음\n\n"
                                        "(이 작업은 되돌릴 수 없습니다!)",
                                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                        QMessageBox.StandardButton.No)

            if reply == QMessageBox.StandardButton.Yes:
                self.color_update_button.setText("업데이트 중...")
                self.color_update_button.setEnabled(False)
                # 백그라운드 스레드로 실행
                self.color_worker = ColorUpdateWorker(excel_path)
                self.color_worker.finished.connect(self.on_color_update_finished)
                self.color_worker.start()

    def on_color_update_finished(self, message):
        self.color_update_button.setText("🎨 연말 색상 업데이트")
        self.color_update_button.setEnabled(True)
        if "오류" in message:
            QMessageBox.critical(self, "업데이트 오류", message)
        else:
            QMessageBox.information(self, "업데이트 완료", message)



# [핵심] 색상 업데이트 작업을 위한 새로운 스레드 클래스 추가
class ColorUpdateWorker(QThread):
    finished = Signal(str)
    def __init__(self, excel_path):
        super().__init__()
        self.excel_path = excel_path
    def run(self):
        result_message = ocr_logic.batch_update_colors(self.excel_path)
        self.finished.emit(result_message)



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



if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow() # 새로운 메인 윈도우를 실행
    window.show()
    sys.exit(app.exec())