# [credit_rating_tab.py 파일 전체를 이 코드로 교체하세요]

import sys
import os
import re
import shutil
import json
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel,
                               QLineEdit, QPushButton, QMessageBox, QFileDialog, QGroupBox, QScrollArea,
                               QTextEdit, QApplication, QDateEdit, QDialog, QComboBox,
                               QDialogButtonBox, QCheckBox)
from PySide6.QtCore import Qt, QRect, QDate, QThread, Signal
from PySide6.QtGui import QPixmap, QImage, QColor, QFont
from PIL import Image
import fitz  # PyMuPDF

# 우리 프로젝트의 다른 파일들
import ocr_logic
import ocr_utils
from ui_widgets import ImageLabel, ZoomableScrollArea
from workers import RoiOcrWorker
from PySide6.QtGui import QTransform
from business_status_tab import BusinessStatusTab

# --- 헬퍼 클래스 및 함수 ---

# 페이지 내보내기 팝업창 클래스
class PdfExportDialog(QDialog):
    def __init__(self, max_page, parent=None):
        super().__init__(parent)
        self.setWindowTitle("페이지 내보내기")
        layout = QVBoxLayout(self)
        self.info_label = QLabel(f"내보낼 페이지를 입력하세요 (총 {max_page}페이지).\n(예: 1, 3-5, 8)")
        self.page_edit = QLineEdit()
        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.accept); self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.info_label); layout.addWidget(self.page_edit); layout.addWidget(self.button_box)
    def get_page_selection(self):
        return self.page_edit.text()

# 페이지 범위 텍스트를 숫자 리스트로 변환하는 함수
def parse_page_ranges(range_string, max_page):
    pages = set()
    try:
        parts = range_string.split(',')
        for part in parts:
            part = part.strip()
            if '-' in part:
                start, end = map(int, part.split('-'))
                if start > 0 and end <= max_page:
                    for i in range(start, end + 1): pages.add(i - 1)
            else:
                page_num = int(part)
                if 0 < page_num <= max_page: pages.add(page_num - 1)
    except: return None
    return sorted(list(pages))

# 신규 업체 정보 입력 팝업창 클래스
class ManualInputDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("신규 업체 정보 입력"); layout = QVBoxLayout(self)
        self.info_label = QLabel("DB에서 업체를 찾을 수 없습니다.\n파일을 보관하려면 아래 정보를 직접 입력해주세요.")
        form_layout = QGridLayout(); self.name_label = QLabel("업체명:"); self.name_edit = QLineEdit(); self.region_label = QLabel("지역:"); self.region_combo = QComboBox()
        self.region_combo.addItems(["서울", "경기", "인천", "부산", "대구", "광주", "대전", "울산", "세종", "강원", "충북", "충남", "전북", "전남", "경북", "경남", "제주"])
        form_layout.addWidget(self.name_label, 0, 0); form_layout.addWidget(self.name_edit, 0, 1); form_layout.addWidget(self.region_label, 1, 0); form_layout.addWidget(self.region_combo, 1, 1)
        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.accept); self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.info_label); layout.addLayout(form_layout); layout.addWidget(self.button_box)
    def get_data(self): return self.name_edit.text().strip(), self.region_combo.currentText()

# --- CreditRatingTab 클래스 ---
class CreditRatingTab(QWidget):
    def __init__(self, reader):
        super().__init__()
        self.reader = reader
        self.original_pixmap = None
        self.scale_factor = 1.0
        self.excel_paths = {}
        self.fields_to_extract = {'사업자등록번호': {}, '신용평가등급': {}}
        self.current_field_to_set = None
        self.pdf_pages = []
        self.current_page_index = 0
        self.found_company_data = None
        self.setup_ui()
        self.connect_signals()
        self.load_paths()
        self.reset_ui()

    def create_viewer_panel(self):
        panel = QGroupBox("1. PDF/이미지 뷰어")
        layout = QVBoxLayout(panel)
        self.image_label = ImageLabel(self)
        self.scroll_area = ZoomableScrollArea(self, self)
        self.scroll_area.setWidget(self.image_label)

        controls_layout = QHBoxLayout()
        self.prev_page_button = QPushButton("< 이전")
        self.next_page_button = QPushButton("다음 >")
        self.page_label = QLabel("0 / 0")
        self.export_pdf_button = QPushButton("📄 페이지 내보내기")

        # [추가] 회전 버튼 생성
        self.rotate_left_button = QPushButton("↩️")
        self.rotate_right_button = QPushButton("↪️")

        self.zoom_in_button, self.zoom_out_button, self.zoom_fit_button = QPushButton("➕"), QPushButton(
            "➖"), QPushButton("🔲")
        self.zoom_label = QLabel("100%")

        controls_layout.addStretch(1)
        controls_layout.addWidget(self.prev_page_button)
        controls_layout.addWidget(self.page_label)
        controls_layout.addWidget(self.next_page_button)
        controls_layout.addWidget(self.export_pdf_button)
        controls_layout.addStretch(1)

        # [추가] 회전 버튼을 레이아웃에 추가
        controls_layout.addWidget(self.rotate_left_button)
        controls_layout.addWidget(self.rotate_right_button)

        controls_layout.addStretch(1)
        controls_layout.addWidget(self.zoom_out_button)
        controls_layout.addWidget(self.zoom_in_button)
        controls_layout.addWidget(self.zoom_fit_button)
        controls_layout.addWidget(self.zoom_label)
        controls_layout.addStretch(1)

        layout.addWidget(self.scroll_area, 1)
        layout.addLayout(controls_layout)
        return panel

    # [credit_rating_tab.py 파일에서 이 함수를 찾아 통째로 교체]

    def create_controls_panel(self):
        panel = QWidget();
        layout = QVBoxLayout(panel);
        panel.setFixedWidth(500)

        # --- 폰트 정의 ---
        group_font = QFont(); group_font.setPointSize(12)
        label_font = QFont(); label_font.setPointSize(11)
        entry_font = QFont(); entry_font.setPointSize(12)

        # [핵심 수정] self.file_box로 변경
        self.file_box = QGroupBox("1. 신용평가 자료 선택"); self.file_box.setFont(group_font)
        file_layout = QHBoxLayout(self.file_box)
        self.file_path_entry = QLineEdit(); self.file_path_entry.setFont(entry_font)
        self.file_path_entry.setReadOnly(True)
        self.file_select_button = QPushButton("📁 파일 열기");
        file_layout.addWidget(self.file_path_entry);
        file_layout.addWidget(self.file_select_button)

        original_data_box = QGroupBox("DB 원본 정보 (변경 전)"); original_data_box.setFont(group_font)
        original_layout = QGridLayout(original_data_box)
        self.original_name_display = QLineEdit(); self.original_name_display.setFont(entry_font)
        self.original_name_display.setReadOnly(True)
        self.original_rating_display = QTextEdit(); self.original_rating_display.setFont(entry_font)
        self.original_rating_display.setReadOnly(True);
        self.original_rating_display.setFixedHeight(60)
        original_layout.addWidget(QLabel("상호:"), 0, 0); original_layout.addWidget(self.original_name_display, 0, 1)
        original_layout.addWidget(QLabel("기존 신용평가:"), 1, 0); original_layout.addWidget(self.original_rating_display, 1, 1)

        roi_box = QGroupBox("데이터 영역 지정 및 입력 (변경 후)"); roi_box.setFont(group_font)
        roi_layout = QGridLayout(roi_box)
        for row, field in enumerate(self.fields_to_extract.keys()):
            lbl, btn, entry = QLabel(f"{field}:"), QPushButton("지정"), QLineEdit();
            lbl.setFont(label_font); entry.setFont(entry_font)
            btn.setProperty("field_name", field)
            roi_layout.addWidget(lbl, row, 0);
            roi_layout.addWidget(btn, row, 1);
            roi_layout.addWidget(entry, row, 2, 1, 3)
            self.fields_to_extract[field].update({"roi": None, "entry": entry, "button": btn})
        self.period_label = QLabel("유효기간:"); self.period_label.setFont(label_font)
        self.start_date_edit = QDateEdit(calendarPopup=True); self.start_date_edit.setFont(entry_font)
        self.end_date_edit = QDateEdit(calendarPopup=True); self.end_date_edit.setFont(entry_font)
        self.start_date_edit.setDate(QDate.currentDate());
        self.end_date_edit.setDate(QDate.currentDate().addYears(1).addDays(-1))
        self.start_date_edit.setDisplayFormat("yyyy.MM.dd");
        self.end_date_edit.setDisplayFormat("yyyy.MM.dd")
        roi_layout.addWidget(self.period_label, 2, 0);
        roi_layout.addWidget(self.start_date_edit, 2, 2)
        roi_layout.addWidget(QLabel("~"), 2, 3, Qt.AlignCenter);
        roi_layout.addWidget(self.end_date_edit, 2, 4)
        self.combined_preview_label = QLabel("<b>최종 저장될 값 (신용평가):</b>")
        self.combined_preview = QTextEdit(); self.combined_preview.setFont(entry_font)
        self.combined_preview.setReadOnly(True);
        self.combined_preview.setFixedHeight(60)
        roi_layout.addWidget(self.combined_preview_label, 3, 0, 1, 5);
        roi_layout.addWidget(self.combined_preview, 4, 0, 1, 5)

        self.archive_box = QGroupBox("처리 완료 파일 보관 경로"); self.archive_box.setFont(group_font)
        archive_layout = QHBoxLayout(self.archive_box)
        self.archive_path_entry = QLineEdit(); self.archive_path_entry.setFont(entry_font)
        self.archive_path_entry.setReadOnly(True)
        self.archive_path_entry.setPlaceholderText("경영 상태 분석 탭의 '경로 설정'에서 지정")
        archive_layout.addWidget(self.archive_path_entry)

        action_box = QGroupBox("실행"); action_box.setFont(group_font)
        action_layout = QVBoxLayout(action_box)
        self.run_ocr_button = QPushButton("1. 이미지에서 글자 분석");
        self.lookup_button = QPushButton("2. 사업자번호로 업체 조회")
        self.data_only_checkbox = QCheckBox("자료 파일 없이 데이터만 저장")
        self.update_button = QPushButton("3. 업데이트 및 파일 보관");
        self.update_button.setStyleSheet("font-weight: bold; background-color: #1E8449; color: white;")
        self.log_display = QTextEdit(); self.log_display.setReadOnly(True)
        action_layout.addWidget(self.run_ocr_button);
        action_layout.addWidget(self.lookup_button)
        action_layout.addWidget(self.data_only_checkbox)
        action_layout.addWidget(self.update_button);
        action_layout.addWidget(QLabel("진행 로그:"));
        action_layout.addWidget(self.log_display)

        layout.addWidget(self.file_box);
        layout.addWidget(original_data_box);
        layout.addWidget(roi_box)
        layout.addWidget(self.archive_box);
        layout.addWidget(action_box);
        layout.addStretch(1)
        return panel

    def connect_signals(self):
        self.prev_page_button.clicked.connect(self.show_previous_page)
        self.next_page_button.clicked.connect(self.show_next_page)
        self.export_pdf_button.clicked.connect(self.export_pdf_pages)
        self.rotate_left_button.clicked.connect(lambda: self.rotate_image(-90))
        self.rotate_right_button.clicked.connect(lambda: self.rotate_image(90))
        self.file_select_button.clicked.connect(lambda: self.open_file())
        self.image_label.roi_selected.connect(self.on_roi_selected)
        self.zoom_in_button.clicked.connect(lambda: self.zoom_image(1.2))
        self.zoom_out_button.clicked.connect(lambda: self.zoom_image(0.8))
        self.run_ocr_button.clicked.connect(self.run_roi_ocr)
        self.lookup_button.clicked.connect(self.run_company_lookup)
        self.update_button.clicked.connect(self.run_final_update)
        for field, data in self.fields_to_extract.items():
            data['button'].clicked.connect(self.prepare_to_set_roi)
            if field == '신용평가등급':
                data['entry'].textChanged.connect(self.update_combined_preview)
            if field == '사업자등록번호':
                data['entry'].textChanged.connect(self.format_biz_no_input)
        self.start_date_edit.dateChanged.connect(self.auto_set_end_date)
        self.end_date_edit.dateChanged.connect(self.update_combined_preview)
        self.data_only_checkbox.stateChanged.connect(self.toggle_file_inputs)

    def open_file(self, file_path=None):
        if file_path is None:
            file_path, _ = QFileDialog.getOpenFileName(self, "신용평가 자료 선택", "", "PDF 및 이미지 파일 (*.pdf *.png *.jpg *.jpeg)")
        if not file_path: return
        self.reset_ui()
        self.file_path_entry.setText(file_path)
        try:
            self.pdf_pages = []
            if file_path.lower().endswith('.pdf'):
                doc = fitz.open(file_path)
                temp_files_to_delete = []  # 삭제할 임시 파일 목록


                for page in doc:
                    pix = page.get_pixmap(dpi=300)
                    img_pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                    # 각 탭에 맞는 고유한 임시 파일 이름 생성
                    tab_name = "bs" if isinstance(self, BusinessStatusTab) else "cr"
                    temp_image_path = f"temp_page_{tab_name}_{page.number}.png"

                    img_pil.save(temp_image_path)
                    self.pdf_pages.append(QPixmap(temp_image_path))
                    temp_files_to_delete.append(temp_image_path)

                doc.close()

                for temp_file in temp_files_to_delete:
                    if os.path.exists(temp_file):
                        os.remove(temp_file)

                if not self.pdf_pages: self.original_pixmap = None
                else:
                    self.original_pixmap = self.pdf_pages[0]
                    if len(self.pdf_pages) > 1:
                        QMessageBox.information(self, "알림", f"이 PDF는 총 {len(self.pdf_pages)}페이지입니다.\n아래 페이지 이동 버튼으로 모든 페이지를 확인하세요.")
            else:
                self.original_pixmap = QPixmap(file_path)
            if self.original_pixmap is None or self.original_pixmap.isNull():
                self.image_label.clear(); self.set_page_controls_visibility(False); return
            self.display_page(0); self.fit_to_window()
        except Exception as e:
            QMessageBox.critical(self, "파일 열기 오류", f"파일을 여는 중 오류가 발생했습니다:\n{e}")

    def display_page(self, page_index):
        if not self.pdf_pages and self.original_pixmap:
            self.image_label.setPixmap(self.original_pixmap); self.set_page_controls_visibility(False); return
        if 0 <= page_index < len(self.pdf_pages):
            self.current_page_index = page_index
            self.original_pixmap = self.pdf_pages[self.current_page_index]
            self.fit_to_window()
            self.page_label.setText(f"{self.current_page_index + 1} / {len(self.pdf_pages)}")
            self.prev_page_button.setEnabled(self.current_page_index > 0)
            self.next_page_button.setEnabled(self.current_page_index < len(self.pdf_pages) - 1)
            self.set_page_controls_visibility(len(self.pdf_pages) > 0)

    def show_previous_page(self): self.display_page(self.current_page_index - 1)
    def show_next_page(self): self.display_page(self.current_page_index + 1)
    def set_page_controls_visibility(self, visible):
        self.prev_page_button.setVisible(visible); self.next_page_button.setVisible(visible)
        self.page_label.setVisible(visible); self.export_pdf_button.setVisible(visible)

    def export_pdf_pages(self):
        source_path = self.file_path_entry.text()
        if not self.pdf_pages or not source_path:
            QMessageBox.warning(self, "알림", "PDF 파일만 페이지를 내보낼 수 있습니다.")
            return

        dialog = PdfExportDialog(len(self.pdf_pages), self)
        if dialog.exec():
            page_selection = dialog.get_page_selection()
            page_indices = parse_page_ranges(page_selection, len(self.pdf_pages))

            if page_indices is None or not page_indices:
                QMessageBox.critical(self, "입력 오류", "페이지 형식이 잘못되었습니다. (예: 1, 3-5, 8)")
                return

            default_name = os.path.splitext(os.path.basename(source_path))[0] + "_분리됨.pdf"
            save_path, _ = QFileDialog.getSaveFileName(self, "분리된 PDF 저장", default_name, "PDF Files (*.pdf)")

            if save_path:
                source_doc = None
                try:
                    source_doc = fitz.open(source_path)
                    new_doc = fitz.open()

                    for page_num in page_indices:
                        new_doc.insert_pdf(source_doc, from_page=page_num, to_page=page_num)

                    new_doc.save(save_path)
                    new_doc.close()
                    QMessageBox.information(self, "성공", f"선택된 페이지를 '{os.path.basename(save_path)}' 파일로 저장했습니다.")

                    page_numbers_str = ", ".join(str(p + 1) for p in page_indices)
                    reply = QMessageBox.question(self, "원본 파일 수정",
                                                 f"원본 PDF 파일에서 내보낸 {page_numbers_str} 페이지를 삭제하시겠습니까?",
                                                 QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                                 QMessageBox.StandardButton.No)

                    if reply == QMessageBox.StandardButton.Yes:
                        for page_num in reversed(page_indices):
                            source_doc.delete_page(page_num)

                        # [핵심 수정] 임시 파일에 먼저 저장한 후, 원본 파일을 대체하는 방식으로 변경
                        temp_path = source_path + ".tmp"
                        source_doc.save(temp_path, garbage=4, deflate=True)
                        source_doc.close()  # 저장 후 반드시 닫기
                        source_doc = None  # 닫혔음을 명시

                        os.remove(source_path)  # 기존 원본 파일 삭제
                        os.rename(temp_path, source_path)  # 임시 파일의 이름을 원본 파일 이름으로 변경

                        QMessageBox.information(self, "작업 완료", "선택한 페이지를 원본에서 삭제했습니다.\n뷰어를 새로고침합니다.")
                        self.open_file(source_path)

                except Exception as e:
                    QMessageBox.critical(self, "내보내기 오류", f"PDF를 처리하는 중 오류가 발생했습니다:\n{e}")
                finally:
                    if source_doc:
                        source_doc.close()

    def reset_ui(self):
        self.file_path_entry.clear();
        self.original_pixmap = None;
        self.image_label.clear();
        self.zoom_label.setText("100%")
        self.pdf_pages = [];
        self.current_page_index = 0;
        self.display_page(0)
        self.found_company_data = None

        for field in self.fields_to_extract.values():
            if field.get('button'): field['button'].setText("지정"); field['button'].setStyleSheet("")
            if field.get('entry'): field['entry'].clear()
            field['roi'] = None

        self.start_date_edit.setDate(QDate.currentDate())
        self.combined_preview.clear();
        self.log_display.clear()

        # [핵심 수정] 원본 정보 표시 칸도 초기화
        self.original_name_display.clear()
        self.original_rating_display.clear()

        self.set_data_input_enabled(False)
        self.update_button.setEnabled(False)

    def run_roi_ocr(self):
        if not self.original_pixmap: QMessageBox.warning(self, "오류", "먼저 분석할 이미지를 열어주세요."); return
        fields_to_process = {k: v for k, v in self.fields_to_extract.items() if v.get('roi')}
        if not fields_to_process: QMessageBox.warning(self, "오류", "하나 이상의 영역을 먼저 지정해주세요."); return
        self.run_ocr_button.setEnabled(False); self.run_ocr_button.setText("분석 중...")
        current_qimage = self.original_pixmap.toImage()
        self.worker = RoiOcrWorker(self.reader, current_qimage, fields_to_process)
        self.worker.progress.connect(self.update_ocr_result)
        self.worker.finished.connect(lambda: self.run_ocr_button.setEnabled(True) or self.run_ocr_button.setText("1. 이미지에서 글자 분석"))
        self.worker.start()

    # --- (이하 나머지 함수들은 이전 최종본과 동일) ---
    def setup_ui(self): main_layout = QHBoxLayout(self); viewer_panel = self.create_viewer_panel(); controls_panel = self.create_controls_panel(); main_layout.addWidget(viewer_panel, 1); main_layout.addWidget(controls_panel)
    def zoom_image(self, factor):
        if self.original_pixmap: self.scale_factor *= factor; self.image_label.setPixmap(self.original_pixmap.scaledToWidth(int(self.original_pixmap.width() * self.scale_factor), Qt.SmoothTransformation)); self.zoom_label.setText(f"{int(self.scale_factor * 100)}%")
    def fit_to_window(self):
        if self.original_pixmap: scaled_pixmap = self.original_pixmap.scaled(self.scroll_area.viewport().size(), Qt.KeepAspectRatio, Qt.SmoothTransformation); self.image_label.setPixmap(scaled_pixmap)
        if self.original_pixmap and self.original_pixmap.width() > 0: self.scale_factor = self.image_label.pixmap().width() / self.original_pixmap.width(); self.zoom_label.setText(f"{int(self.scale_factor * 100)}%")
    def prepare_to_set_roi(self): sender = self.sender(); self.current_field_to_set = sender.property("field_name"); self.image_label.selecting = True; self.setCursor(Qt.CrossCursor)
    def on_roi_selected(self, rect):
        if self.current_field_to_set:
            original_rect = QRect(int(rect.x() / self.scale_factor), int(rect.y() / self.scale_factor), int(rect.width() / self.scale_factor), int(rect.height() / self.scale_factor))
            self.fields_to_extract[self.current_field_to_set]['roi'] = original_rect
            btn = self.fields_to_extract[self.current_field_to_set]['button']; btn.setText("지정됨"); btn.setStyleSheet("background-color: #2ECC71;")
        self.image_label.selecting = False; self.setCursor(Qt.ArrowCursor)
    def update_ocr_result(self, field_name, text):
        cleaned_text = text
        if field_name == '사업자등록번호': cleaned_text = ocr_utils.clean_biz_number(text)
        self.fields_to_extract[field_name]['entry'].setText(cleaned_text)
    def load_paths(self):
        try:
            # 설정 파일 기준으로 항상 최신 상태를 반영한다.
            self.excel_paths.clear()
            if os.path.exists("ocr_config.json"):
                with open("ocr_config.json", 'r', encoding='utf-8') as f:
                    paths = json.load(f)
                    self.excel_paths.update(paths)
                    archive_path = paths.get("archive_path", "")
                    self.archive_path_entry.setText(archive_path)
        except Exception as e:
            print(f"설정 파일 로드 오류: {e}")

    def on_tab_activated(self):
        self.load_paths()
    def auto_set_end_date(self):
        start_date = self.start_date_edit.date(); end_date = start_date.addYears(1).addDays(-1)
        self.end_date_edit.blockSignals(True); self.end_date_edit.setDate(end_date); self.end_date_edit.blockSignals(False)
        self.update_combined_preview()
    def update_combined_preview(self):
        grade = self.fields_to_extract['신용평가등급']['entry'].text().strip(); start_date = self.start_date_edit.date().toString("yy.MM.dd"); end_date = self.end_date_edit.date().toString("yy.MM.dd")
        period = f"({start_date}~{end_date})"; combined_text = f"{grade}\n{period}"; self.combined_preview.setText(combined_text)

    def format_biz_no_input(self, text):
        sender = self.sender()
        if not isinstance(sender, QLineEdit): return

        cleaned_text = re.sub(r'[^0-9]', '', text)[:10]
        formatted_text = cleaned_text
        if len(cleaned_text) > 5:
            formatted_text = f"{cleaned_text[:3]}-{cleaned_text[3:5]}-{cleaned_text[5:]}"
        elif len(cleaned_text) > 3:
            formatted_text = f"{cleaned_text[:3]}-{cleaned_text[3:]}"

        if text != formatted_text:
            sender.blockSignals(True)
            sender.setText(formatted_text)
            sender.blockSignals(False)
            sender.end(False)

    def run_company_lookup(self):
        # 조회 직전에 경로를 다시 읽어 최신 설정을 사용한다.
        self.load_paths()
        biz_no = self.fields_to_extract['사업자등록번호']['entry'].text().strip()
        if not biz_no: QMessageBox.warning(self, "정보 부족", "사업자등록번호를 먼저 입력(또는 분석)해야 합니다."); return
        self.lookup_button.setText("조회 중...");
        QApplication.processEvents()

        company_data = None
        for db_path in self.excel_paths.values():
            if db_path and os.path.exists(db_path):
                data, _ = ocr_logic.find_company_data(db_path, biz_no)
                if data: company_data = data; break

        self.lookup_button.setText("2. 사업자번호로 업체 조회")

        if company_data:
            self.found_company_data = company_data

            # [핵심 수정] 조회된 원본 데이터를 UI에 표시
            company_name = company_data.get('상호', {}).get('value', '')
            original_rating = company_data.get('신용평가', {}).get('value', '')
            self.original_name_display.setText(company_name)
            self.original_rating_display.setText(str(original_rating))

            self.set_data_input_enabled(True)
            QMessageBox.information(self, "업체 확인",
                                    f"<b>{company_name}</b>\n\nDB에서 업체를 찾았습니다.\n새로운 등급과 유효기간을 입력 후 '업데이트' 버튼을 누르세요.")
        else:
            self.found_company_data = None
            self.set_data_input_enabled(False)
            self.original_name_display.clear()
            self.original_rating_display.clear()
            QMessageBox.warning(self, "조회 실패", "DB에서 해당 업체를 찾을 수 없습니다.\n신규 업체인 경우, '업데이트 및 파일 보관' 버튼을 눌러주세요.")

    # [credit_rating_tab.py 파일에서 이 함수를 찾아 통째로 교체]

    def run_final_update(self):
        # 업데이트 직전에 경로를 다시 읽어 최신 설정을 사용한다.
        self.load_paths()
        biz_no = self.fields_to_extract['사업자등록번호']['entry'].text().strip()
        if not biz_no:
            QMessageBox.warning(self, "정보 부족", "사업자등록번호가 필요합니다.");
            return

        data_only_mode = self.data_only_checkbox.isChecked()

        # 데이터만 저장 모드가 아닐 때만 파일/보관 경로를 필수로 확인
        if not data_only_mode:
            source_file_path = self.file_path_entry.text().strip()
            base_archive_path = self.archive_path_entry.text().strip()
            if not (source_file_path and base_archive_path):
                QMessageBox.warning(self, "경로 부족", "자료 파일과 보관 경로를 모두 지정해야 합니다.");
                return

        # --- 실제 작업 실행 ---
        if self.found_company_data:
            company_name = self.found_company_data.get('상호', {}).get('value')
            region_full = self.found_company_data.get('지역', {}).get('value', '기타')
            region_name = region_full.split(' ')[0] if region_full else '기타'
            self.perform_excel_update_and_archive(company_name, region_name, update_excel=True,
                                                  archive_file=not data_only_mode)
        else:
            # 데이터만 저장 모드는 기존 업체만 가능
            if data_only_mode:
                QMessageBox.warning(self, "오류", "'데이터만 저장' 모드는 업체 조회에 성공한 경우에만 사용할 수 있습니다.");
                return

            # 신규 업체 파일 보관 로직 (기존과 동일)
            dialog = ManualInputDialog(self)
            if dialog.exec():
                manual_name, manual_region = dialog.get_data()
                if not (manual_name and manual_region):
                    QMessageBox.warning(self, "정보 부족", "업체명과 지역을 모두 입력해야 합니다.")
                else:
                    self.perform_excel_update_and_archive(manual_name, manual_region, update_excel=False,
                                                          archive_file=True)

    def perform_excel_update_and_archive(self, company_name, region_name, update_excel=True, archive_file=True):
        credit_rating = self.combined_preview.toPlainText().strip()
        biz_no = self.fields_to_extract['사업자등록번호']['entry'].text().strip()

        if update_excel and (not credit_rating or "\n" not in credit_rating):
            QMessageBox.warning(self, "정보 부족", "신용평가등급과 유효기간을 모두 입력해야 합니다.");
            return

        self.update_button.setText("실행 중...");
        self.log_display.clear();
        QApplication.processEvents()

        total_updates = 0
        if update_excel:
            for db_type, excel_path in self.excel_paths.items():
                if excel_path and os.path.exists(excel_path):
                    self.log_display.append(f"'{db_type}' DB 업데이트 시도...");
                    QApplication.processEvents()
                    result_msg, error = ocr_logic.update_credit_rating_only(excel_path, biz_no, credit_rating)
                    if error:
                        self.log_display.append(f"  -> 오류: {error}")
                    else:
                        self.log_display.append(f"  -> {result_msg}"); total_updates += 1 if "완료" in result_msg else 0

        if archive_file:
            try:
                source_file_path = self.file_path_entry.text().strip()
                base_archive_path = self.archive_path_entry.text().strip()
                destination_folder = os.path.join(base_archive_path, region_name)
                company_name_normalized = company_name.replace('㈜', '(주)')
                sanitized_company_name = re.sub(r'[<>:"/\\|?*]', '', company_name_normalized).strip()
                new_filename = f"{sanitized_company_name}_신용평가{os.path.splitext(source_file_path)[1]}"
                destination_path = os.path.join(destination_folder, new_filename)
                self.log_display.append("\n파일 보관 작업 중...")
                os.makedirs(destination_folder, exist_ok=True);
                shutil.move(source_file_path, destination_path)
                self.log_display.append(f"  -> '{new_filename}' 이름으로 변경하여\n  -> '{destination_folder}' 경로에 저장 완료!")
            except Exception as e:
                QMessageBox.critical(self, "파일 이동 오류", f"파일 이동 중 오류 발생:\n{e}")

        final_msg = "작업을 완료했습니다."
        if update_excel and not archive_file: final_msg = "엑셀 업데이트를 완료했습니다."
        QMessageBox.information(self, "작업 완료", f"{final_msg} 로그를 확인하세요.")
        self.reset_ui()

        self.update_button.setText("3. 업데이트 및 파일 보관")

    def set_data_input_enabled(self, enabled):
        self.fields_to_extract['신용평가등급']['entry'].setEnabled(enabled); self.start_date_edit.setEnabled(enabled); self.end_date_edit.setEnabled(enabled); self.update_button.setEnabled(True)

    def rotate_image(self, angle):
        """현재 이미지를 주어진 각도만큼 회전시킵니다."""
        if not self.original_pixmap:
            return

        # QTransform 객체를 사용하여 회전 적용
        transform = QTransform().rotate(angle)
        rotated_pixmap = self.original_pixmap.transformed(transform, Qt.TransformationMode.SmoothTransformation)

        # 원본 이미지를 회전된 이미지로 교체
        self.original_pixmap = rotated_pixmap

        # 현재 페이지 리스트에도 회전된 이미지를 반영 (PDF인 경우)
        if self.pdf_pages:
            self.pdf_pages[self.current_page_index] = self.original_pixmap

        # 화면을 새로고침하여 회전된 이미지를 표시
        self.fit_to_window()

    def toggle_file_inputs(self, state):
        """체크박스 상태에 따라 파일 관련 위젯을 활성화/비활성화합니다."""
        is_enabled = not bool(state)
        self.file_box.setEnabled(is_enabled)
        self.archive_box.setEnabled(is_enabled)
