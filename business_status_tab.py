# [business_status_tab.py 파일 전체를 이 코드로 교체하세요]

import sys
import os
import re
import shutil
import json
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel,
                               QLineEdit, QPushButton, QMessageBox, QFileDialog, QGroupBox, QScrollArea,
                               QTableWidget, QTableWidgetItem, QHeaderView, QComboBox, QInputDialog,
                               QDialog, QDialogButtonBox, QApplication, QCheckBox)
from PySide6.QtCore import QThread, Signal, Qt, QRect
from PySide6.QtGui import QPixmap, QImage, QColor
from PIL import Image
from PySide6.QtGui import QFont
import fitz  # PyMuPDF

# 우리 프로젝트의 다른 파일들
import ocr_logic
import config
import ocr_utils
from ui_widgets import ImageLabel, ZoomableScrollArea
from workers import RoiOcrWorker, ColorUpdateWorker
from PySide6.QtGui import QTransform


# --- 헬퍼 클래스 ---

class PdfExportDialog(QDialog):
    def __init__(self, max_page, parent=None):
        super().__init__(parent)
        self.setWindowTitle("페이지 내보내기")
        layout = QVBoxLayout(self)
        self.info_label = QLabel(f"내보낼 페이지를 입력하세요 (총 {max_page}페이지).\n(예: 1, 3-5, 8)")
        self.page_edit = QLineEdit()
        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.accept);
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.info_label);
        layout.addWidget(self.page_edit);
        layout.addWidget(self.button_box)

    def get_page_selection(self):
        return self.page_edit.text()


class ManualInputDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("신규 업체 정보 입력")
        layout = QVBoxLayout(self)
        self.info_label = QLabel("DB에서 업체를 찾을 수 없습니다.\n파일을 보관하려면 아래 정보를 직접 입력해주세요.")
        form_layout = QGridLayout();
        self.name_label = QLabel("업체명:");
        self.name_edit = QLineEdit();
        self.region_label = QLabel("지역:");
        self.region_combo = QComboBox()
        self.region_combo.addItems(
            ["서울", "경기", "인천", "부산", "대구", "광주", "대전", "울산", "세종", "강원", "충북", "충남", "전북", "전남", "경북", "경남", "제주"])
        form_layout.addWidget(self.name_label, 0, 0);
        form_layout.addWidget(self.name_edit, 0, 1);
        form_layout.addWidget(self.region_label, 1, 0);
        form_layout.addWidget(self.region_combo, 1, 1)
        self.button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.button_box.accepted.connect(self.accept);
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.info_label);
        layout.addLayout(form_layout);
        layout.addWidget(self.button_box)

    def get_data(self): return self.name_edit.text().strip(), self.region_combo.currentText()


class CreditColorUpdateWorker(QThread):
    finished = Signal(str)

    def __init__(self, excel_path): super().__init__(); self.excel_path = excel_path

    def run(self): self.finished.emit(ocr_logic.batch_update_credit_rating_colors(self.excel_path))


# --- 헬퍼 함수 ---

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
    except:
        return None
    return sorted(list(pages))


# --- BusinessStatusTab 클래스 ---

class BusinessStatusTab(QWidget):
    def __init__(self, reader):
        super().__init__()
        self.reader = reader
        self.original_pixmap = None
        self.scale_factor = 1.0
        self.fields_to_extract = {key: {} for key in config.COLUMN_MAP.keys()}
        self.current_field_to_set = None
        self.current_company_name = None
        self.current_before_data = None
        self.excel_paths = {"전기": "", "통신": "", "소방": ""}
        self.pdf_pages = []
        self.current_page_index = 0
        self.setup_ui()
        self.connect_signals()
        self.load_excel_paths()
        print("BusinessStatusTab 객체 생성 완료")

    def setup_ui(self):
        main_layout = QHBoxLayout(self)
        viewer_panel = self.create_viewer_panel()
        preview_panel = self.create_preview_panel()
        controls_panel = self.create_controls_panel()
        main_layout.addWidget(viewer_panel, 3)
        main_layout.addWidget(preview_panel, 2)
        main_layout.addWidget(controls_panel)

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

    def create_preview_table(self):
        table = QTableWidget()
        table.setColumnCount(2)
        table.setRowCount(len(self.fields_to_extract))
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setVisible(False)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)

        for row, field_name in enumerate(self.fields_to_extract.keys()):
            item = QTableWidgetItem(field_name)
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
            table.setItem(row, 0, item)
        return table



    def create_preview_panel(self):
        panel = QGroupBox("4. 변경 전/후 미리보기")
        layout = QHBoxLayout(panel)
        self.before_table = self.create_preview_table()
        self.after_table = self.create_preview_table()

        before_vbox = QVBoxLayout()
        before_vbox.addWidget(QLabel("<b>변경 전 (엑셀 원본)</b>"))
        # [핵심] 테이블 위젯이 세로로 남는 공간을 모두 차지하도록 stretch 값(1)을 추가
        before_vbox.addWidget(self.before_table, 1)

        after_vbox = QVBoxLayout()
        after_vbox.addWidget(QLabel("<b>변경 후 (OCR 결과)</b>"))
        # [핵심] 테이블 위젯이 세로로 남는 공간을 모두 차지하도록 stretch 값(1)을 추가
        after_vbox.addWidget(self.after_table, 1)

        layout.addLayout(before_vbox)
        layout.addLayout(after_vbox)
        return panel

    def create_controls_panel(self):
        panel = QWidget();
        layout = QVBoxLayout(panel);
        panel.setFixedWidth(450)

        # 1. 원본 파일 선택 (self.file_box로 변경)
        self.file_box = QGroupBox("1. 원본 파일 선택");
        file_layout = QHBoxLayout(self.file_box)
        self.file_path_entry = QLineEdit();
        self.file_path_entry.setReadOnly(True)
        self.file_select_button = QPushButton("📁 PDF/이미지 열기");
        file_layout.addWidget(self.file_path_entry);
        file_layout.addWidget(self.file_select_button)

        # 2. 업데이트 대상 설정
        excel_box = QGroupBox("2. 업데이트 대상 설정");
        excel_layout = QGridLayout(excel_box)
        self.file_type_combo = QComboBox();
        self.file_type_combo.addItems(["-- 자료 종류 선택 --", "전기경영상태", "통신경영상태", "소방경영상태"])
        self.excel_file_path_entry = QLineEdit();
        self.excel_file_path_entry.setReadOnly(True);
        self.excel_file_path_entry.setPlaceholderText("자료 종류 선택 시 자동 지정")
        self.excel_path_config_button = QPushButton("🔧 DB 경로 설정");
        self.color_update_button = QPushButton("🎨 연말 색상 업데이트");
        self.credit_color_update_button = QPushButton("✨ 신용평가 유효기간 갱신")
        excel_layout.addWidget(QLabel("자료 종류:"), 0, 0);
        excel_layout.addWidget(self.file_type_combo, 0, 1, 1, 2)
        excel_layout.addWidget(QLabel("DB 경로:"), 1, 0);
        excel_layout.addWidget(self.excel_file_path_entry, 1, 1);
        excel_layout.addWidget(self.excel_path_config_button, 1, 2)
        excel_layout.addWidget(self.color_update_button, 2, 0, 1, 3);
        excel_layout.addWidget(self.credit_color_update_button, 3, 0, 1, 3)

        # 3. 데이터 영역 지정
        roi_box = QGroupBox("3. 데이터 영역 지정");
        roi_layout = QGridLayout(roi_box)
        for row, field in enumerate(self.fields_to_extract.keys()):
            lbl, btn, entry = QLabel(f"{field}:"), QPushButton("지정"), QLineEdit();
            btn.setProperty("field_name", field)
            roi_layout.addWidget(lbl, row, 0);
            roi_layout.addWidget(btn, row, 1);
            roi_layout.addWidget(entry, row, 2)
            self.fields_to_extract[field].update({"roi": None, "entry": entry, "button": btn})

        # 4. 처리 완료 파일 보관 경로 (self.archive_box로 변경)
        self.archive_box = QGroupBox("4. 처리 완료 파일 보관 경로");
        archive_layout = QGridLayout(self.archive_box)
        self.archive_path_entry = QLineEdit();
        self.archive_path_entry.setPlaceholderText("자료를 보관할 최상위 폴더를 선택하세요.")
        self.archive_select_button = QPushButton("📂 기본 경로 선택");
        archive_layout.addWidget(QLabel("기본 보관 경로:"), 0, 0)
        archive_layout.addWidget(self.archive_path_entry, 0, 1);
        archive_layout.addWidget(self.archive_select_button, 0, 2)

        # 5. 실행
        action_box = QGroupBox("5. 실행");
        action_layout = QVBoxLayout(action_box)
        self.run_ocr_button = QPushButton("1. 지정 영역 분석");
        self.compare_button = QPushButton("2. 원본 데이터 비교")

        # [핵심] 체크박스 생성
        self.data_only_checkbox = QCheckBox("자료 파일 없이 데이터만 저장")

        self.save_button = QPushButton("3. 확정 및 엑셀 저장");
        self.save_button.setEnabled(False);
        self.save_button.setStyleSheet("font-weight: bold; background-color: #A93226;")
        action_layout.addWidget(self.run_ocr_button);
        action_layout.addWidget(self.compare_button)
        action_layout.addWidget(self.data_only_checkbox)  # 체크박스 추가
        action_layout.addWidget(self.save_button)

        layout.addWidget(self.file_box);
        layout.addWidget(excel_box);
        layout.addWidget(roi_box)
        layout.addWidget(self.archive_box);
        layout.addStretch(1);
        layout.addWidget(action_box)
        return panel

    def connect_signals(self):
        self.prev_page_button.clicked.connect(self.show_previous_page)
        self.next_page_button.clicked.connect(self.show_next_page)
        self.export_pdf_button.clicked.connect(self.export_pdf_pages)
        self.rotate_left_button.clicked.connect(lambda: self.rotate_image(-90))
        self.rotate_right_button.clicked.connect(lambda: self.rotate_image(90))
        self.data_only_checkbox.stateChanged.connect(self.toggle_file_inputs)

        # 각 데이터 영역 지정 버튼에 대한 시그널 연결
        for key, field_data in self.fields_to_extract.items():
            field_data['button'].clicked.connect(self.prepare_to_set_roi)

            # [핵심 수정] 필드 종류에 따라 다른 서식 함수를 연결
            if key in ['시평액', '3년실적', '5년실적']:
                field_data['entry'].textChanged.connect(self.format_number_input)
            elif key in ['부채비율', '유동비율']:
                field_data['entry'].textChanged.connect(self.format_ratio_input)

        self.file_select_button.clicked.connect(lambda: self.open_file())
        self.excel_path_config_button.clicked.connect(self.configure_excel_paths)
        self.archive_select_button.clicked.connect(self.select_archive_folder)
        self.image_label.roi_selected.connect(self.on_roi_selected)
        self.zoom_in_button.clicked.connect(lambda: self.zoom_image(1.2))
        self.zoom_out_button.clicked.connect(lambda: self.zoom_image(0.8))
        self.zoom_fit_button.clicked.connect(self.fit_to_window)
        self.run_ocr_button.clicked.connect(self.run_roi_ocr)
        self.compare_button.clicked.connect(self.compare_data)
        self.save_button.clicked.connect(self.save_data_to_excel)
        self.file_type_combo.currentTextChanged.connect(self.on_file_type_changed)
        self.color_update_button.clicked.connect(self.start_color_update)
        self.credit_color_update_button.clicked.connect(self.start_credit_color_update)

    def open_file(self, file_path=None):
        print("'PDF/이미지 열기' 버튼 클릭됨! 파일 선택창을 엽니다...")
        if file_path is None:
            file_path, _ = QFileDialog.getOpenFileName(self, "파일 선택", "", "PDF 및 이미지 파일 (*.pdf *.png *.jpg *.jpeg)")
        if not file_path: return
        self.reset_ui_for_next_file()
        self.file_path_entry.setText(file_path)

        try:
            self.pdf_pages = []
            if file_path.lower().endswith('.pdf'):
                doc = fitz.open(file_path)
                temp_files_to_delete = []  # 삭제할 임시 파일 목록


                for page in doc:
                    pix = page.get_pixmap(dpi=300)
                    img_pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                    tab_name = "bs" if isinstance(self, BusinessStatusTab) else "cr"
                    temp_image_path = f"temp_page_{tab_name}_{page.number}.png"

                    img_pil.save(temp_image_path)
                    self.pdf_pages.append(QPixmap(temp_image_path))
                    temp_files_to_delete.append(temp_image_path)

                doc.close()

                # [핵심 수정] 사용이 끝난 임시 파일들을 삭제
                for temp_file in temp_files_to_delete:
                    if os.path.exists(temp_file):
                        os.remove(temp_file)

                if not self.pdf_pages:
                    self.original_pixmap = None
                else:
                    self.original_pixmap = self.pdf_pages[0]
                    if len(self.pdf_pages) > 1:
                        QMessageBox.information(self, "알림",
                                                f"이 PDF는 총 {len(self.pdf_pages)}페이지입니다.\n아래 페이지 이동 버튼으로 모든 페이지를 확인하세요.")
            else:
                self.original_pixmap = QPixmap(file_path)
            if self.original_pixmap is None or self.original_pixmap.isNull():
                self.image_label.clear();
                self.set_page_controls_visibility(False);
                return
            self.display_page(0);
            self.fit_to_window()
        except Exception as e:
            QMessageBox.critical(self, "파일 열기 오류", f"파일을 여는 중 오류가 발생했습니다:\n{e}")

    def display_page(self, page_index):
        if not self.pdf_pages and self.original_pixmap:  # 일반 이미지 파일 처리
            self.image_label.setPixmap(self.original_pixmap)
            self.set_page_controls_visibility(False)
            return

        if 0 <= page_index < len(self.pdf_pages):
            self.current_page_index = page_index
            self.original_pixmap = self.pdf_pages[self.current_page_index]
            self.fit_to_window()

            self.page_label.setText(f"{self.current_page_index + 1} / {len(self.pdf_pages)}")
            self.prev_page_button.setEnabled(self.current_page_index > 0)
            self.next_page_button.setEnabled(self.current_page_index < len(self.pdf_pages) - 1)

            # [핵심 수정] 페이지 컨트롤 UI가 보이게 하는 조건을 (페이지 수 > 0) 으로 변경
            self.set_page_controls_visibility(len(self.pdf_pages) > 0)

    def show_previous_page(self):
        self.display_page(self.current_page_index - 1)

    def show_next_page(self):
        self.display_page(self.current_page_index + 1)

    def set_page_controls_visibility(self, visible):
        self.prev_page_button.setVisible(visible);
        self.next_page_button.setVisible(visible)
        self.page_label.setVisible(visible);
        self.export_pdf_button.setVisible(visible)

    # [business_status_tab.py와 credit_rating_tab.py 두 파일 모두 이 함수로 교체]

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

    def compare_data(self):
        print(">>> 새 버전의 compare_data 함수 실행됨 <<<")  # <-- 진단용 메시지 추가
        excel_path = self.excel_file_path_entry.text()
        biz_no = self.fields_to_extract['사업자등록번호']['entry'].text().strip()
        if not (excel_path and biz_no):
            QMessageBox.warning(self, "오류", "업데이트할 DB(자료 종류)와 사업자등록번호가 모두 필요합니다.");
            return

        before_data, error = ocr_logic.find_company_data(excel_path, biz_no)

        # 1. 엑셀 파일 자체에 오류가 있는지 먼저 확인
        if error:
            QMessageBox.critical(self, "엑셀 파일 오류", error)
            self.save_button.setEnabled(False)
            return

        # 2. 파일에 문제가 없을 때, 업체 존재 여부 확인
        if before_data:
            # [기존 업체 처리 로직]
            original_biz_no = before_data.get('사업자등록번호', {}).get('value')
            if original_biz_no:
                self.fields_to_extract['사업자등록번호']['entry'].setText(str(original_biz_no))

            self.current_before_data = before_data
            self.current_company_name = before_data.get('상호', {}).get('value')
            after_data = {key: info.get('value') for key, info in before_data.items()}
            for key, field_data in self.fields_to_extract.items():
                ui_text = field_data['entry'].text().strip()
                if ui_text: after_data[key] = ui_text
            self.populate_preview_table(self.before_table, before_data, is_after=False)
            self.populate_preview_table(self.after_table, after_data, is_after=True)
            self.save_button.setEnabled(True)
            QMessageBox.information(self, "비교 완료", "내용을 확인하고 '3. 확정 및 엑셀 저장' 버튼을 누르세요.")

        else:
            # [신규 업체 처리 로직]
            self.current_before_data = None
            self.before_table.clearContents()
            self.after_table.clearContents()

            dialog = ManualInputDialog(self)
            if dialog.exec():
                manual_name, manual_region = dialog.get_data()
                if not (manual_name and manual_region):
                    QMessageBox.warning(self, "정보 부족", "업체명과 지역을 모두 입력해야 합니다.")
                    self.save_button.setEnabled(False)
                else:
                    self.current_company_name = manual_name
                    self.current_before_data = {'지역': {'value': manual_region}}
                    QMessageBox.information(self, "신규 업체 확인",
                                            "입력된 정보로 파일 보관을 준비합니다.\n'3. 확정 및 엑셀 저장' 버튼을 눌러 파일을 이동하세요.")
                    self.save_button.setEnabled(True)
            else:
                self.save_button.setEnabled(False)

    def save_data_to_excel(self):
        # --- 정보 수집 ---
        update_data = {k: v['entry'].text() for k, v in self.fields_to_extract.items()}
        biz_no = update_data.get('사업자등록번호', '').strip()
        file_type = self.file_type_combo.currentText()
        excel_key = "전기" if "전기" in file_type else "통신" if "통신" in file_type else "소방" if "소방" in file_type else None
        excel_path = self.excel_paths.get(excel_key) if excel_key else None

        # [핵심 수정] 체크박스 상태에 따라 유효성 검사 및 작업 분기
        data_only_mode = self.data_only_checkbox.isChecked()

        # --- 데이터만 저장 모드 ---
        if data_only_mode:
            if not (biz_no and excel_path):
                QMessageBox.warning(self, "정보 부족", "DB 경로와 사업자등록번호가 모두 필요합니다.");
                return

            reply = QMessageBox.question(self, "최종 확인",
                                         f"<b>[엑셀 업데이트]</b>\n- 대상 파일: {os.path.basename(excel_path)}\n\n데이터만 업데이트하시겠습니까?",
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.Yes:
                updated_log, error = ocr_logic.update_company_data(excel_path, biz_no, update_data, excel_key)
                if error: QMessageBox.critical(self, "엑셀 업데이트 오류", error); return
                QMessageBox.information(self, "작업 완료", "엑셀 업데이트를 완료했습니다.")
                self.reset_ui_for_next_file()
                self.save_button.setEnabled(False)

        # --- 기존 모드 (파일 보관 포함) ---
        else:
            source_file_path = self.file_path_entry.text()
            base_archive_path = self.archive_path_entry.text()

            if not (biz_no and source_file_path and base_archive_path):
                QMessageBox.warning(self, "정보 부족", "자료 파일, 보관 경로, 사업자등록번호가 모두 필요합니다.");
                return
            if file_type == "-- 자료 종류 선택 --":
                QMessageBox.warning(self, "종류 선택 필요", "'자료 종류'를 먼저 선택해주세요.");
                return
            if not self.current_company_name or not self.current_before_data:
                QMessageBox.warning(self, "정보 오류", "'2. 원본 데이터 비교'를 먼저 실행하여 업체 정보를 확인해주세요.");
                return

            # (이하 기존 최종 확인창 및 파일 보관 로직과 동일)
            is_existing_company, _ = ocr_logic.find_company_data(excel_path, biz_no) if excel_path and os.path.exists(
                excel_path) else (False, None)
            try:
                region_info_dict = self.current_before_data.get('지역', {});
                region_full_name = region_info_dict.get('value', '기타');
                region_name = region_full_name.split(' ')[0] if region_full_name else '기타';
                destination_folder = os.path.join(base_archive_path, region_name)
                company_name_normalized = self.current_company_name.replace('㈜', '(주)');
                sanitized_company_name = re.sub(r'[<>:"/\\|?*]', '', company_name_normalized).strip();
                _, file_extension = os.path.splitext(source_file_path);
                new_filename = f"{sanitized_company_name}_{file_type}{file_extension}"
            except Exception as e:
                QMessageBox.critical(self, "경로 생성 오류", f"파일 저장 경로를 만드는 중 오류가 발생했습니다:\n{e}"); return
            confirm_message = (
                f"아래 작업을 실행하시겠습니까?\n\n<b>[엑셀 업데이트]</b>\n- 대상 파일: {os.path.basename(excel_path) if is_existing_company else '없음 (신규 업체)'}\n\n<b>[파일 보관]</b>\n- 새 이름: {new_filename}\n- 저장 위치: {destination_folder}");
            reply = QMessageBox.question(self, "최종 확인", confirm_message,
                                         QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                         QMessageBox.StandardButton.No)
            if reply == QMessageBox.StandardButton.No: return

            if is_existing_company:
                updated_log, error = ocr_logic.update_company_data(excel_path, biz_no, update_data, excel_key)
                if error: QMessageBox.critical(self, "엑셀 업데이트 오류", error); return
                self.archive_file(destination_folder, new_filename)
            else:
                QMessageBox.information(self, "파일 보관 실행", "신규 업체로 인식되어 파일 보관만 실행합니다.");
                self.archive_file(destination_folder, new_filename)

    def archive_file(self, destination_folder, new_filename):
        try:
            source_file_path = self.file_path_entry.text();
            destination_path = os.path.join(destination_folder, new_filename)
            os.makedirs(destination_folder, exist_ok=True);
            shutil.move(source_file_path, destination_path)
            QMessageBox.information(self, "작업 완료",
                                    f"<b>[파일 보관 완료]</b>\n'{os.path.basename(source_file_path)}' 파일을\n'{os.path.basename(destination_path)}'(으)로 변경하여 저장했습니다.")
            self.reset_ui_for_next_file()
        except Exception as e:
            QMessageBox.critical(self, "파일 보관 오류", f"파일을 보관하는 중 오류가 발생했습니다:\n{e}")
        self.save_button.setEnabled(False)

    def format_number_input(self, text):
        sender = self.sender()
        if not isinstance(sender, QLineEdit): return

        # [핵심 수정] 입력된 텍스트가 숫자인지 먼저 확인
        number_str = re.sub(r'[^0-9]', '', text)
        if not number_str:  # 비어있으면 그대로 둠
            return

        # 입력된 텍스트 전체가 숫자로만 구성되지 않았다면 (예: "판단안됨")
        # 서식을 적용하지 않고 그대로 둠
        if not text.replace(',', '').isdigit():
            return

        try:
            number = int(number_str)
            formatted_text = f"{number:,}"
        except ValueError:
            formatted_text = text

        if text != formatted_text:
            sender.blockSignals(True)
            sender.setText(formatted_text)
            sender.blockSignals(False)
            sender.end(False)

    def format_ratio_input(self, text):
        sender = self.sender()
        if not isinstance(sender, QLineEdit): return

        # [핵심 수정] 입력된 텍스트가 숫자인지 먼저 확인 (소수점은 허용)
        cleaned_text = text.replace('.', '')
        if not cleaned_text.isdigit() and cleaned_text != "":
            return  # 숫자가 아니면 서식 적용 안함

        number_str = re.sub(r'[^0-9]', '', text)
        if not number_str:
            return

        if len(number_str) > 2:
            formatted_text = f"{number_str[:-2]}.{number_str[-2:]}"
        else:
            formatted_text = number_str

        if text != formatted_text:
            sender.blockSignals(True)
            sender.setText(formatted_text)
            sender.blockSignals(False)
            sender.end(False)


    def reset_ui_for_next_file(self):
        self.file_path_entry.clear()
        self.original_pixmap = None;
        self.image_label.clear()
        self.pdf_pages = [];
        self.current_page_index = 0;
        self.display_page(0)
        for field in self.fields_to_extract.values():
            if field.get('button'): field['button'].setText("지정"); field['button'].setStyleSheet("")
            if field.get('entry'): field['entry'].clear()
            field['roi'] = None
        self.before_table.clearContents();
        self.after_table.clearContents()
        self.file_type_combo.setCurrentIndex(0);
        self.save_button.setEnabled(False)

    def select_archive_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "보관할 폴더 선택")
        if folder_path:
            self.archive_path_entry.setText(folder_path)

    def zoom_image(self, factor):
        if self.original_pixmap: self.scale_factor *= factor; new_width = int(
            self.original_pixmap.width() * self.scale_factor); scaled_pixmap = self.original_pixmap.scaledToWidth(
            new_width, Qt.SmoothTransformation); self.image_label.setPixmap(scaled_pixmap); self.zoom_label.setText(
            f"{int(self.scale_factor * 100)}%")

    def fit_to_window(self):
        if self.original_pixmap: scaled_pixmap = self.original_pixmap.scaled(self.scroll_area.viewport().size(),
                                                                             Qt.KeepAspectRatio,
                                                                             Qt.SmoothTransformation); self.image_label.setPixmap(
            scaled_pixmap)
        if self.original_pixmap and self.original_pixmap.width() > 0: self.scale_factor = self.image_label.pixmap().width() / self.original_pixmap.width(); self.zoom_label.setText(
            f"{int(self.scale_factor * 100)}%")

    def prepare_to_set_roi(self):
        sender = self.sender(); self.current_field_to_set = sender.property(
            "field_name"); self.image_label.selecting = True; self.setCursor(
            Qt.CrossCursor); self.image_label.setCursor(Qt.CrossCursor)

    def on_roi_selected(self, rect):
        if self.current_field_to_set: original_rect = QRect(int(rect.x() / self.scale_factor),
                                                            int(rect.y() / self.scale_factor),
                                                            int(rect.width() / self.scale_factor),
                                                            int(rect.height() / self.scale_factor));
        self.fields_to_extract[self.current_field_to_set]['roi'] = original_rect; button = \
        self.fields_to_extract[self.current_field_to_set]['button']; button.setText(
            f"지정됨({original_rect.x()},{original_rect.y()})"); button.setStyleSheet("background-color: #2ECC71;")
        self.image_label.selecting = False;
        self.current_field_to_set = None;
        self.setCursor(Qt.ArrowCursor);
        self.image_label.setCursor(Qt.ArrowCursor)

    def run_roi_ocr(self):
        if not self.original_pixmap: QMessageBox.warning(self, "오류", "먼저 분석할 이미지를 열어주세요."); return
        fields_to_process = {k: v for k, v in self.fields_to_extract.items() if v.get('roi')};
        if not fields_to_process: QMessageBox.warning(self, "오류", "하나 이상의 영역을 먼저 지정해주세요."); return
        self.run_ocr_button.setEnabled(False);
        self.run_ocr_button.setText("분석 중...");
        current_qimage = self.original_pixmap.toImage();
        self.worker = RoiOcrWorker(self.reader, current_qimage, fields_to_process);
        self.worker.progress.connect(self.update_ocr_result);
        self.worker.finished.connect(self.on_ocr_finished);
        self.worker.start()

    def update_ocr_result(self, field_name, text):
        cleaned_text = text;
        if field_name == '사업자등록번호':
            cleaned_text = ocr_utils.clean_biz_number(text)
        elif '실적' in field_name or '시평액' in field_name:
            cleaned_text = ocr_utils.clean_ocr_number(text)
        elif '비율' in field_name:
            cleaned_text = "".join(re.findall(r'[\d.]', text))
        self.fields_to_extract[field_name]['entry'].setText(cleaned_text)

    def on_ocr_finished(self, message):
        self.run_ocr_button.setEnabled(True); self.run_ocr_button.setText("1. 지정 영역 분석"); QMessageBox.information(self,
                                                                                                                  "분석 완료",
                                                                                                                  message)

    def populate_preview_table(self, table, data, is_after=False):
        if not table: return

        from PySide6.QtGui import QFont, QColor
        data_font = QFont()
        data_font.setPointSize(11)

        for row, key in enumerate(self.fields_to_extract.keys()):
            label_item = QTableWidgetItem(key)
            label_item.setFlags(label_item.flags() & ~Qt.ItemIsEditable)
            table.setItem(row, 0, label_item)

            value = None
            color_hex = '#FFFFFF'

            if not is_after:  # '변경 전' 패널
                cell_info = data.get(key, {})
                value = cell_info.get('value')
                color_hex = cell_info.get('color', '#FFFFFF')
            else:  # '변경 후' 패널
                value = data.get(key)

            display_text = ""
            if value is not None and str(value).strip() != "":
                try:
                    # [핵심 수정] is_after 플래그에 따라 숫자 변환 로직을 분리
                    if '비율' in key:
                        numeric_value = float(str(value).replace('%', ''))
                        # '변경 후' 데이터는 UI 입력값(e.g., "12.23")이거나 DB 원본(e.g., 0.1223)일 수 있음
                        if is_after and isinstance(value, str):
                            display_text = f"{numeric_value:.2f}%"
                        else:  # '변경 전' 데이터 또는 '변경 후'의 DB 원본 데이터는 항상 100을 곱해야 함
                            display_text = f"{numeric_value * 100:.2f}%"

                    elif '실적' in key or '시평액' in key:
                        numeric_value = int(float(str(value).replace(',', '')))
                        # '변경 후' 데이터는 UI 입력값(단위: 천원, e.g., "6042281")
                        if is_after and isinstance(value, str):
                            display_text = f"{numeric_value * 1000:,}"
                        else:  # '변경 전' 데이터 또는 '변경 후'의 DB 원본 데이터 (단위: 원, e.g., 6042281000)
                            display_text = f"{numeric_value:,}"

                    else:
                        display_text = str(value)
                except (ValueError, TypeError):
                    display_text = str(value)

            value_item = QTableWidgetItem(display_text)
            bg_color = QColor(color_hex)
            if bg_color.lightnessF() > 0.9: value_item.setForeground(QColor(0, 0, 0))
            value_item.setBackground(bg_color)
            value_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            value_item.setFont(data_font)
            table.setItem(row, 1, value_item)

        table.resizeRowsToContents()


    def start_color_update(self):
        excel_path = self.excel_file_path_entry.text()
        if not excel_path: QMessageBox.warning(self, "파일 선택 오류", "먼저 색상을 업데이트할 엑셀 파일을 선택해주세요."); return
        reply = QMessageBox.question(self, "연말 색상 업데이트 확인",
                                     f"'{os.path.basename(excel_path)}' 파일의 모든 데이터 상태 색상을 갱신하시겠습니까?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes: self.color_update_button.setText(
            "업데이트 중..."); self.color_update_button.setEnabled(False); self.color_worker = ColorUpdateWorker(
            excel_path); self.color_worker.finished.connect(self.on_color_update_finished); self.color_worker.start()

    def on_color_update_finished(self, message):
        self.color_update_button.setText("🎨 연말 색상 업데이트");
        self.color_update_button.setEnabled(True)
        if "오류" in message:
            QMessageBox.critical(self, "업데이트 오류", message)
        else:
            QMessageBox.information(self, "업데이트 완료", message)

    def start_credit_color_update(self):
        excel_path = self.excel_file_path_entry.text()
        if not excel_path: QMessageBox.warning(self, "파일 선택 오류", "먼저 색상을 업데이트할 엑셀 파일을 선택해주세요."); return
        reply = QMessageBox.question(self, "신용평가 유효기간 갱신 확인",
                                     f"'{os.path.basename(excel_path)}' 파일의 모든 '신용평가' 셀의 색상을 유효기간에 따라 갱신하시겠습니까?",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes: self.credit_color_update_button.setText(
            "갱신 중..."); self.credit_color_update_button.setEnabled(False); self.credit_worker = CreditColorUpdateWorker(
            excel_path); self.credit_worker.finished.connect(
            self.on_credit_color_update_finished); self.credit_worker.start()

    def on_credit_color_update_finished(self, message):
        self.credit_color_update_button.setText("✨ 신용평가 유효기간 갱신");
        self.credit_color_update_button.setEnabled(True)
        if "오류" in message:
            QMessageBox.critical(self, "갱신 오류", message)
        else:
            QMessageBox.information(self, "갱신 완료", message)

    def load_excel_paths(self):
        try:
            if os.path.exists("ocr_config.json"):
                with open("ocr_config.json", 'r', encoding='utf-8') as f: self.excel_paths.update(json.load(f))
        except Exception as e:
            print(f"설정 파일 로드 오류: {e}")

    def save_excel_paths(self):
        try:
            with open("ocr_config.json", 'w', encoding='utf-8') as f:
                json.dump(self.excel_paths, f, ensure_ascii=False, indent=4)
        except Exception as e:
            QMessageBox.critical(self, "설정 저장 오류", f"설정 파일 저장 중 오류가 발생했습니다:\n{e}")

    def configure_excel_paths(self):
        items = ["전기", "통신", "소방"];
        item, ok = QInputDialog.getItem(self, "엑셀 경로 설정", "어떤 DB 파일의 경로를 설정하시겠습니까?", items, 0, False)
        if ok and item:
            file_path, _ = QFileDialog.getOpenFileName(self, f"{item} DB 파일 선택", "", "Excel 파일 (*.xlsx *.xls)")
            if file_path: self.excel_paths[item] = file_path; self.save_excel_paths(); self.on_file_type_changed(
                self.file_type_combo.currentText()); QMessageBox.information(self, "설정 완료",
                                                                             f"'{item}' DB의 경로가 저장되었습니다.")

    def on_file_type_changed(self, text):
        if "전기" in text:
            key = "전기"
        elif "통신" in text:
            key = "통신"
        elif "소방" in text:
            key = "소방"
        else:
            key = None
        if key and self.excel_paths.get(key):
            self.excel_file_path_entry.setText(self.excel_paths[key])
        else:
            self.excel_file_path_entry.clear()

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