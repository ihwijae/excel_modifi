from __future__ import annotations

import os
from collections import OrderedDict
from typing import Dict, List, Optional

from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QGridLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QComboBox,
    QFileDialog,
    QHBoxLayout,
    QProgressBar,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QMessageBox,
)
from PySide6.QtCore import Qt

import smpp_excel
from smpp_excel import CompanyEntry
from smpp_service import CorpCheckResult
from workers import SmppLookupWorker


class SmppLookupDialog(QDialog):
    def __init__(
        self,
        excel_paths: Optional[Dict[str, str]] = None,
        default_excel_path: str = "",
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle("SMPP 중소/여성기업 조회")
        self.resize(900, 600)

        self.excel_paths = excel_paths or {}
        self.selected_excel_path = default_excel_path
        self.smpp_worker = None
        self.smpp_result_map: "OrderedDict[str, Dict[str, str]]" = OrderedDict()
        self.biz_numbers: List[str] = []
        self.company_entries: List[CompanyEntry] = []

        self._build_ui()
        self._populate_excel_sources()
        if self.selected_excel_path:
            self.excel_path_entry.setText(self.selected_excel_path)

    def _build_ui(self):
        layout = QVBoxLayout(self)

        # Excel 경로 선택
        path_group = QGridLayout()
        self.db_combo = QComboBox()
        self.excel_path_entry = QLineEdit()
        self.excel_path_entry.setReadOnly(True)
        self.browse_button = QPushButton("파일 선택")
        path_group.addWidget(QLabel("DB 선택:"), 0, 0)
        path_group.addWidget(self.db_combo, 0, 1)
        path_group.addWidget(self.browse_button, 0, 2)
        path_group.addWidget(QLabel("선택된 파일:"), 1, 0)
        path_group.addWidget(self.excel_path_entry, 1, 1, 1, 2)
        self.load_biz_button = QPushButton("사업자번호 가져오기")
        path_group.addWidget(self.load_biz_button, 2, 0, 1, 3)
        layout.addLayout(path_group)
        self.biz_count_label = QLabel("총 0건 로드됨")
        layout.addWidget(self.biz_count_label)

        # 업체명 검색
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("업체명 검색:"))
        self.search_entry = QLineEdit()
        self.search_entry.setPlaceholderText("예) 거성전력")
        search_layout.addWidget(self.search_entry, 1)
        self.clear_search_button = QPushButton("초기화")
        search_layout.addWidget(self.clear_search_button)
        layout.addLayout(search_layout)

        # SMPP 자격 정보
        creds_layout = QGridLayout()
        self.smpp_id_entry = QLineEdit()
        self.smpp_password_entry = QLineEdit()
        self.smpp_password_entry.setEchoMode(QLineEdit.Password)
        creds_layout.addWidget(QLabel("SMPP ID"), 0, 0)
        creds_layout.addWidget(self.smpp_id_entry, 0, 1)
        creds_layout.addWidget(QLabel("SMPP PW"), 1, 0)
        creds_layout.addWidget(self.smpp_password_entry, 1, 1)
        layout.addLayout(creds_layout)

        # 실행 버튼
        self.smpp_lookup_button = QPushButton("여성/중소기업 유효기간 조회")
        layout.addWidget(self.smpp_lookup_button)

        # 진행 상태
        progress_layout = QHBoxLayout()
        self.smpp_progress_bar = QProgressBar()
        self.smpp_status_label = QLabel("대기 중")
        progress_layout.addWidget(self.smpp_progress_bar, 1)
        progress_layout.addWidget(self.smpp_status_label)
        layout.addLayout(progress_layout)

        # 결과 테이블
        self.smpp_results_table = QTableWidget()
        self.smpp_results_table.setColumnCount(5)
        self.smpp_results_table.setHorizontalHeaderLabels(
            ["업체명", "대표자명", "사업자등록번호", "여성기업", "중소기업"]
        )
        self.smpp_results_table.verticalHeader().setVisible(False)
        header = self.smpp_results_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        layout.addWidget(self.smpp_results_table, 1)

        # 닫기 버튼
        self.close_button = QPushButton("닫기")
        layout.addWidget(self.close_button, alignment=Qt.AlignRight)

        # 시그널 연결
        self.db_combo.currentIndexChanged.connect(self._on_db_combo_changed)
        self.browse_button.clicked.connect(self._browse_excel_file)
        self.load_biz_button.clicked.connect(self.load_biz_numbers)
        self.search_entry.textChanged.connect(self.apply_company_filter)
        self.clear_search_button.clicked.connect(lambda: self.search_entry.clear())
        self.smpp_lookup_button.clicked.connect(self.start_smpp_lookup)
        self.close_button.clicked.connect(self.accept)

    def _populate_excel_sources(self):
        self.db_combo.clear()
        self.db_combo.addItem("-- DB 선택 --", "")
        for key, path in self.excel_paths.items():
            if path:
                self.db_combo.addItem(f"{key} ({os.path.basename(path)})", path)

    def _on_db_combo_changed(self, index: int):
        path = self.db_combo.itemData(index) or ""
        if path:
            self.selected_excel_path = path
            self.excel_path_entry.setText(path)

    def _browse_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "엑셀 파일 선택",
            "",
            "Excel Files (*.xlsx *.xls)",
        )
        if file_path:
            self.selected_excel_path = file_path
            self.excel_path_entry.setText(file_path)

    def load_biz_numbers(self):
        excel_path = self.selected_excel_path or self.excel_path_entry.text().strip()
        if not excel_path:
            QMessageBox.warning(self, "DB 경로", "엑셀 파일을 먼저 선택해주세요.")
            return
        try:
            entries = smpp_excel.extract_company_entries_from_excel(excel_path)
        except Exception as exc:
            QMessageBox.critical(self, "사업자번호 추출 오류", str(exc))
            return
        if not entries:
            QMessageBox.warning(self, "사업자번호 없음", "선택한 파일에서 사업자등록번호를 찾지 못했습니다.")
            return

        self.selected_excel_path = excel_path
        self.company_entries = entries
        self.biz_numbers = [entry.biz_no for entry in entries]
        self.smpp_result_map.clear()
        for entry in self.company_entries:
            biz = entry.biz_no
            self.smpp_result_map[biz] = {"women": "", "small": "", "error": ""}
        self.refresh_smpp_results_table()
        self.smpp_progress_bar.setMaximum(len(self.biz_numbers))
        self.smpp_progress_bar.setValue(0)
        self.smpp_status_label.setText("사업자번호 로드 완료")
        self.biz_count_label.setText(f"총 {len(self.biz_numbers)}건 로드됨")
        QMessageBox.information(
            self,
            "사업자번호 가져오기",
            f"{len(self.biz_numbers)}건의 사업자번호를 불러왔습니다.",
        )

    def start_smpp_lookup(self):
        if self.smpp_worker and self.smpp_worker.isRunning():
            QMessageBox.warning(self, "진행 중", "이미 SMPP 조회가 진행 중입니다.")
            return

        if not self.biz_numbers:
            QMessageBox.warning(self, "사업자번호 없음", "먼저 '사업자번호 가져오기' 버튼을 눌러주세요.")
            return

        user_id = self.smpp_id_entry.text().strip()
        password = self.smpp_password_entry.text().strip()
        if not (user_id and password):
            QMessageBox.warning(self, "SMPP 자격증명", "SMPP ID와 PW를 모두 입력해주세요.")
            return

        biz_nos = [biz.replace("-", "") for biz in self.biz_numbers]
        for entry in self.smpp_result_map.values():
            entry.update({"women": "", "small": "", "error": ""})
        self.refresh_smpp_results_table()
        self.smpp_progress_bar.setMaximum(len(biz_nos))
        self.smpp_progress_bar.setValue(0)
        self.smpp_status_label.setText("조회 준비 중...")
        self._set_controls_enabled(False)

        self.smpp_worker = SmppLookupWorker(user_id, password, biz_nos)
        self.smpp_worker.progress.connect(self.update_smpp_progress)
        self.smpp_worker.finished.connect(
            lambda results, error: self.on_smpp_finished(results, error)
        )
        self.smpp_worker.start()

    def _set_controls_enabled(self, enabled: bool):
        self.smpp_lookup_button.setEnabled(enabled)
        self.db_combo.setEnabled(enabled)
        self.browse_button.setEnabled(enabled)
        self.load_biz_button.setEnabled(enabled)
        self.smpp_id_entry.setEnabled(enabled)
        self.smpp_password_entry.setEnabled(enabled)
        self.close_button.setEnabled(enabled)

    def update_smpp_progress(self, idx: int, total: int, biz_no: str):
        display_no = smpp_excel.normalize_biz_no(biz_no) or biz_no
        self.smpp_progress_bar.setMaximum(total)
        self.smpp_progress_bar.setValue(idx)
        self.smpp_status_label.setText(f"{idx}/{total} - {display_no}")

    def on_smpp_finished(self, results, error):
        self._set_controls_enabled(True)
        self.smpp_worker = None

        if error:
            QMessageBox.critical(self, "SMPP 조회 오류", str(error))
            self.smpp_status_label.setText("오류 발생")
            return

        self.merge_smpp_results(results)
        self.refresh_smpp_results_table()
        success = sum(1 for r in results if not r.error)
        self.smpp_status_label.setText(f"{success}/{len(results)}개 완료")

    def merge_smpp_results(self, results):
        for result in results:
            display_no = smpp_excel.normalize_biz_no(result.biz_no) or result.biz_no
            entry = self.smpp_result_map.setdefault(
                display_no, {"women": "", "small": "", "error": ""}
            )

            if result.error:
                entry["error"] = result.error
                continue

            entry["error"] = ""
            entry["women"] = self.format_smpp_validity(result, "women")
            entry["small"] = self.format_smpp_validity(result, "small")

    def refresh_smpp_results_table(self):
        table = self.smpp_results_table
        table.setRowCount(len(self.company_entries))
        for row, entry in enumerate(self.company_entries):
            result = self.smpp_result_map.get(entry.biz_no, {})
            error_text = result.get("error") or ""

            company_item = QTableWidgetItem(entry.company_name or "")
            ceo_item = QTableWidgetItem(entry.ceo_name or "")
            biz_item = QTableWidgetItem(entry.biz_no)
            women_text = result.get("women") or ""
            small_text = result.get("small") or ""

            if error_text:
                women_text = women_text or "오류"
                small_text = small_text or "오류"

            women_item = QTableWidgetItem(women_text)
            small_item = QTableWidgetItem(small_text)

            if error_text:
                tooltip = f"SMPP 오류: {error_text}"
                for item in (company_item, ceo_item, biz_item, women_item, small_item):
                    item.setToolTip(tooltip)

            table.setItem(row, 0, company_item)
            table.setItem(row, 1, ceo_item)
            table.setItem(row, 2, biz_item)
            table.setItem(row, 3, women_item)
            table.setItem(row, 4, small_item)
        self.apply_company_filter(self.search_entry.text())

    def format_smpp_validity(self, result: CorpCheckResult, column_key: str) -> str:
        features = result.features
        if not features:
            return ""

        if column_key == "women":
            exists = features.women_exists
            confirm = features.women_confirm_date
            expire = features.women_expire_date
        else:
            exists = features.small_exists
            confirm = features.small_confirm_date
            expire = features.small_expire_date

        if not exists:
            return "해당사항 없음"
        if confirm and expire:
            return f"{confirm}~{expire}"
        return confirm or expire or "정보 없음"

    def apply_company_filter(self, keyword: str):
        keyword = (keyword or "").strip().lower()
        table = self.smpp_results_table
        for row in range(table.rowCount()):
            if not keyword:
                table.setRowHidden(row, False)
                continue
            company_item = table.item(row, 0)
            company_name = company_item.text().lower() if company_item else ""
            match = keyword in company_name
            table.setRowHidden(row, not match)

    def reject(self):
        if self.smpp_worker and self.smpp_worker.isRunning():
            QMessageBox.warning(self, "진행 중", "조회가 완료될 때까지 기다려주세요.")
            return
        super().reject()
