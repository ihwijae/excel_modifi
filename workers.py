from PySide6.QtCore import QThread, Signal
from PySide6.QtGui import QImage
from PIL import Image

import ocr_logic
import ocr_utils
from smpp_service import check_corps


class RoiOcrWorker(QThread):
    progress = Signal(str, str)
    finished = Signal(str)

    def __init__(self, reader, image_qimage: QImage, fields_to_process):
        super().__init__()
        self.reader = reader
        self.image_qimage = image_qimage
        self.fields_to_process = fields_to_process

    def run(self):
        try:
            pil_image = Image.fromqimage(self.image_qimage)
            for field, data in self.fields_to_process.items():
                rect = data.get("roi")
                if not rect:
                    self.progress.emit(field, "[지정안됨]")
                    continue

                cropped_pil = pil_image.crop(
                    (rect.x(), rect.y(), rect.x() + rect.width(), rect.y() + rect.height())
                )
                preprocessed_img = ocr_utils.preprocess_image_for_ocr(cropped_pil)
                result = self.reader.readtext(preprocessed_img, detail=0, paragraph=True)
                text = " ".join(result) if result else ""
                self.progress.emit(field, text.strip())

            self.finished.emit("모든 영역 분석 완료!")
        except Exception as exc:  # pylint: disable=broad-except
            self.finished.emit(f"분석 중 오류 발생: {exc}")


class ColorUpdateWorker(QThread):
    finished = Signal(str)

    def __init__(self, excel_path: str):
        super().__init__()
        self.excel_path = excel_path

    def run(self):
        result_message = ocr_logic.batch_update_colors(self.excel_path)
        self.finished.emit(result_message)


class SmppLookupWorker(QThread):
    progress = Signal(int, int, str)
    finished = Signal(list, object)

    def __init__(self, user_id: str, password: str, biz_nos, delay_seconds: float = 0.0):
        super().__init__()
        self.user_id = user_id
        self.password = password
        self.biz_nos = list(biz_nos)
        self.delay_seconds = delay_seconds

    def run(self):
        try:
            results = check_corps(
                self.user_id,
                self.password,
                self.biz_nos,
                delay_seconds=self.delay_seconds,
                progress_callback=self._emit_progress,
            )
            self.finished.emit(results, None)
        except Exception as exc:  # pylint: disable=broad-except
            self.finished.emit([], exc)

    def _emit_progress(self, idx: int, total: int, biz_no: str):
        self.progress.emit(idx, total, biz_no)
