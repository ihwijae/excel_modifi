from PySide6.QtCore import QThread, Signal
from PySide6.QtGui import QImage
from PIL import Image
import numpy as np
import ocr_logic
import ocr_utils

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

class ColorUpdateWorker(QThread):
    finished = Signal(str)
    def __init__(self, excel_path):
        super().__init__(); self.excel_path = excel_path
    def run(self):
        result_message = ocr_logic.batch_update_colors(self.excel_path)
        self.finished.emit(result_message)