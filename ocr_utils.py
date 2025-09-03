# [ocr_utils.py 파일 전체]
import re
import numpy as np
import cv2
from PIL import Image

# 🔹 숫자 데이터 정리 (시평액, 실적액 등)
def clean_ocr_number(text):
    # 공백, 쉼표, '천원', '원' 등 모든 불필요한 문자 제거
    text = re.sub(r'[^0-9]', '', text)
    return text if text else "0"

# 🔹 사업자등록번호 정리
def clean_biz_number(text):
    # 공백, 일반 하이픈, 특수 하이픈 등 모두 제거
    text = text.replace(' ', '').replace('—', '-').replace('–', '-')
    # OCR이 자주 실수하는 알파벳 'O'와 'I'를 숫자로 변경
    text = text.upper().replace('O', '0').replace('I', '1')
    # XXX-XX-XXXXX 형식의 패턴만 정확히 추출
    match = re.search(r'(\d{3}-?\d{2}-?\d{5})', text)
    return match.group(1) if match else ''

# 🔹 OCR 전 이미지 전처리 (흑백+이진화)
def preprocess_image_for_ocr(pil_img):
    """
    Pillow 이미지를 받아, OCR에 최적화된 OpenCV 이미지(Numpy 배열)로 변환합니다.
    """
    # 1. Pillow 이미지를 Numpy 배열로 변환
    img_np = np.array(pil_img)
    
    # 2. 컬러 이미지를 흑백으로 변환
    #   이미지가 이미 흑백이거나, 채널 정보가 없는 경우를 대비
    if len(img_np.shape) == 3 and img_np.shape[2] == 3:
        img_gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
    else:
        img_gray = img_np

    # 3. 이미지 노이즈를 줄이기 위해 약간의 블러 처리
    img_blurred = cv2.GaussianBlur(img_gray, (3, 3), 0)

    # 4. adaptiveThreshold를 사용하여, 조명이 균일하지 않은 문서에서도 글자와 배경을 명확하게 분리
    img_thresh = cv2.adaptiveThreshold(
        img_blurred, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        11, 2
    )
    return img_thresh