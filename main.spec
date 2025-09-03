# -*- mode: python ; coding: utf-8 -*-

# =============================================================================
# PyInstaller 사양 파일 (main.spec) - 최종본
# =============================================================================

# easyocr 라이브러리의 경로를 자동으로 찾습니다.
import easyocr
from pathlib import Path
easyocr_path = Path(easyocr.__file__).parent

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        # easyocr가 사용하는 모델 파일들을 포함
        (str(easyocr_path / 'model'), 'easyocr/model'),
        # 사용자가 수정할 설정 파일을 포함
        ('ocr_config.json', '.')
    ],
    hiddenimports=[
        # PySide6(UI) 관련 숨겨진 import
        'PySide6.QtSvg',
        'PySide6.QtOpenGL',
        # PyMuPDF(PDF) 관련 숨겨진 import (런타임 오류 예방)
        'fitz_new',
        # easyocr 및 그 의존성(scipy) 관련 숨겨진 import
        'scipy._cyutility',
        'pkg_resources.py2_warn'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='협력업체 관리 프로그램',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon=None, # 아이콘 파일이 있다면 'my_icon.ico' 와 같이 경로를 지정
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='main',
)