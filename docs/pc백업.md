# PC 백업/복구 인계 문서

이 문서는 새 PC에서 이 프로젝트(`엑셀업데이트`)를 동일하게 재구성하기 위한 기준 문서입니다.

## 1. 개발 환경 기준
- OS: Windows
- Python: 3.13.x (현재 개발 PC 기준: 3.13.7)
- Shell: PowerShell

## 2. 저장소 클론
```powershell
git clone <repo-url>
cd 엑셀업데이트
```

## 3. 자동 세팅 (권장)
아래 스크립트 1회 실행으로 가상환경 생성 + 라이브러리 설치를 진행합니다.

```powershell
.\scripts\setup_new_pc.ps1
```

## 4. 수동 세팅 (필요 시)
```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## 5. 실행 방법
```powershell
.\scripts\run_app.ps1
```

또는 직접 실행:
```powershell
.\.venv\Scripts\python.exe .\main.py
```

## 6. 빌드 방법 (PyInstaller)
```powershell
.\scripts\build_exe.ps1
```

산출물:
- `dist\main\`

## 7. 필수 라이브러리
`requirements.txt` 기준:
- PySide6
- easyocr
- pdf2image
- numpy
- Pillow
- opencv-python
- openpyxl
- PyMuPDF
- requests
- beautifulsoup4
- pyinstaller

## 8. 외부 도구 주의사항
`pdf2image`를 쓰는 기능은 Poppler가 필요할 수 있습니다.
- Poppler 미설치 시 PDF 이미지 변환 단계에서 오류가 발생할 수 있음
- 새 PC에서 관련 오류가 나면 Poppler 설치 후 PATH 설정 또는 코드 내 경로 설정 확인

## 9. 환경변수(.env) 처리
- 실제 키 파일: `.env` (Git 커밋 금지)
- 템플릿 파일: `.env.example`

초기화:
```powershell
Copy-Item .env.example .env
```

그 다음 `.env`에 실제 값을 넣어 사용합니다.

## 10. Git 제외 항목
`.gitignore`에 아래 항목이 반영되어 있음:
- `dist/`
- `build/`
- `__pycache__/`
- `*.pyc`
- `.venv/`
- `.env`
- `.idea/`
- `.vscode/`

## 11. 다음 세션 Codex 인계용 메모
다음 세션에서 아래 순서로 요청하면 빠르게 복구 가능:
1. `docs/pc백업.md` 기준으로 환경 재구성
2. `.\scripts\setup_new_pc.ps1` 실행
3. `.\scripts\run_app.ps1`로 실행 확인
4. 필요 시 `.\scripts\build_exe.ps1`로 배포 빌드
