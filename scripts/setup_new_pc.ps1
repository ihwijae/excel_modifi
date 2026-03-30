param(
    [string]$PythonExe = "python"
)

$ErrorActionPreference = "Stop"

Write-Host "[1/4] Python 버전 확인"
& $PythonExe --version

Write-Host "[2/4] 가상환경 생성 (.venv)"
& $PythonExe -m venv .venv

Write-Host "[3/4] pip 업그레이드"
& .\.venv\Scripts\python.exe -m pip install --upgrade pip

Write-Host "[4/4] 라이브러리 설치 (requirements.txt)"
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt

Write-Host ""
Write-Host "설치 완료"
Write-Host "- 실행: .\\scripts\\run_app.ps1"
Write-Host "- EXE 빌드: .\\scripts\\build_exe.ps1"
Write-Host "- 참고: pdf2image 사용 시 Poppler 설치 필요"
