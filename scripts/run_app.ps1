$ErrorActionPreference = "Stop"

if (-not (Test-Path .\.venv\Scripts\python.exe)) {
    Write-Error ".venv가 없습니다. 먼저 .\\scripts\\setup_new_pc.ps1 를 실행하세요."
}

& .\.venv\Scripts\python.exe .\main.py
