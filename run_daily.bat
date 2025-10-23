@echo off
setlocal enabledelayedexpansion

REM === 프로젝트 루트(이 파일이 있는 폴더)로 이동 ===
cd /d "%~dp0"

REM === 가상환경 생성 (없으면) ===
if not exist .venv (
    echo [INFO] Creating virtual environment...
    py -3 -m venv .venv
)

REM === 가상환경 활성화 ===
call .venv\Scripts\activate.bat

REM === 의존성 설치/업데이트 ===
python -m pip install --upgrade pip
pip install -r requirements.txt

REM === 로그 폴더 준비 ===
if not exist logs mkdir logs

REM === 수집 실행 (도시 변경 원하면 --city suwon 등 사용) ===
echo [RUN] %date% %time% >> logs\run.log
python fetch_weather.py --city seoul >> logs\run.log 2>&1

endlocal
