@echo off
REM 배치 파일은 python apis.py 스크립트를 여러 포트에서 실행합니다.

echo Starting apis.py on port 8501
start cmd /k "python ./apis.py --port 8501"

echo Starting apis.py on port 8502
start cmd /k "python ./apis.py --port 8502"

echo Starting apis.py on port 8503
start cmd /k "python ./apis.py --port 8503"

echo All scripts have been started.
