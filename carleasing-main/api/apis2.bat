@echo off
REM 배치 파일은 python apis.py 스크립트를 여러 포트에서 실행합니다.

echo Starting apis.py on port 8504
start cmd /k "python ./apis.py --port 8504 --excel ../data/dgb.xlsm"

echo Starting apis.py on port 8505
start cmd /k "python ./apis.py --port 8505 --excel ../data/mz.xlsx"

echo All scripts have been started.
