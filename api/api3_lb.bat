@echo off
REM 배치 파일은 python apis.py 스크립트를 여러 포트에서 실행합니다.

start cmd /k "python ./lb.py --port 8600"

start cmd /k "python ./apis.py --port 8501 --excel ../data/bnk.xlsx"

start cmd /k "python ./apis.py --port 8502 --excel ../data/bnk.xlsx"

start cmd /k "python ./apis.py --port 8511 --excel ../data_detail/bnk.xlsx"

start cmd /k "python ./apis.py --port 8512 --excel ../data_detail/bnk.xlsx"

echo All scripts have been started.
    