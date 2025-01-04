@echo off
REM 배치 파일은 python apis.py 스크립트를 여러 포트에서 실행합니다.

powershell -File ../utils/pid_kill.ps1

start cmd /k "python ./lb.py --port 8600"

start cmd /k "python ./apis.py --port 8501 --excel ../data/bnk.xlsm"

start cmd /k "python ./apis.py --port 8502 --excel ../data/bnk.xlsm"

start cmd /k "python ./apis.py --port 8511 --excel ../data_detail/bnk.xlsm"

start cmd /k "python ./apis.py --port 8512 --excel ../data_detail/bnk.xlsm"

start cmd /k "python ./apis_pdf.py --port 8507"

echo All scripts have been started.
    