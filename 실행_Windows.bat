@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  py -3 -m venv .venv
  if errorlevel 1 goto error
)
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto error
if not exist "INPUT-부가세_확장.xlsx" (
  if not exist "INPUT-부가세.xlsx" (
    echo 원래 INPUT-부가세.xlsx 파일을 이 폴더에 복사한 뒤 다시 실행하세요.
    pause
    exit /b 2
  )
  .venv\Scripts\python.exe prepare_input.py --source "INPUT-부가세.xlsx" --target "INPUT-부가세_확장.xlsx" --year 2026
  if errorlevel 1 goto error
  echo 2026년 확장 입력파일을 만들었습니다. 엑셀 내용을 확인하고 저장한 뒤 다시 실행하세요.
  pause
  exit /b 0
)
.venv\Scripts\python.exe hometax_vat_downloader.py --input "INPUT-부가세_확장.xlsx" --resume
if errorlevel 1 goto error
pause
exit /b 0
:error
 echo 실행 오류 또는 일부 작업 실패. 위 메시지를 확인하세요.
 pause
 exit /b 1
