@echo off
chcp 65001 >nul
cd /d "%~dp0"
python hometax_vat_downloader.py %*
set "VAT_EXIT_CODE=%ERRORLEVEL%"
if not "%VAT_EXIT_CODE%"=="0" echo 위 오류 원인과 오류진단 경로를 확인하세요.
pause
exit /b %VAT_EXIT_CODE%
