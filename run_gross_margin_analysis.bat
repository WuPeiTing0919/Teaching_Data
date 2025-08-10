@echo off
echo === 毛利分析器執行腳本 ===
echo.

cd /d "%~dp0main"
echo 切換到 main 目錄...

echo.
echo 執行毛利分析...
python gross_margin_analyzer.py

echo.
echo 按任意鍵結束...
pause > nul
