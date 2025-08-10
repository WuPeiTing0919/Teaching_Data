@echo off
echo 正在生成靜態 HTML 報表...
echo.
echo 請確保已安裝所需的 Python 套件
echo 如果尚未安裝，請先執行: pip install -r requirements.txt
echo.
echo 正在生成報表...
echo.
cd /d "%~dp0"
python main/generate_static_report.py
echo.
echo 按任意鍵開啟報表資料夾...
pause >nul
explorer static_report
