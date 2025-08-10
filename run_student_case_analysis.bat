@echo off
echo 啟動 Student Case Orders 分析儀表板...
echo.
echo 請確保已安裝所需的 Python 套件
echo 如果尚未安裝，請先執行: pip install -r requirements.txt
echo.
echo 正在啟動儀表板...
echo 啟動完成後，請在瀏覽器中開啟: http://127.0.0.1:8050
echo.
echo 按 Ctrl+C 可以停止儀表板
echo.
cd /d "%~dp0"
python main/student_case_analysis_dashboard.py
pause
