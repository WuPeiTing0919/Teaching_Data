Write-Host "啟動 Student Case Orders 分析儀表板..." -ForegroundColor Green
Write-Host ""
Write-Host "請確保已安裝所需的 Python 套件" -ForegroundColor Yellow
Write-Host "如果尚未安裝，請先執行: pip install -r requirements.txt" -ForegroundColor Yellow
Write-Host ""
Write-Host "正在啟動儀表板..." -ForegroundColor Cyan
Write-Host "啟動完成後，請在瀏覽器中開啟: http://127.0.0.1:8050" -ForegroundColor Cyan
Write-Host ""
Write-Host "按 Ctrl+C 可以停止儀表板" -ForegroundColor Red
Write-Host ""

# 切換到腳本所在目錄
Set-Location $PSScriptRoot

# 執行 Python 腳本
try {
    python main/student_case_analysis_dashboard.py
}
catch {
    Write-Host "執行時發生錯誤: $_" -ForegroundColor Red
    Write-Host "請檢查 Python 是否已安裝，以及所需套件是否已安裝" -ForegroundColor Red
}

Write-Host ""
Write-Host "按任意鍵繼續..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
