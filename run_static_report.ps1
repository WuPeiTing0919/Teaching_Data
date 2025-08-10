Write-Host "正在生成靜態 HTML 報表..." -ForegroundColor Green
Write-Host ""
Write-Host "請確保已安裝所需的 Python 套件" -ForegroundColor Yellow
Write-Host "如果尚未安裝，請先執行: pip install -r requirements.txt" -ForegroundColor Yellow
Write-Host ""
Write-Host "正在生成報表..." -ForegroundColor Cyan
Write-Host ""

Set-Location $PSScriptRoot

try {
    python main/generate_static_report.py
    
    if (Test-Path "static_report") {
        Write-Host ""
        Write-Host "✅ 報表生成成功！" -ForegroundColor Green
        Write-Host "📁 報表位置: static_report/main_report.html" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "正在開啟報表資料夾..." -ForegroundColor Yellow
        Start-Process "static_report"
    } else {
        Write-Host "❌ 報表生成失敗，未找到 static_report 資料夾" -ForegroundColor Red
    }
}
catch {
    Write-Host "執行時發生錯誤: $_" -ForegroundColor Red
    Write-Host "請檢查 Python 是否已安裝，以及所需套件是否已安裝" -ForegroundColor Red
}

Write-Host ""
Write-Host "按任意鍵繼續..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
