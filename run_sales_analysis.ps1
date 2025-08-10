# 銷售資料分析儀表板 PowerShell 腳本
Write-Host "銷售資料分析儀表板" -ForegroundColor Green
Write-Host "===================" -ForegroundColor Green
Write-Host ""

Write-Host "正在執行銷售資料分析..." -ForegroundColor Yellow
Write-Host ""

# 切換到 main 目錄
Set-Location "main"

# 執行 Python 腳本
python sales_analysis_simple.py

Write-Host ""
Write-Host "分析完成！HTML 儀表板已生成在專案根目錄" -ForegroundColor Green
Write-Host "檔案名稱: sales_dashboard.html" -ForegroundColor Cyan
Write-Host ""

# 詢問是否要開啟瀏覽器
$openBrowser = Read-Host "是否要在瀏覽器中開啟儀表板？(y/n)"
if ($openBrowser -eq "y" -or $openBrowser -eq "Y") {
    Start-Process "..\sales_dashboard.html"
    Write-Host "已在瀏覽器中開啟儀表板" -ForegroundColor Green
}

Write-Host ""
Write-Host "按任意鍵繼續..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
