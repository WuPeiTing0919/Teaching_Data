# 毛利分析器執行腳本 (PowerShell)
Write-Host "=== 毛利分析器執行腳本 ===" -ForegroundColor Green
Write-Host ""

# 切換到 main 目錄
Set-Location ".\main"
Write-Host "切換到 main 目錄..." -ForegroundColor Yellow

Write-Host ""
Write-Host "執行毛利分析..." -ForegroundColor Cyan
python gross_margin_analyzer.py

Write-Host ""
Write-Host "按任意鍵結束..." -ForegroundColor Yellow
Read-Host
