Write-Host "開始執行 student_case_cleaner.py..." -ForegroundColor Green
Write-Host ""

# 切換到腳本所在目錄
Set-Location $PSScriptRoot

# 執行 Python 程式
python main/student_case_cleaner.py

Write-Host ""
Write-Host "程式執行完成！" -ForegroundColor Green
Read-Host "按 Enter 鍵繼續"
