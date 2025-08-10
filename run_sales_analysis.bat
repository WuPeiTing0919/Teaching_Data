@echo off
echo 銷售資料分析儀表板
echo ===================
echo.
echo 正在執行銷售資料分析...
echo.

cd main
python sales_analysis_simple.py

echo.
echo 分析完成！HTML 儀表板已生成在專案根目錄
echo 檔案名稱: sales_dashboard.html
echo.
echo 請在瀏覽器中開啟此檔案以查看分析結果
echo.
pause
