# Student Case Orders 分析儀表板

這是一個基於 Dash 和 Plotly 的互動式網頁儀表板，用於分析 `student_case_clean.xlsx` 檔案中的 `orders_clean` 資料。

## 功能特色

### 📊 統計摘要
- **總訂單數**: 顯示所有訂單的總數量
- **總營收**: 顯示所有訂單的總營收金額
- **產品種類**: 顯示不同產品的數量
- **產品類別**: 顯示產品類別的數量

### 📈 分析圖表
1. **營收趨勢圖**: 按月份顯示營收變化趨勢
2. **產品類別營收分布**: 圓餅圖顯示各類別的營收占比
3. **產品營收排行**: 水平長條圖顯示各產品的營收排行
4. **折扣分析**: 散點圖分析折扣與營收的關係

### 📋 資料表格
- 互動式資料表格，顯示所有訂單的詳細資訊
- 支援分頁瀏覽
- 包含訂單日期、產品名稱、類別、數量、單價、折扣、總金額等欄位

### 🔍 資料篩選
- **產品類別篩選**: 可選擇一個或多個產品類別
- **日期範圍篩選**: 可選擇特定的日期範圍
- **最小金額篩選**: 可設定最小金額門檻

## 安裝需求

確保已安裝以下 Python 套件：

```bash
pip install -r requirements.txt
```

主要套件包括：
- `pandas`: 資料處理
- `plotly`: 互動式圖表
- `dash`: 網頁應用框架
- `dash-bootstrap-components`: Bootstrap 樣式元件
- `dash-table`: 資料表格元件
- `openpyxl`: Excel 檔案讀取

## 使用方法

### 方法 1: 使用批次檔案 (Windows)
雙擊執行 `run_student_case_analysis.bat`

### 方法 2: 使用 PowerShell 腳本
在 PowerShell 中執行：
```powershell
.\run_student_case_analysis.ps1
```

### 方法 3: 直接執行 Python 腳本
```bash
python main/student_case_analysis_dashboard.py
```

## 啟動後的操作

1. 腳本執行成功後，會顯示類似以下的訊息：
   ```
   儀表板已創建完成！
   請在瀏覽器中開啟 http://127.0.0.1:8050 來查看儀表板
   ```

2. 在瀏覽器中開啟 `http://127.0.0.1:8050`

3. 儀表板會自動載入並顯示分析結果

4. 使用篩選器來調整顯示的資料範圍

5. 圖表支援互動操作（縮放、平移、懸停顯示詳細資訊等）

## 資料來源

儀表板會自動讀取 `clean/student_case_clean.xlsx` 檔案中的 `orders_clean` 工作表。

## 資料欄位說明

- **order_date**: 訂單日期
- **product_id**: 產品編號
- **product_name**: 產品名稱
- **category**: 產品類別
- **qty**: 訂購數量
- **unit_price**: 單價
- **discount**: 折扣百分比
- **subtotal**: 小計金額
- **total_with_tax**: 含稅總金額

## 停止儀表板

在終端機中按 `Ctrl+C` 即可停止儀表板服務。

## 故障排除

### 常見問題

1. **套件未安裝**
   - 執行 `pip install -r requirements.txt`

2. **埠號被占用**
   - 修改腳本中的埠號（預設為 8050）

3. **檔案路徑錯誤**
   - 確保 `clean/student_case_clean.xlsx` 檔案存在

4. **瀏覽器無法連線**
   - 檢查防火牆設定
   - 確認埠號是否正確

### 錯誤訊息

如果遇到錯誤，請檢查：
- Python 版本是否為 3.7 或以上
- 所需套件是否已正確安裝
- Excel 檔案是否存在且可讀取
- 檔案路徑是否正確

## 技術架構

- **後端**: Python + Dash
- **前端**: HTML + CSS + JavaScript (透過 Dash 自動生成)
- **圖表**: Plotly.js
- **樣式**: Bootstrap CSS 框架
- **資料處理**: Pandas

## 自訂與擴展

您可以修改 `student_case_analysis_dashboard.py` 來：
- 新增更多圖表類型
- 調整儀表板佈局
- 新增更多篩選條件
- 自訂圖表樣式和顏色
- 新增資料匯出功能

## 支援

如有問題或建議，請檢查：
1. 錯誤訊息和日誌
2. 套件版本相容性
3. 檔案路徑和權限設定
