# 資料清洗與分析程式集

這個專案包含多個 Python 程式，用於清洗和分析各種資料檔案。

## 主要程式

### 1. 客戶資料清洗程式 (`customer_data_cleaner.py`)

專門用於清洗客戶資料，解決以下資料品質問題：

### 1. 空白和大小寫問題
- 去除所有欄位的前後空白
- 姓名轉換為 Title Case (首字母大寫)

### 2. Email 驗證
- 使用正則表達式驗證 Email 格式
- 無效的 Email 會被清除 (設為 NaN)
- 支援標準的 Email 格式：`user@domain.com`

### 3. 電話格式標準化
- 手機號碼：統一為 `09xxxxxxxx` 格式
- 市話：統一為 `(02)xxxx-xxxx` 格式
- 自動處理國際碼 `8869` → `09`
- 不符合格式的電話設為 NaN

### 4. 日期標準化
- 支援多種日期格式：
  - `2025/8/9` → `2025-08-09`
  - `09-08-2025` → `2025-08-09`
  - `8月1日2025年` → `2025-08-01`
  - `2025-08-05` → `2025-08-05`
- 統一輸出為 `YYYY-MM-DD` 格式

### 5. 城市名稱映射
- 建立城市名稱字典映射
- 支援多種寫法：
  - `Taipei` / `taipei` / `TPE` / `台北` → `Taipei`
  - `Taoyuan` / `taoyuan` / `桃園` → `Taoyuan`

### 6. 金額清理
- 移除 `NT$` 符號和逗號
- 破折號 `—` 和 `-` 轉換為 NaN
- 轉換為數值型別

### 7. 重複記錄處理
- 以 `Customer ID` 為主要去重鍵
- 優先保留有 Email 的記錄
- 統一 Customer ID 為大寫

### 8. 資料品質標記
- 新增 `Valid` 欄位
- 標記資料是否有效 (Email 和電話都有效)

## 使用方法

### 執行環境
- Python 3.12
- 路徑：`C:\Users\User\AppData\Local\Programs\Python\Python312\python.exe`

### 執行命令
```bash
python main/customer_data_cleaner.py
```

### 輸入輸出
- 輸入：`dirty/1.customers_dirty.xlsx`
- 輸出：`clean/customers_clean.xlsx`

## 清洗結果

### 原始資料問題
- 總筆數：5 筆
- 無效 Email：2 筆 (`bob@@example.com`, `charlie@example,com`)
- 重複 Customer ID：C004 有 2 筆記錄
- 日期格式不一致
- 城市名稱不統一
- 電話格式多樣化
- 金額帶符號和破折號

### 清洗後結果
- 總筆數：4 筆 (去除重複)
- 有效 Email：3 筆
- 有效電話：4 筆
- 有效日期：3 筆
- 有效金額：3 筆
- 有效記錄：3 筆

### 欄位說明
| 欄位 | 說明 | 範例 |
|------|------|------|
| Customer ID | 客戶編號 | C001, C002, C003, C004 |
| Name | 客戶姓名 (Title Case) | Alice, Bob, Charlie, Dávid |
| Email | 電子郵件 | ALICE@example.com, david@example.com |
| Phone | 電話號碼 | 0912345678, (02)2345-6789 |
| Join Date | 加入日期 (YYYY-MM-DD) | 2025-08-09, 2025-08-01 |
| City | 城市名稱 | Taipei, Taoyuan |
| Spend (NT$) | 消費金額 (數值) | 1200.0, 800.0, 2500.0 |
| Valid | 資料有效性 | True/False |

## 技術特點

- 使用 pandas 進行資料處理
- 正則表達式進行格式驗證
- 字典映射處理城市名稱
- 多種日期格式支援
- 完整的錯誤處理和日誌輸出
- 資料品質統計和摘要

## 注意事項

1. 程式會自動處理各種資料格式問題
2. 無效資料會被設為 NaN 而不是刪除
3. 重複記錄會保留第一筆有效記錄
4. 所有清洗過程都有詳細的日誌輸出
5. 清洗後的資料會自動儲存到 clean 資料夾

---

## 2. 銷售資料分析儀表板

### 功能特色

✅ **當日銷售總額趨勢圖** - 顯示各日期的銷售總額變化  
✅ **前 5 名商品銷售排行** - 橫向柱狀圖顯示最熱銷商品  
✅ **按區域彙總分析** - 圓餅圖和柱狀圖顯示各區域銷售情況  
✅ **摘要統計資訊** - 總銷售額、訂單數、商品種類、銷售區域等關鍵指標  
✅ **資料明細表格** - 顯示前 50 筆銷售記錄  

### 使用方法

#### 方法一：使用批次檔案 (Windows)
```bash
# 雙擊執行
run_sales_analysis.bat
```

#### 方法二：使用 PowerShell 腳本
```bash
# 在 PowerShell 中執行
.\run_sales_analysis.ps1
```

#### 方法三：直接執行 Python 腳本
```bash
cd main
python sales_analysis_simple.py
```

### 輸出結果

- 腳本會生成 `sales_dashboard.html` 檔案
- 在瀏覽器中開啟此檔案即可查看完整的分析儀表板
- 包含互動式圖表和詳細的統計資訊

### 技術架構

- **資料處理**: pandas, openpyxl
- **圖表繪製**: Chart.js (CDN)
- **網頁框架**: 純 HTML/CSS/JavaScript
- **樣式設計**: 響應式設計，支援各種螢幕尺寸

### 資料來源

腳本會讀取 `clean/sales_clean.xlsx` 檔案中的 `Cleaned_Data` 工作表，該工作表包含以下欄位：

- **OrderID**: 訂單編號
- **Order Date**: 訂單日期
- **Product**: 產品名稱
- **Qty**: 數量
- **Unit Price**: 單價
- **line_amount**: 小計金額
- **Region**: 銷售區域
