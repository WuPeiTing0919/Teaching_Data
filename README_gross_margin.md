# 毛利分析器使用說明

## 功能概述

這個毛利分析器可以幫助您：
1. **計算產品毛利率**：使用公式 `gross_margin = (price - cost) / price`
2. **自動回推建議售價**：使用公式 `suggested_price = cost / (1 - target_margin)`
3. **生成毛利報表**：包含所有產品的詳細分析
4. **低毛利警示Email**：自動生成適合Gmail的警示郵件

## 檔案結構

```
main/
├── gross_margin_analyzer.py    # 主要分析程式
├── gross_margin_report.xlsx    # 生成的毛利報表
├── gross_margin_email_draft.txt        # 純文本Email草稿
└── gross_margin_gmail_email.html      # Gmail HTML Email
```

## 使用方法

### 方法1：使用批次檔案 (Windows)
```bash
# 雙擊執行
run_gross_margin_analysis.bat
```

### 方法2：使用PowerShell腳本
```powershell
# 在PowerShell中執行
.\run_gross_margin_analysis.ps1
```

### 方法3：直接執行Python
```bash
cd main
python gross_margin_analyzer.py
```

## 輸出結果

執行完成後，您會得到：

### 1. 毛利報表 (gross_margin_report.xlsx)
- 所有產品的原始數據
- 計算出的毛利率
- 建議售價
- 低毛利標記

### 2. Email草稿 (gross_margin_email_draft.txt)
- 純文本格式
- 可直接複製到任何郵件系統

### 3. Gmail Email (gross_margin_gmail_email.html)
- HTML格式，適合Gmail
- 美觀的表格和樣式
- 可直接複製到Gmail

## 公式說明

### 毛利率計算
```
gross_margin = (price - cost) / price
毛利率 = (售價 - 成本) / 售價
```

### 建議售價計算
```
suggested_price = cost / (1 - target_margin)
建議售價 = 成本 / (1 - 目標毛利率)
```

**範例：**
- 成本：$100
- 目標毛利率：40%
- 建議售價：$100 / (1 - 0.4) = $166.67

## 設定參數

在 `gross_margin_analyzer.py` 中，您可以調整：

```python
# 目標毛利率 (預設 40%)
target_margin = 0.40

# 輸入文件路徑
input_file = "../clean/products_clean.xlsx"
```

## 數據要求

您的 `products_clean.xlsx` 需要包含以下欄位：
- `sku`: 產品編號
- `name`: 產品名稱
- `category`: 產品類別
- `cost`: 成本
- `price`: 售價
- `active`: 是否啟用

## 常見問題

### Q: 如何修改目標毛利率？
A: 編輯 `gross_margin_analyzer.py` 中的 `target_margin` 變數

### Q: 如何自訂Email內容？
A: 修改 `generate_low_margin_email()` 和 `generate_gmail_email()` 方法

### Q: 如何添加更多分析指標？
A: 在 `calculate_gross_margin()` 方法中添加新的計算邏輯

## 技術特點

- **自動化**：一鍵生成完整報表和Email
- **靈活性**：支援自訂目標毛利率
- **多格式**：同時生成Excel報表和Email草稿
- **Gmail優化**：HTML格式適合Gmail使用
- **錯誤處理**：完善的異常處理和提示訊息

## 聯絡支援

如有任何問題或建議，請隨時聯繫。
