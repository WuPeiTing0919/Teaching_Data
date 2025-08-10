import pandas as pd
import re
import numpy as np

def clean_products_data(input_file, output_file):
    """
    清洗產品資料，處理各種資料品質問題
    
    清洗邏輯：
    - sku: 一律大寫、去空白、去除重複
    - name: 去空白 + Title Case
    - category: 標準化（accessories/cables...）
    - cost/price: 轉數值（去空白/NT$/中文），非數字變 NaN
    - active: 欄位轉 True/False（"TRUE","True","yes","Y"→True；空字串→False）
    """
    # 讀取資料
    df = pd.read_excel(input_file)
    print(f"原始資料筆數: {len(df)}")
    print(f"原始欄位: {df.columns.tolist()}")
    
    # 1. 清洗 SKU：大寫、去空白、標準化格式、去除重複
    def standardize_sku(sku):
        if pd.isna(sku):
            return sku
        
        sku = str(sku).strip().upper()
        
        # 標準化 SKU 格式：移除連字號，統一為 P001 格式
        # 處理 P-001 格式，移除連字號變成 P001
        sku = re.sub(r'P-(\d+)', r'P\1', sku)
        
        return sku
    
    df['sku'] = df['sku'].apply(standardize_sku)
    
    # 去除重複的 SKU，保留第一個出現的
    df = df.drop_duplicates(subset=['sku'], keep='first')
    print(f"去除重複 SKU 後筆數: {len(df)}")
    
    # 2. 清洗品名：去空白 + Title Case
    df['name'] = df['name'].str.strip().str.title()
    
    # 3. 標準化類別
    def standardize_category(category):
        if pd.isna(category):
            return 'unknown'
        
        category = str(category).strip().lower()
        
        # 標準化類別名稱
        category_mapping = {
            'accessories': 'accessories',
            'cables': 'cables',
            'cable': 'cables',
            'unknown': 'unknown',
            'nan': 'unknown'
        }
        
        # 模糊匹配
        for key, value in category_mapping.items():
            if key in category:
                return value
        
        return 'other'
    
    df['category'] = df['category'].apply(standardize_category)
    
    # 4. 清洗成本和價格：轉數值，非數字變 NaN
    def clean_numeric_value(value):
        if pd.isna(value):
            return np.nan
        
        value = str(value).strip()
        
        # 處理特殊符號
        if value in ['—', '-', 'nan', 'NaN']:
            return np.nan
        
        # 移除貨幣符號和空白
        value = re.sub(r'NT\$', '', value)  # 移除 NT$
        value = re.sub(r'[\s,，]', '', value)  # 移除空白和逗號
        
        # 處理中文數字
        chinese_to_number = {
            'one': '1', 'two': '2', 'three': '3', 'four': '4', 'five': '5',
            'six': '6', 'seven': '7', 'eight': '8', 'nine': '9', 'ten': '10',
            'hundred': '100', 'thousand': '1000'
        }
        
        for chinese, number in chinese_to_number.items():
            value = value.replace(chinese, number)
        
        # 嘗試轉換為數值
        try:
            return float(value)
        except ValueError:
            return np.nan
    
    df['cost'] = df['cost'].apply(clean_numeric_value)
    df['price'] = df['price'].apply(clean_numeric_value)
    
    # 5. 標準化 active 欄位
    def standardize_boolean(value):
        if pd.isna(value):
            return False
        
        value = str(value).strip().lower()
        
        # 轉換為 True 的值
        true_values = ['true', 'yes', 'y', '1', 'active']
        if value in true_values:
            return True
        
        # 轉換為 False 的值
        false_values = ['false', 'no', 'n', '0', 'inactive', 'nan']
        if value in false_values:
            return False
        
        # 空字串或無法識別的值設為 False
        return False
    
    df['active'] = df['active'].apply(standardize_boolean)
    
    # 6. 資料品質檢查
    print(f"\n清洗後資料統計:")
    print(f"總筆數: {len(df)}")
    print(f"有效成本筆數: {df['cost'].notna().sum()}")
    print(f"有效價格筆數: {df['price'].notna().sum()}")
    print(f"啟用產品筆數: {df['active'].sum()}")
    
    print(f"\n類別分布:")
    print(df['category'].value_counts())
    
    print(f"\n清洗後資料:")
    print(df.to_string(index=False))
    
    # 7. 儲存清洗後的資料
    df.to_excel(output_file, index=False)
    print(f"\n清洗後的資料已儲存至: {output_file}")
    
    return df

def main():
    """主程式"""
    input_file = "dirty/hw.products_dirty.xlsx"
    output_file = "clean/products_clean.xlsx"
    
    try:
        cleaned_df = clean_products_data(input_file, output_file)
        print(f"\n資料清洗完成！")
        print(f"原始資料筆數: 5")
        print(f"清洗後資料筆數: {len(cleaned_df)}")
        
    except Exception as e:
        print(f"錯誤: {e}")

if __name__ == "__main__":
    main()
