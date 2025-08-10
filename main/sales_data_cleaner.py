import pandas as pd
import numpy as np
import re
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

def clean_sales_data(input_file, output_file):
    """
    清洗 sales 資料的主要函數
    """
    print("開始讀取資料...")
    
    # 讀取 Excel 檔案
    try:
        df = pd.read_excel(input_file)
        print(f"成功讀取資料，共 {len(df)} 筆記錄")
    except Exception as e:
        print(f"讀取檔案時發生錯誤: {e}")
        return
    
    print("\n原始資料前5筆:")
    print(df.head())
    print(f"\n資料欄位: {df.columns.tolist()}")
    print(f"資料型別:\n{df.dtypes}")
    
    # 1. 日期格式標準化
    print("\n1. 處理日期格式...")
    df['Order Date'] = df['Order Date'].apply(standardize_date)
    
    # 2. 產品名稱清理
    print("2. 清理產品名稱...")
    df['Product'] = df['Product'].apply(clean_product_name)
    
    # 3. Qty 轉數值
    print("3. 處理數量欄位...")
    df['Qty'] = df['Qty'].apply(clean_quantity)
    
    # 4. 單價清理
    print("4. 清理單價...")
    df['Unit Price'] = df['Unit Price'].apply(clean_unit_price)
    
    # 5. 地區標準化
    print("5. 標準化地區...")
    df['Region'] = df['Region'].apply(standardize_region)
    
    # 6. 新增 line_amount
    print("6. 計算 line_amount...")
    df['line_amount'] = df['Qty'] * df['Unit Price']
    
    # 7. 去除重複（以 OrderID + Product 為準）
    print("7. 去除重複記錄...")
    initial_count = len(df)
    df = df.drop_duplicates(subset=['OrderID', 'Product'], keep='first')
    final_count = len(df)
    print(f"去除重複後: {initial_count} -> {final_count} 筆記錄")
    
    # 顯示清洗後的結果
    print("\n清洗後的資料:")
    print(df)
    
    # 建立彙總報表
    print("\n建立彙總報表...")
    summary_sheets = create_summary_reports(df)
    
    # 儲存到 Excel
    print(f"\n儲存清洗後的資料到: {output_file}")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Cleaned_Data', index=False)
        
        # 儲存彙總報表
        for sheet_name, summary_df in summary_sheets.items():
            summary_df.to_excel(writer, sheet_name=sheet_name, index=True)
    
    print("資料清洗完成！")
    return df

def standardize_date(date_value):
    """
    將各種日期格式轉換為 YYYY-MM-DD
    """
    if pd.isna(date_value):
        return np.nan
    
    # 如果是字串，嘗試解析
    if isinstance(date_value, str):
        # 處理中文日期格式
        if '年' in str(date_value) and '月' in str(date_value) and '日' in str(date_value):
            try:
                # 2025年7月28日 -> 2025-07-28
                match = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})日', str(date_value))
                if match:
                    year, month, day = match.groups()
                    return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
            except:
                pass
        
        # 處理英文日期格式
        try:
            # Jul 30, 2025 -> 2025-07-30
            parsed_date = pd.to_datetime(date_value, format='%b %d, %Y')
            return parsed_date.strftime('%Y-%m-%d')
        except:
            pass
        
        # 處理其他格式
        try:
            # 31/07/2025 -> 2025-07-31
            parsed_date = pd.to_datetime(date_value, dayfirst=True)
            return parsed_date.strftime('%Y-%m-%d')
        except:
            pass
    
    # 如果是 datetime 物件
    try:
        return pd.to_datetime(date_value).strftime('%Y-%m-%d')
    except:
        return np.nan

def clean_product_name(product):
    """
    清理產品名稱：去除前後空白 + Title Case
    """
    if pd.isna(product):
        return np.nan
    
    # 去除前後空白
    cleaned = str(product).strip()
    
    # 轉為 Title Case
    cleaned = cleaned.title()
    
    return cleaned

def clean_quantity(qty):
    """
    清理數量：轉為數值，支援英文數字轉換，無法轉的用 NaN
    """
    if pd.isna(qty):
        return np.nan
    
    # 如果是數字，直接返回
    if isinstance(qty, (int, float)):
        return float(qty)
    
    # 如果是字串，嘗試轉換
    qty_str = str(qty).strip().lower()
    
    # 處理特殊情況
    if qty_str in ['', 'nan', 'none', 'null']:
        return np.nan
    
    # 英文數字對照表
    word_to_number = {
        'zero': 0, 'one': 1, 'two': 2, 'three': 3, 'four': 4, 'five': 5,
        'six': 6, 'seven': 7, 'eight': 8, 'nine': 9, 'ten': 10,
        'eleven': 11, 'twelve': 12, 'thirteen': 13, 'fourteen': 14, 'fifteen': 15,
        'sixteen': 16, 'seventeen': 17, 'eighteen': 18, 'nineteen': 19, 'twenty': 20
    }
    
    # 檢查是否為英文數字
    if qty_str in word_to_number:
        return float(word_to_number[qty_str])
    
    # 嘗試轉換為數字
    try:
        return float(qty_str)
    except ValueError:
        return np.nan

def clean_unit_price(price):
    """
    清理單價：移除逗號/NT$/空白，破折號 → NaN
    """
    if pd.isna(price):
        return np.nan
    
    # 如果是數字，直接返回
    if isinstance(price, (int, float)):
        return float(price)
    
    price_str = str(price).strip()
    
    # 處理破折號
    if price_str in ['—', '-', 'nan', 'none', 'null', '']:
        return np.nan
    
    # 移除 NT$ 符號
    price_str = re.sub(r'NT\$', '', price_str)
    
    # 移除逗號
    price_str = re.sub(r',', '', price_str)
    
    # 移除空白
    price_str = re.sub(r'\s+', '', price_str)
    
    # 嘗試轉換為數字
    try:
        return float(price_str)
    except ValueError:
        return np.nan

def standardize_region(region):
    """
    標準化地區名稱
    """
    if pd.isna(region):
        return np.nan
    
    region_str = str(region).strip()
    
    # 建立地區對照表
    region_mapping = {
        'north': 'North',
        'NORTH': 'North',
        'North': 'North',
        'east': 'East', 
        'EAST': 'East',
        'East': 'East',
        'south': 'South',
        'SOUTH': 'South',
        'South': 'South',
        'west': 'West',
        'WEST': 'West',
        'West': 'West',
        'taichung': 'Taichung',
        'TAICHUNG': 'Taichung',
        'Taichung': 'Taichung',
        'taipei': 'Taipei',
        'TAIPEI': 'Taipei',
        'Taipei': 'Taipei',
        'kaohsiung': 'Kaohsiung',
        'KAOHSIUNG': 'Kaohsiung',
        'Kaohsiung': 'Kaohsiung'
    }
    
    return region_mapping.get(region_str.lower(), region_str)

def create_summary_reports(df):
    """
    建立彙總報表，包含樞紐表格式
    """
    summary_sheets = {}
    
    # 1. 按地區彙總
    print("建立地區彙總報表...")
    region_summary = df.groupby('Region').agg({
        'OrderID': 'count',
        'Qty': 'sum',
        'line_amount': 'sum'
    }).rename(columns={
        'OrderID': 'Order_Count',
        'Qty': 'Total_Qty',
        'line_amount': 'Total_Amount'
    })
    summary_sheets['Region_Summary'] = region_summary
    
    # 2. 按產品彙總
    print("建立產品彙總報表...")
    product_summary = df.groupby('Product').agg({
        'OrderID': 'count',
        'Qty': 'sum',
        'line_amount': 'sum'
    }).rename(columns={
        'OrderID': 'Order_Count',
        'Qty': 'Total_Qty',
        'line_amount': 'Total_Amount'
    })
    summary_sheets['Product_Summary'] = product_summary
    
    # 3. 樞紐表：產品為標題，地區為分類，line_amount為數值
    print("建立樞紐表...")
    pivot_table = df.pivot_table(
        values='line_amount',
        index='Product',
        columns='Region',
        aggfunc='sum',
        fill_value=0
    )
    
    # 新增總計欄位
    pivot_table['Total'] = pivot_table.sum(axis=1)
    
    # 新增總計列
    total_row = pivot_table.sum()
    pivot_table.loc['Total'] = total_row
    
    summary_sheets['Pivot_Table'] = pivot_table
    
    # 4. 按日期彙總
    print("建立日期彙總報表...")
    date_summary = df.groupby('Order Date').agg({
        'OrderID': 'count',
        'Qty': 'sum',
        'line_amount': 'sum'
    }).rename(columns={
        'OrderID': 'Order_Count',
        'Qty': 'Total_Qty',
        'line_amount': 'Total_Amount'
    })
    summary_sheets['Date_Summary'] = date_summary
    
    return summary_sheets

if __name__ == "__main__":
    # 設定檔案路徑
    input_file = "dirty/3.sales_dirty.xlsx"
    output_file = "clean/sales_clean.xlsx"
    
    # 執行資料清洗
    cleaned_df = clean_sales_data(input_file, output_file)
    
    if cleaned_df is not None:
        print(f"\n清洗完成！輸出檔案: {output_file}")
        print(f"包含以下工作表:")
        print("- Cleaned_Data: 清洗後的資料")
        print("- Region_Summary: 地區彙總")
        print("- Product_Summary: 產品彙總") 
        print("- Pivot_Table: 樞紐表 (產品為標題，地區為分類)")
        print("- Date_Summary: 日期彙總")
