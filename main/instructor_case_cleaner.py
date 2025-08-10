import pandas as pd
import numpy as np
import re
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

def clean_orders_data(df, products_master_df):
    """
    清洗 orders_dirty 資料並增加計算欄位
    """
    print("開始清洗 orders_dirty 資料...")
    
    # 複製資料避免修改原始資料
    df_clean = df.copy()
    
    # 1. 清洗日期欄位 - 統一為 YYYY-MM-DD 格式
    def standardize_date(date_str):
        if pd.isna(date_str):
            return np.nan
        
        date_str = str(date_str).strip()
        
        # 處理各種日期格式
        try:
            # 已經是 YYYY-MM-DD 格式
            if re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
                return date_str
            
            # 處理 MM/DD/YYYY 格式
            if re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', date_str):
                return pd.to_datetime(date_str).strftime('%Y-%m-%d')
            
            # 處理 MM-DD-YYYY 格式
            if re.match(r'^\d{1,2}-\d{1,2}-\d{4}$', date_str):
                return pd.to_datetime(date_str).strftime('%Y-%m-%d')
            
            # 處理 DD/MM/YYYY 格式
            if re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', date_str):
                # 假設是 DD/MM/YYYY 格式
                parts = date_str.split('/')
                if len(parts) == 3:
                    return f"{parts[2]}-{parts[1].zfill(2)}-{parts[0].zfill(2)}"
            
            # 處理其他格式
            parsed_date = pd.to_datetime(date_str)
            return parsed_date.strftime('%Y-%m-%d')
            
        except:
            return np.nan
    
    # 清洗所有日期欄位
    date_columns = ['order_date', 'ship_date', 'due_date']
    for col in date_columns:
        df_clean[col] = df_clean[col].apply(standardize_date)
        print(f"  已清洗 {col} 欄位")
    
    # 2. 清洗 region - 首字大寫
    def clean_region(region):
        if pd.isna(region):
            return np.nan
        return str(region).strip().title()
    
    df_clean['region'] = df_clean['region'].apply(clean_region)
    print("  已清洗 region 欄位")
    
    # 3. 清洗 product - 去空白 + Title Case
    def clean_product(product):
        if pd.isna(product):
            return np.nan
        return str(product).strip().title()
    
    df_clean['product'] = df_clean['product'].apply(clean_product)
    print("  已清洗 product 欄位")
    
    # 4. 清洗 qty - 轉數值（three→3；空白→NaN）
    def clean_qty(qty):
        if pd.isna(qty):
            return np.nan
        
        qty_str = str(qty).strip().lower()
        
        # 數字轉換字典
        word_to_num = {
            'one': 1, 'two': 2, 'three': 3, 'four': 4, 'five': 5,
            'six': 6, 'seven': 7, 'eight': 8, 'nine': 9, 'ten': 10
        }
        
        # 如果是英文數字
        if qty_str in word_to_num:
            return word_to_num[qty_str]
        
        # 如果是數字字串，轉為數值
        try:
            return int(float(qty_str))
        except:
            return np.nan
    
    df_clean['qty'] = df_clean['qty'].apply(clean_qty)
    print("  已清洗 qty 欄位")
    
    # 5. 清洗 unit_price - 去逗號/貨幣/空白，「—」→NaN → 轉浮點
    def clean_unit_price(price):
        if pd.isna(price):
            return np.nan
        
        price_str = str(price).strip()
        
        # 處理破折號
        if price_str == '—' or price_str == '-':
            return np.nan
        
        # 移除貨幣符號和逗號
        price_str = re.sub(r'[NT$,\s]', '', price_str)
        
        # 轉為浮點數
        try:
            return float(price_str)
        except:
            return np.nan
    
    df_clean['unit_price'] = df_clean['unit_price'].apply(clean_unit_price)
    print("  已清洗 unit_price 欄位")
    
    # 6. 清洗 discount(%) - 去掉 %，無值→0
    def clean_discount(discount):
        if pd.isna(discount):
            return 0
        
        discount_str = str(discount).strip()
        
        # 移除 % 符號
        discount_str = discount_str.replace('%', '')
        
        # 轉為數值
        try:
            return float(discount_str)
        except:
            return 0
    
    df_clean['discount(%)'] = df_clean['discount(%)'].apply(clean_discount)
    print("  已清洗 discount(%) 欄位")
    
    print("orders_dirty 資料清洗完成！")
    
    # 增加計算欄位
    print("開始增加計算欄位...")
    
    # 7. 計算 subtotal = qty * unit_price * (1 - discount/100)
    df_clean['subtotal'] = df_clean.apply(
        lambda row: row['qty'] * row['unit_price'] * (1 - row['discount(%)'] / 100) 
        if pd.notna(row['qty']) and pd.notna(row['unit_price']) else np.nan, 
        axis=1
    )
    print("  已計算 subtotal 欄位")
    
    # 8. 計算 total_with_tax = subtotal * 1.05
    df_clean['total_with_tax'] = df_clean['subtotal'] * 1.05
    print("  已計算 total_with_tax 欄位")
    
    # 9. 計算 lead_time_days = ship_date - order_date（天）
    def calculate_lead_time(row):
        if pd.isna(row['ship_date']) or pd.isna(row['order_date']):
            return np.nan
        try:
            ship_date = pd.to_datetime(row['ship_date'])
            order_date = pd.to_datetime(row['order_date'])
            return (ship_date - order_date).days
        except:
            return np.nan
    
    df_clean['lead_time_days'] = df_clean.apply(calculate_lead_time, axis=1)
    print("  已計算 lead_time_days 欄位")
    
    # 10. 計算 overdue = ship_date > due_date（True/False）
    def calculate_overdue(row):
        if pd.isna(row['ship_date']) or pd.isna(row['due_date']):
            return np.nan
        try:
            ship_date = pd.to_datetime(row['ship_date'])
            due_date = pd.to_datetime(row['due_date'])
            return ship_date > due_date
        except:
            return np.nan
    
    df_clean['overdue'] = df_clean.apply(calculate_overdue, axis=1)
    print("  已計算 overdue 欄位")
    
    # 11. 參考 products_master 資料，並以 product 資料做比對，並補上對應 category
    # 建立 product 到 category 的對應字典
    product_category_map = dict(zip(products_master_df['product'], products_master_df['category']))
    
    # 新增 category 欄位
    df_clean['category'] = df_clean['product'].map(product_category_map)
    print("  已新增 category 欄位")
    
    print("計算欄位新增完成！")
    return df_clean

def clean_monthly_sales_wide(df):
    """
    清洗 monthly_sales_wide 資料
    """
    print("開始清洗 monthly_sales_wide 資料...")
    
    # 複製資料避免修改原始資料
    df_clean = df.copy()
    
    # 清洗數值欄位（Jan, Feb, Mar, Apr）
    numeric_columns = ['Jan', 'Feb', 'Mar', 'Apr']
    
    for col in numeric_columns:
        df_clean[col] = df_clean[col].apply(lambda x: 
            float(str(x).replace(',', '').strip()) if pd.notna(x) else np.nan
        )
        print(f"  已清洗 {col} 欄位")
    
    print("monthly_sales_wide 資料清洗完成！")
    return df_clean

def main():
    """
    主函數：讀取、清洗並儲存資料
    """
    print("開始處理 instructor_case_dirty.xlsx 檔案...")
    
    # 讀取 Excel 檔案
    input_file = 'dirty/2.instructor_case_dirty.xlsx'
    
    try:
        # 讀取各個工作表
        orders_df = pd.read_excel(input_file, sheet_name='orders_dirty')
        monthly_sales_df = pd.read_excel(input_file, sheet_name='monthly_sales_wide')
        products_df = pd.read_excel(input_file, sheet_name='products_master')
        
        print(f"成功讀取檔案，包含 {len(orders_df)} 筆訂單資料")
        
        # 清洗資料
        orders_clean = clean_orders_data(orders_df, products_df)
        monthly_sales_clean = clean_monthly_sales_wide(monthly_sales_df)
        
        # 建立輸出檔案名稱
        output_file = 'clean/instructor_case_clean.xlsx'
        
        # 建立樞紐分析表
        print("開始建立樞紐分析表...")
        
        # 建立樞紐分析：index=region, columns=product, values=total_with_tax, aggfunc=sum
        pivot_table = pd.pivot_table(
            orders_clean,
            index='region',
            columns='product',
            values='total_with_tax',
            aggfunc='sum',
            fill_value=0
        )
        
        # 計算每個 region 的總計
        pivot_table['Total'] = pivot_table.sum(axis=1)
        
        # 計算每個 product 的總計
        product_totals = pivot_table.sum()
        pivot_table.loc['Total'] = product_totals
        
        print("樞紐分析表建立完成！")
        print(f"樞紐分析表形狀: {pivot_table.shape}")
        
        # 儲存清洗後的資料和樞紐分析表
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            orders_clean.to_excel(writer, sheet_name='orders_clean', index=False)
            monthly_sales_clean.to_excel(writer, sheet_name='monthly_sales_wide_clean', index=False)
            products_df.to_excel(writer, sheet_name='products_master', index=False)
            pivot_table.to_excel(writer, sheet_name='pivot_region_product')
        
        print(f"\n資料清洗完成！已儲存至 {output_file}")
        print(f"orders_clean: {orders_clean.shape}")
        print(f"monthly_sales_wide_clean: {monthly_sales_clean.shape}")
        print(f"products_master: {products_df.shape}")
        
        # 顯示清洗前後的對比
        print("\n=== 清洗前後對比 ===")
        print("\norders_dirty 清洗前（前5筆）:")
        print(orders_df.head())
        print("\norders_clean 清洗後（前5筆）:")
        print(orders_clean.head())
        
        print("\nmonthly_sales_wide 清洗前:")
        print(monthly_sales_df)
        print("\nmonthly_sales_wide_clean 清洗後:")
        print(monthly_sales_clean)
        
    except Exception as e:
        print(f"處理過程中發生錯誤: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
