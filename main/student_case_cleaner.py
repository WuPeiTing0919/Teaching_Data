import pandas as pd
import numpy as np
import re
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

def clean_orders_data(df):
    """
    清洗 orders 資料
    
    清洗邏輯：
    - 日期欄（order_date）統一為 YYYY-MM-DD（無法解析 → NaN）
    - qty 轉數值：「seven」→ 7；不可解析 → NaN
    - discount 去掉 % 後轉數值；無效或空字串 → 0
    """
    print("開始清洗 orders 資料...")
    
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
    
    # 清洗 order_date 欄位
    df_clean['order_date'] = df_clean['order_date'].apply(standardize_date)
    print("  已清洗 order_date 欄位")
    
    # 2. 清洗 qty - 轉數值（seven→7；不可解析→NaN）
    def clean_qty(qty):
        if pd.isna(qty):
            return np.nan
        
        qty_str = str(qty).strip().lower()
        
        # 數字轉換字典
        word_to_num = {
            'one': 1, 'two': 2, 'three': 3, 'four': 4, 'five': 5,
            'six': 6, 'seven': 7, 'eight': 8, 'nine': 9, 'ten': 10,
            'eleven': 11, 'twelve': 12, 'thirteen': 13, 'fourteen': 14, 'fifteen': 15,
            'sixteen': 16, 'seventeen': 17, 'eighteen': 18, 'nineteen': 19, 'twenty': 20
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
    
    # 3. 清洗 discount - 去掉 % 後轉數值；無效或空字串 → 0
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
    
    df_clean['discount'] = df_clean['discount'].apply(clean_discount)
    print("  已清洗 discount 欄位")
    
    print("orders 資料清洗完成！")
    return df_clean

def clean_products_master_data(df):
    """
    清洗 products_master 資料
    
    清洗邏輯：
    - product_name 去空白、Title Case
    - category 規範大小寫（首字大寫）
    - unit_price 去掉「USD」「逗號/空白」後轉數值
    - tax_rate：若是「5%」→ 0.05；確保為小數
    """
    print("開始清洗 products_master 資料...")
    
    # 複製資料避免修改原始資料
    df_clean = df.copy()
    
    # 1. 清洗 product_name - 去空白、Title Case
    def clean_product_name(name):
        if pd.isna(name):
            return np.nan
        return str(name).strip().title()
    
    df_clean['product_name'] = df_clean['product_name'].apply(clean_product_name)
    print("  已清洗 product_name 欄位")
    
    # 2. 清洗 category - 首字大寫
    def clean_category(category):
        if pd.isna(category):
            return np.nan
        return str(category).strip().title()
    
    df_clean['category'] = df_clean['category'].apply(clean_category)
    print("  已清洗 category 欄位")
    
    # 3. 清洗 unit_price - 去掉「USD」「逗號/空白」後轉數值
    def clean_unit_price(price):
        if pd.isna(price):
            return np.nan
        
        price_str = str(price).strip()
        
        # 處理破折號或無效值
        if price_str in ['—', '-', 'nan', 'NaN', '']:
            return np.nan
        
        # 移除 USD 和逗號
        price_str = re.sub(r'USD\s*', '', price_str, flags=re.IGNORECASE)
        price_str = re.sub(r'[,\s]', '', price_str)
        
        # 轉為浮點數
        try:
            return float(price_str)
        except:
            return np.nan
    
    df_clean['unit_price'] = df_clean['unit_price'].apply(clean_unit_price)
    print("  已清洗 unit_price 欄位")
    
    # 4. 清洗 tax_rate - 若是「5%」→ 0.05；確保為小數
    def clean_tax_rate(tax_rate):
        if pd.isna(tax_rate):
            return np.nan
        
        tax_str = str(tax_rate).strip()
        
        # 處理百分比格式
        if '%' in tax_str:
            # 移除 % 符號並轉為小數
            try:
                tax_value = float(tax_str.replace('%', ''))
                return tax_value / 100
            except:
                return np.nan
        
        # 如果已經是數值，確保是小數格式
        try:
            tax_value = float(tax_str)
            # 如果大於1，可能是百分比格式，轉為小數
            if tax_value > 1:
                return tax_value / 100
            return tax_value
        except:
            return np.nan
    
    df_clean['tax_rate'] = df_clean['tax_rate'].apply(clean_tax_rate)
    print("  已清洗 tax_rate 欄位")
    
    print("products_master 資料清洗完成！")
    return df_clean


def clean_monthly_sales_wide_data(df):
    """
    清洗 monthly_sales_wide 資料
    
    清洗邏輯：
    - region 欄位去空白、Title Case
    - 月份欄位（Jan, Feb, Mar）確保為數值格式
    """
    print("開始清洗 monthly_sales_wide 資料...")
    
    # 複製資料避免修改原始資料
    df_clean = df.copy()
    
    # 1. 清洗 region 欄位 - 去空白、Title Case
    def clean_region(region):
        if pd.isna(region):
            return np.nan
        return str(region).strip().title()
    
    df_clean['region'] = df_clean['region'].apply(clean_region)
    print("  已清洗 region 欄位")
    
    # 2. 清洗月份欄位 - 確保為數值格式
    month_columns = ['Jan', 'Feb', 'Mar']
    for month in month_columns:
        if month in df_clean.columns:
            df_clean[month] = pd.to_numeric(df_clean[month], errors='coerce')
            print(f"  已清洗 {month} 欄位")
    
    print("monthly_sales_wide 資料清洗完成！")
    return df_clean


def main():
    """
    主函數：讀取、清洗並儲存資料
    """
    print("開始處理 hw2.student_case_dirty.xlsx 檔案...")
    
    # 讀取 Excel 檔案
    input_file = 'dirty/hw2.student_case_dirty.xlsx'
    
    try:
        # 讀取各個工作表
        orders_df = pd.read_excel(input_file, sheet_name='orders')
        products_master_df = pd.read_excel(input_file, sheet_name='products_master')
        monthly_sales_wide_df = pd.read_excel(input_file, sheet_name='monthly_sales_wide')
        
        print(f"成功讀取檔案")
        print(f"orders 工作表: {len(orders_df)} 筆資料")
        print(f"products_master 工作表: {len(products_master_df)} 筆資料")
        print(f"monthly_sales_wide 工作表: {len(monthly_sales_wide_df)} 筆資料")
        
        # 顯示原始資料結構
        print(f"\norders 欄位: {orders_df.columns.tolist()}")
        print(f"products_master 欄位: {products_master_df.columns.tolist()}")
        print(f"monthly_sales_wide 欄位: {monthly_sales_wide_df.columns.tolist()}")
        
        # 清洗資料
        orders_clean = clean_orders_data(orders_df)
        products_master_clean = clean_products_master_data(products_master_df)
        monthly_sales_wide_clean = clean_monthly_sales_wide_data(monthly_sales_wide_df)
        
        # 將 products_master_clean 的資料合併到 orders_clean
        print("\n開始合併 products_master 資料到 orders...")
        orders_clean = orders_clean.merge(
            products_master_clean[['product_id', 'product_name', 'category', 'unit_price', 'tax_rate']], 
            on='product_id', 
            how='left'
        )
        print("  已合併 products_master 資料")
        
        # 在 orders_clean 中增加計算欄位
        print("開始計算訂單金額...")
        
        # 計算 subtotal = qty * unit_price * (1 - discount/100)
        orders_clean['subtotal'] = orders_clean['qty'] * orders_clean['unit_price'] * (1 - orders_clean['discount'] / 100)
        print("  已計算 subtotal 欄位")
        
        # 計算 total_with_tax = subtotal * (1 + tax_rate)
        orders_clean['total_with_tax'] = orders_clean['subtotal'] * (1 + orders_clean['tax_rate'])
        print("  已計算 total_with_tax 欄位")
        
        print("訂單金額計算完成！")
        
        # 建立樞紐表分析
        print("\n開始建立樞紐表分析...")
        
        # 建立樞紐表：以 total_with_tax 做匯總，index=category，columns=product_name
        pivot_table = pd.pivot_table(
            orders_clean,
            values='total_with_tax',
            index='category',
            columns='product_name',
            aggfunc='sum',
            fill_value=0
        )
        
        # 計算每個 category 的總計
        pivot_table['Total'] = pivot_table.sum(axis=1)
        
        # 計算每個 product_name 的總計
        pivot_table.loc['Total'] = pivot_table.sum()
        
        print("  已建立樞紐表分析")
        print(f"  樞紐表大小: {pivot_table.shape}")
        
        # 建立資料轉置（unpivoting/melting）
        print("\n開始建立資料轉置...")
        
        # 將 monthly_sales_wide_clean 從寬格式轉為長格式
        # 欄位：region, month, revenue
        monthly_sales_long = monthly_sales_wide_clean.melt(
            id_vars=['region'],
            value_vars=['Jan', 'Feb', 'Mar'],
            var_name='month',
            value_name='revenue'
        )
        
        print("  已建立資料轉置")
        print(f"  轉置後資料大小: {monthly_sales_long.shape}")
        print(f"  轉置後欄位: {list(monthly_sales_long.columns)}")
        
        # 排序 monthly_sales_long 資料
        print("\n開始排序 monthly_sales_long 資料...")
        
        # 定義月份的正確順序（Jan → Feb → Mar）
        month_order = ['Jan', 'Feb', 'Mar']
        
        # 將 month 欄位轉換為 Categorical 類型，確保正確的排序順序
        monthly_sales_long['month'] = pd.Categorical(
            monthly_sales_long['month'], 
            categories=month_order, 
            ordered=True
        )
        
        # 按 region 和 month 排序
        monthly_sales_long = monthly_sales_long.sort_values(['region', 'month'])
        print("  已排序 monthly_sales_long 資料（月份按 Jan → Feb → Mar 順序）")
        
        # 建立輸出檔案名稱
        output_file = 'clean/student_case_clean.xlsx'
        
        # 儲存清洗後的資料
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            orders_clean.to_excel(writer, sheet_name='orders_clean', index=False)
            products_master_clean.to_excel(writer, sheet_name='products_master_clean', index=False)
            monthly_sales_wide_clean.to_excel(writer, sheet_name='monthly_sales_wide_clean', index=False)
            pivot_table.to_excel(writer, sheet_name='pivot_analysis')
            monthly_sales_long.to_excel(writer, sheet_name='monthly_sales_long', index=False)
        
        print(f"\n資料清洗完成！已儲存至 {output_file}")
        print(f"orders_clean: {orders_clean.shape}")
        print(f"products_master_clean: {products_master_clean.shape}")
        print(f"monthly_sales_wide_clean: {monthly_sales_wide_clean.shape}")
        print(f"monthly_sales_long: {monthly_sales_long.shape}")
        
        # 顯示清洗前後的對比
        print("\n=== 清洗前後對比 ===")
        print("\norders 清洗前（前5筆）:")
        print(orders_df.head())
        print("\norders 清洗後（前5筆）:")
        print(orders_clean.head())
        
        print("\nproducts_master 清洗前（前5筆）:")
        print(products_master_df.head())
        print("\nproducts_master 清洗後（前5筆）:")
        print(products_master_clean.head())
        
        print("\nmonthly_sales_wide 清洗前（前5筆）:")
        print(monthly_sales_wide_df.head())
        print("\nmonthly_sales_wide 清洗後（前5筆）:")
        print(monthly_sales_wide_clean.head())
        
        print("\nmonthly_sales_wide 轉置後（前5筆）:")
        print(monthly_sales_long.head())
        
        # 顯示資料品質統計
        print("\n=== 資料品質統計 ===")
        print(f"\norders 清洗前後比較:")
        print(f"  原始資料筆數: {len(orders_df)}")
        print(f"  清洗後資料筆數: {len(orders_clean)}")
        print(f"  日期欄位有效值: {orders_clean['order_date'].notna().sum()}/{len(orders_clean)}")
        print(f"  qty 欄位有效值: {orders_clean['qty'].notna().sum()}/{len(orders_clean)}")
        print(f"  discount 欄位有效值: {orders_clean['discount'].notna().sum()}/{len(orders_clean)}")
        
        print(f"\nproducts_master 清洗前後比較:")
        print(f"  原始資料筆數: {len(products_master_df)}")
        print(f"  清洗後資料筆數: {len(products_master_clean)}")
        print(f"  product_name 欄位有效值: {products_master_clean['product_name'].notna().sum()}/{len(products_master_clean)}")
        print(f"  category 欄位有效值: {products_master_clean['category'].notna().sum()}/{len(products_master_clean)}")
        print(f"  unit_price 欄位有效值: {products_master_clean['unit_price'].notna().sum()}/{len(products_master_clean)}")
        print(f"  tax_rate 欄位有效值: {products_master_clean['tax_rate'].notna().sum()}/{len(products_master_clean)}")
        
        print(f"\nmonthly_sales_wide 清洗前後比較:")
        print(f"  原始資料筆數: {len(monthly_sales_wide_df)}")
        print(f"  清洗後資料筆數: {len(monthly_sales_wide_clean)}")
        print(f"  region 欄位有效值: {monthly_sales_wide_clean['region'].notna().sum()}/{len(monthly_sales_wide_clean)}")
        print(f"  Jan 欄位有效值: {monthly_sales_wide_clean['Jan'].notna().sum()}/{len(monthly_sales_wide_clean)}")
        print(f"  Feb 欄位有效值: {monthly_sales_wide_clean['Feb'].notna().sum()}/{len(monthly_sales_wide_clean)}")
        print(f"  Mar 欄位有效值: {monthly_sales_wide_clean['Mar'].notna().sum()}/{len(monthly_sales_wide_clean)}")
        
    except Exception as e:
        print(f"處理過程中發生錯誤: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
