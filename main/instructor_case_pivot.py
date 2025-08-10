import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

def create_instructor_case_pivot(input_file, output_file):
    """
    使用 instructor_case_clean.xlsx 中的 orders_clean 資料創建樞紐分析表
    - index: region
    - columns: product
    - values: total_with_tax
    - aggfunc: sum
    """
    print("開始讀取 instructor_case_clean.xlsx 中的 orders_clean 資料...")
    
    try:
        # 讀取 orders_clean 工作表
        df = pd.read_excel(input_file, sheet_name='orders_clean')
        print(f"成功讀取資料，共 {len(df)} 筆記錄")
    except Exception as e:
        print(f"讀取檔案時發生錯誤: {e}")
        return
    
    print("\n資料欄位:")
    print(df.columns.tolist())
    print(f"\n資料前5筆:")
    print(df.head())
    
    # 檢查必要欄位是否存在
    required_columns = ['region', 'product', 'total_with_tax']
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        print(f"缺少必要欄位: {missing_columns}")
        return
    
    # 處理缺失值
    print("\n處理缺失值...")
    initial_count = len(df)
    df_clean = df.dropna(subset=['region', 'product', 'total_with_tax'])
    final_count = len(df_clean)
    print(f"去除缺失值後: {initial_count} -> {final_count} 筆記錄")
    
    # 創建樞紐表
    print("\n創建樞紐表...")
    print("設定:")
    print("- index: region")
    print("- columns: product")
    print("- values: total_with_tax")
    print("- aggfunc: sum")
    
    pivot_table = df_clean.pivot_table(
        values='total_with_tax',
        index='region',
        columns='product',
        aggfunc='sum',
        fill_value=0
    )
    
    # 新增總計欄位 (產品總計)
    pivot_table['Total_By_Product'] = pivot_table.sum(axis=1)
    
    # 新增總計列 (地區總計)
    total_row = pivot_table.sum()
    pivot_table.loc['Total_By_Region'] = total_row
    
    print("\n樞紐表完成！")
    print(f"地區數量: {len(pivot_table) - 1}")  # 減1是因為有總計列
    print(f"產品數量: {len(pivot_table.columns) - 1}")  # 減1是因為有總計欄
    
    # 創建額外的分析工作表
    print("\n創建額外分析工作表...")
    
    # 1. 地區排名 (按總金額)
    region_ranking = df_clean.groupby('region')['total_with_tax'].sum().sort_values(ascending=False)
    region_ranking = region_ranking.reset_index()
    region_ranking.columns = ['Region', 'Total_Amount']
    region_ranking['Percentage'] = (region_ranking['Total_Amount'] / region_ranking['Total_Amount'].sum() * 100).round(2)
    
    # 2. 產品排名 (按總金額)
    product_ranking = df_clean.groupby('product')['total_with_tax'].sum().sort_values(ascending=False)
    product_ranking = product_ranking.reset_index()
    product_ranking.columns = ['Product', 'Total_Amount']
    product_ranking['Percentage'] = (product_ranking['Total_Amount'] / product_ranking['Total_Amount'].sum() * 100).round(2)
    
    # 3. 地區-產品組合分析
    region_product_analysis = df_clean.groupby(['region', 'product'])['total_with_tax'].sum().reset_index()
    region_product_analysis = region_product_analysis.sort_values(['region', 'total_with_tax'], ascending=[True, False])
    
    # 4. 按類別分析
    if 'category' in df_clean.columns:
        category_analysis = df_clean.groupby('category')['total_with_tax'].sum().sort_values(ascending=False)
        category_analysis = category_analysis.reset_index()
        category_analysis.columns = ['Category', 'Total_Amount']
        category_analysis['Percentage'] = (category_analysis['Total_Amount'] / category_analysis['Total_Amount'].sum() * 100).round(2)
    else:
        category_analysis = None
    
    # 5. 按訂單日期分析
    if 'order_date' in df_clean.columns:
        date_analysis = df_clean.groupby('order_date')['total_with_tax'].sum().sort_index()
        date_analysis = date_analysis.reset_index()
        date_analysis.columns = ['Order_Date', 'Total_Amount']
    else:
        date_analysis = None
    
    return pivot_table, region_ranking, product_ranking, region_product_analysis, category_analysis, date_analysis

def transform_monthly_sales_wide(input_file):
    """
    轉置 monthly_sales_wide 資料，將寬表轉換為長表
    從: product | Jan | Feb | Mar | Apr
    到: product | month | revenue
    """
    print("\n開始處理 monthly_sales_wide 資料轉置...")
    
    try:
        # 讀取 monthly_sales_wide_clean 工作表
        df = pd.read_excel(input_file, sheet_name='monthly_sales_wide_clean')
        print(f"成功讀取 monthly_sales_wide_clean 資料，共 {len(df)} 筆記錄")
        
        print("\n原始資料結構:")
        print(df.head())
        print(f"\n資料欄位: {df.columns.tolist()}")
        
        # 檢查是否有月份欄位
        month_columns = [col for col in df.columns if col in ['Jan', 'Feb', 'Mar', 'Apr']]
        if not month_columns:
            print("警告: 未找到月份欄位 (Jan, Feb, Mar, Apr)")
            return None
        
        print(f"\n找到月份欄位: {month_columns}")
        
        # 使用 melt 函數進行轉置
        # 將月份欄位轉換為長表格式
        df_melted = df.melt(
            id_vars=['product'],  # 保持 product 欄位不變
            value_vars=month_columns,  # 要轉置的月份欄位
            var_name='month',  # 新的月份欄位名稱
            value_name='revenue'  # 新的營收欄位名稱
        )
        
        # 排序資料：按照 product, month, revenue 的順序
        # 首先按照 product 排序，然後按照月份順序，最後按照 revenue 降序
        month_order = ['Jan', 'Feb', 'Mar', 'Apr']  # 定義月份順序
        df_melted['month'] = pd.Categorical(df_melted['month'], categories=month_order, ordered=True)
        df_melted = df_melted.sort_values(['product', 'month', 'revenue'], ascending=[True, True, False])
        
        print("\n轉置後的資料結構:")
        print(df_melted.head(10))
        print(f"\n轉置後資料形狀: {df_melted.shape}")
        
        # 顯示轉置統計
        print(f"\n轉置統計:")
        print(f"- 產品數量: {df_melted['product'].nunique()}")
        print(f"- 月份數量: {df_melted['month'].nunique()}")
        print(f"- 總記錄數: {len(df_melted)}")
        
        # 按產品和月份顯示營收
        print(f"\n各產品各月份營收:")
        pivot_summary = df_melted.pivot_table(
            values='revenue',
            index='product',
            columns='month',
            aggfunc='sum',
            fill_value=0
        )
        print(pivot_summary)
        
        return df_melted
        
    except Exception as e:
        print(f"處理 monthly_sales_wide 資料時發生錯誤: {e}")
        return None

def main():
    """
    主函數：執行樞紐分析和資料轉置
    """
    # 設定檔案路徑
    input_file = "clean/instructor_case_clean.xlsx"
    output_file = "clean/instructor_case_pivot_analysis.xlsx"
    
    # 執行樞紐分析
    print("=== 執行樞紐分析 ===")
    pivot_result = create_instructor_case_pivot(input_file, output_file)
    
    if pivot_result is None:
        print("樞紐分析失敗，無法繼續執行")
        return
    
    pivot_table, region_ranking, product_ranking, region_product_analysis, category_analysis, date_analysis = pivot_result
    
    # 執行 monthly_sales_wide 轉置
    print("\n=== 執行 monthly_sales_wide 轉置 ===")
    monthly_sales_transformed = transform_monthly_sales_wide(input_file)
    
    # 儲存到 Excel
    print(f"\n儲存分析結果到: {output_file}")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 主要樞紐表
        pivot_table.to_excel(writer, sheet_name='Pivot_Table', index=True)
        
        # 地區排名
        region_ranking.to_excel(writer, sheet_name='Region_Ranking', index=False)
        
        # 產品排名
        product_ranking.to_excel(writer, sheet_name='Product_Ranking', index=False)
        
        # 地區-產品組合分析
        region_product_analysis.to_excel(writer, sheet_name='Region_Product_Analysis', index=False)
        
        # 類別分析 (如果有的話)
        if category_analysis is not None:
            category_analysis.to_excel(writer, sheet_name='Category_Analysis', index=False)
        
        # 日期分析 (如果有的話)
        if date_analysis is not None:
            date_analysis.to_excel(writer, sheet_name='Date_Analysis', index=False)
        
        # monthly_sales_wide 轉置後的資料
        if monthly_sales_transformed is not None:
            monthly_sales_transformed.to_excel(writer, sheet_name='Monthly_Sales_Transformed', index=False)
            
            # 額外創建轉置後的樞紐表
            monthly_pivot = monthly_sales_transformed.pivot_table(
                values='revenue',
                index='product',
                columns='month',
                aggfunc='sum',
                fill_value=0
            )
            monthly_pivot.to_excel(writer, sheet_name='Monthly_Sales_Pivot', index=True)
        
        # 來源資料 (用於參考)
        df_clean = pd.read_excel(input_file, sheet_name='orders_clean')
        df_clean = df_clean.dropna(subset=['region', 'product', 'total_with_tax'])
        df_clean.to_excel(writer, sheet_name='Source_Data', index=False)
    
    print("所有分析完成！")
    print(f"\n輸出檔案包含以下工作表:")
    print("- Pivot_Table: 主要樞紐表 (地區為列，產品為欄)")
    print("- Region_Ranking: 地區排名 (按總金額)")
    print("- Product_Ranking: 產品排名 (按總金額)")
    print("- Region_Product_Analysis: 地區-產品組合分析")
    if category_analysis is not None:
        print("- Category_Analysis: 類別分析")
    if date_analysis is not None:
        print("- Date_Analysis: 日期分析")
    if monthly_sales_transformed is not None:
        print("- Monthly_Sales_Transformed: 轉置後的月度銷售資料 (product, month, revenue)")
        print("- Monthly_Sales_Pivot: 月度銷售樞紐表")
    print("- Source_Data: 來源資料")
    
    # 顯示樞紐表摘要
    display_pivot_summary(pivot_table)

def display_pivot_summary(pivot_table):
    """
    顯示樞紐表摘要資訊
    """
    print("\n=== 樞紐表摘要 ===")
    print(f"總地區數: {len(pivot_table) - 1}")
    print(f"總產品數: {len(pivot_table.columns) - 1}")
    
    # 顯示各地區總計
    print("\n各地區總計:")
    region_totals = pivot_table['Total_By_Product'].sort_values(ascending=False)
    for region, total in region_totals.items():
        if region != 'Total_By_Region':
            print(f"  {region}: {total:,.2f}")
    
    # 顯示各產品總計
    print("\n各產品總計:")
    product_totals = pivot_table.loc['Total_By_Region'].sort_values(ascending=False)
    for product, total in product_totals.items():
        if product != 'Total_By_Product':
            print(f"  {product}: {total:,.2f}")
    
    # 總計
    grand_total = pivot_table.loc['Total_By_Region', 'Total_By_Product']
    print(f"\n總計: {grand_total:,.2f}")

if __name__ == "__main__":
    main()
