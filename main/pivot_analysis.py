import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

def create_pivot_analysis(input_file, output_file):
    """
    創建樞紐分析表
    - index: region
    - columns: product
    - values: total_with_tax (line_amount)
    - aggfunc: sum
    """
    print("開始讀取清洗後的銷售資料...")
    
    try:
        # 讀取清洗後的資料
        df = pd.read_excel(input_file, sheet_name='Cleaned_Data')
        print(f"成功讀取資料，共 {len(df)} 筆記錄")
    except Exception as e:
        print(f"讀取檔案時發生錯誤: {e}")
        return
    
    print("\n資料欄位:")
    print(df.columns.tolist())
    print(f"\n資料前5筆:")
    print(df.head())
    
    # 檢查必要欄位是否存在
    required_columns = ['Region', 'Product', 'line_amount']
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        print(f"缺少必要欄位: {missing_columns}")
        return
    
    # 處理缺失值
    print("\n處理缺失值...")
    initial_count = len(df)
    df_clean = df.dropna(subset=['Region', 'Product', 'line_amount'])
    final_count = len(df_clean)
    print(f"去除缺失值後: {initial_count} -> {final_count} 筆記錄")
    
    # 創建樞紐表
    print("\n創建樞紐表...")
    print("設定:")
    print("- index: Region")
    print("- columns: Product")
    print("- values: line_amount")
    print("- aggfunc: sum")
    
    pivot_table = df_clean.pivot_table(
        values='line_amount',
        index='Region',
        columns='Product',
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
    region_ranking = df_clean.groupby('Region')['line_amount'].sum().sort_values(ascending=False)
    region_ranking = region_ranking.reset_index()
    region_ranking.columns = ['Region', 'Total_Amount']
    region_ranking['Percentage'] = (region_ranking['Total_Amount'] / region_ranking['Total_Amount'].sum() * 100).round(2)
    
    # 2. 產品排名 (按總金額)
    product_ranking = df_clean.groupby('Product')['line_amount'].sum().sort_values(ascending=False)
    product_ranking = product_ranking.reset_index()
    product_ranking.columns = ['Product', 'Total_Amount']
    product_ranking['Percentage'] = (product_ranking['Total_Amount'] / product_ranking['Total_Amount'].sum() * 100).round(2)
    
    # 3. 地區-產品組合分析
    region_product_analysis = df_clean.groupby(['Region', 'Product'])['line_amount'].sum().reset_index()
    region_product_analysis = region_product_analysis.sort_values(['Region', 'line_amount'], ascending=[True, False])
    
    # 儲存到 Excel
    print(f"\n儲存樞紐分析結果到: {output_file}")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 主要樞紐表
        pivot_table.to_excel(writer, sheet_name='Pivot_Table', index=True)
        
        # 地區排名
        region_ranking.to_excel(writer, sheet_name='Region_Ranking', index=False)
        
        # 產品排名
        product_ranking.to_excel(writer, sheet_name='Product_Ranking', index=False)
        
        # 地區-產品組合分析
        region_product_analysis.to_excel(writer, sheet_name='Region_Product_Analysis', index=False)
        
        # 原始資料 (用於參考)
        df_clean.to_excel(writer, sheet_name='Source_Data', index=False)
    
    print("樞紐分析完成！")
    print(f"\n輸出檔案包含以下工作表:")
    print("- Pivot_Table: 主要樞紐表 (地區為列，產品為欄)")
    print("- Region_Ranking: 地區排名 (按總金額)")
    print("- Product_Ranking: 產品排名 (按總金額)")
    print("- Region_Product_Analysis: 地區-產品組合分析")
    print("- Source_Data: 來源資料")
    
    return pivot_table

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
    # 設定檔案路徑
    input_file = "clean/sales_clean.xlsx"
    output_file = "clean/sales_pivot_analysis.xlsx"
    
    # 執行樞紐分析
    pivot_result = create_pivot_analysis(input_file, output_file)
    
    if pivot_result is not None:
        print(f"\n樞紐分析完成！輸出檔案: {output_file}")
        display_pivot_summary(pivot_result)
