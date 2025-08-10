import pandas as pd
import numpy as np
from datetime import datetime
import os

class GrossMarginAnalyzer:
    def __init__(self, target_margin=0.40):
        """
        初始化毛利分析器
        
        Args:
            target_margin (float): 目標毛利率，預設為 40%
        """
        self.target_margin = target_margin
        self.df = None
        
    def load_products_data(self, file_path):
        """
        載入產品數據
        
        Args:
            file_path (str): 產品數據文件路徑
        """
        try:
            self.df = pd.read_excel(file_path)
            print(f"成功載入產品數據，共 {len(self.df)} 筆記錄")
            print(f"欄位: {list(self.df.columns)}")
            return True
        except Exception as e:
            print(f"載入產品數據失敗: {e}")
            return False
    
    def calculate_gross_margin(self):
        """
        計算毛利率
        
        Returns:
            pd.DataFrame: 包含毛利率的數據框
        """
        if self.df is None:
            print("請先載入產品數據")
            return None
        
        # 複製數據框避免修改原始數據
        df_analysis = self.df.copy()
        
        # 計算毛利率
        df_analysis['gross_margin'] = np.where(
            (df_analysis['price'].notna()) & (df_analysis['cost'].notna()) & (df_analysis['price'] > 0),
            (df_analysis['price'] - df_analysis['cost']) / df_analysis['price'],
            np.nan
        )
        
        # 轉換為百分比
        df_analysis['gross_margin_pct'] = df_analysis['gross_margin'] * 100
        
        # 計算建議售價
        df_analysis['suggested_price'] = np.where(
            (df_analysis['cost'].notna()) & (df_analysis['cost'] > 0),
            np.round(df_analysis['cost'] / (1 - self.target_margin), 2),
            np.nan
        )
        
        # 計算建議售價的毛利率
        df_analysis['suggested_margin'] = np.where(
            (df_analysis['suggested_price'].notna()) & (df_analysis['cost'].notna()),
            (df_analysis['suggested_price'] - df_analysis['cost']) / df_analysis['suggested_price'],
            np.nan
        )
        
        # 轉換為百分比
        df_analysis['suggested_margin_pct'] = df_analysis['suggested_margin'] * 100
        
        # 標記低毛利產品
        df_analysis['low_margin_flag'] = df_analysis['gross_margin_pct'] < (self.target_margin * 100)
        
        return df_analysis
    
    def generate_margin_report(self, output_file=None):
        """
        生成毛利報表
        
        Args:
            output_file (str): 輸出文件路徑
            
        Returns:
            pd.DataFrame: 分析結果
        """
        df_analysis = self.calculate_gross_margin()
        if df_analysis is None:
            return None
        
        # 生成報表摘要
        print("\n=== 毛利分析報表 ===")
        print(f"分析日期: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"目標毛利率: {self.target_margin * 100:.1f}%")
        print(f"總產品數: {len(df_analysis)}")
        
        # 統計有效數據
        valid_margin = df_analysis['gross_margin'].notna()
        print(f"有效毛利率數據: {valid_margin.sum()} 筆")
        
        if valid_margin.sum() > 0:
            print(f"平均毛利率: {df_analysis.loc[valid_margin, 'gross_margin_pct'].mean():.2f}%")
            print(f"最高毛利率: {df_analysis.loc[valid_margin, 'gross_margin_pct'].max():.2f}%")
            print(f"最低毛利率: {df_analysis.loc[valid_margin, 'gross_margin_pct'].min():.2f}%")
        
        # 低毛利產品統計
        low_margin_count = df_analysis['low_margin_flag'].sum()
        print(f"低毛利產品數: {low_margin_count} 筆")
        
        # 按類別統計
        if 'category' in df_analysis.columns:
            print("\n=== 按類別統計 ===")
            category_stats = df_analysis.groupby('category').agg({
                'gross_margin_pct': ['count', 'mean', 'min', 'max'],
                'low_margin_flag': 'sum'
            }).round(2)
            print(category_stats)
        
        # 低毛利產品明細
        if low_margin_count > 0:
            print(f"\n=== 低毛利產品明細 (毛利率 < {self.target_margin * 100:.1f}%) ===")
            low_margin_products = df_analysis[df_analysis['low_margin_flag']].copy()
            low_margin_products = low_margin_products.sort_values('gross_margin_pct')
            
            display_columns = ['sku', 'name', 'category', 'cost', 'price', 'gross_margin_pct', 'suggested_price']
            available_columns = [col for col in display_columns if col in low_margin_products.columns]
            
            print(low_margin_products[available_columns].to_string(index=False))
        
        # 保存報表
        if output_file:
            try:
                df_analysis.to_excel(output_file, index=False)
                print(f"\n毛利報表已保存至: {output_file}")
            except Exception as e:
                print(f"保存報表失敗: {e}")
        
        return df_analysis
    
    def generate_low_margin_email(self, df_analysis=None):
        """
        生成低毛利警示Email
        
        Args:
            df_analysis (pd.DataFrame): 分析結果數據框
            
        Returns:
            str: Email內容
        """
        if df_analysis is None:
            df_analysis = self.calculate_gross_margin()
            if df_analysis is None:
                return "無法生成Email：數據載入失敗"
        
        # 找出低毛利產品
        low_margin_products = df_analysis[df_analysis['low_margin_flag']].copy()
        
        if len(low_margin_products) == 0:
            email_content = f"""
主旨：毛利分析報告 - 無低毛利產品

內容：
各位好，

今日毛利分析報告：
✅ 所有產品毛利率均達到目標標準 ({self.target_margin * 100:.1f}%)
✅ 無需調整售價的產品

如有任何問題，請隨時聯繫。

謝謝！
            """
            return email_content
        
        # 生成低毛利產品表格
        table_html = self._generate_email_table(low_margin_products)
        
        # 生成Email內容
        today = datetime.now().strftime("%Y年%m月%d日")
        email_content = f"""
主旨：毛利分析報告 - 低毛利產品警示 ({today})

內容：
各位好，

以下是 {today} 毛利分析報告：

⚠️ **低毛利產品警示**
目標毛利率：{self.target_margin * 100:.1f}%
發現 {len(low_margin_products)} 個產品毛利率過低，需要關注：

{table_html}

📊 **建議行動：**
1. 檢視低毛利產品的成本結構
2. 考慮調整售價至建議售價
3. 評估是否有成本優化空間
4. 重新評估產品定位和市場策略

如有任何問題，請隨時聯繫。

謝謝！
        """
        
        return email_content
    
    def _generate_email_table(self, low_margin_products):
        """
        生成Email表格
        
        Args:
            low_margin_products (pd.DataFrame): 低毛利產品數據
            
        Returns:
            str: HTML表格
        """
        # 選擇要顯示的欄位
        display_columns = ['sku', 'name', 'category', 'cost', 'price', 'gross_margin_pct', 'suggested_price']
        available_columns = [col for col in display_columns if col in low_margin_products.columns]
        
        # 重新命名欄位為中文
        column_mapping = {
            'sku': 'SKU',
            'name': '產品名稱',
            'category': '類別',
            'cost': '成本',
            'price': '現行售價',
            'gross_margin_pct': '毛利率(%)',
            'suggested_price': '建議售價'
        }
        
        # 生成表格HTML
        table_html = """
        <table style="width: 100%; border-collapse: collapse; margin: 20px 0; font-family: Arial, sans-serif; font-size: 12px;">
            <thead>
                <tr style="background-color: #f8f9fa; border: 1px solid #dee2e6;">
        """
        
        # 表頭
        for col in available_columns:
            col_name = column_mapping.get(col, col)
            table_html += f'<th style="padding: 8px; text-align: center; border: 1px solid #dee2e6; font-weight: bold; color: #495057;">{col_name}</th>'
        
        table_html += """
                </tr>
            </thead>
            <tbody>
        """
        
        # 表格內容
        for _, row in low_margin_products.iterrows():
            table_html += '<tr style="border: 1px solid #dee2e6;">'
            
            for col in available_columns:
                value = row[col]
                
                # 格式化數值
                if col in ['cost', 'price', 'suggested_price']:
                    if pd.notna(value):
                        formatted_value = f"${value:,.2f}"
                    else:
                        formatted_value = "N/A"
                elif col == 'gross_margin_pct':
                    if pd.notna(value):
                        formatted_value = f"{value:.2f}%"
                        # 低毛利用紅色標示
                        cell_style = "color: #dc3545; font-weight: bold;"
                    else:
                        formatted_value = "N/A"
                        cell_style = ""
                else:
                    formatted_value = str(value) if pd.notna(value) else "N/A"
                    cell_style = ""
                
                # 應用樣式
                if 'cell_style' in locals():
                    table_html += f'<td style="padding: 8px; text-align: center; border: 1px solid #dee2e6; {cell_style}">{formatted_value}</td>'
                else:
                    table_html += f'<td style="padding: 8px; text-align: center; border: 1px solid #dee2e6;">{formatted_value}</td>'
            
            table_html += '</tr>'
        
        table_html += """
            </tbody>
        </table>
        """
        
        return table_html
    
    def generate_gmail_email(self, df_analysis=None):
        """
        生成適合Gmail的HTML Email
        
        Args:
            df_analysis (pd.DataFrame): 分析結果數據框
            
        Returns:
            str: Gmail格式的Email內容
        """
        if df_analysis is None:
            df_analysis = self.calculate_gross_margin()
            if df_analysis is None:
                return "無法生成Email：數據載入失敗"
        
        # 找出低毛利產品
        low_margin_products = df_analysis[df_analysis['low_margin_flag']].copy()
        
        if len(low_margin_products) == 0:
            email_content = f"""
<div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 800px; margin: 0 auto;">
    <h2 style="color: #28a745; border-bottom: 2px solid #28a745; padding-bottom: 10px;">
        ✅ 毛利分析報告 - 無低毛利產品
    </h2>
    
    <p>各位好，</p>
    
    <div style="background-color: #d4edda; border: 1px solid #c3e6cb; border-radius: 5px; padding: 15px; margin: 20px 0;">
        <h3 style="color: #155724; margin: 0;">今日毛利分析報告</h3>
        <p style="margin: 10px 0 0 0; color: #155724;">
            ✅ 所有產品毛利率均達到目標標準 ({self.target_margin * 100:.1f}%)<br>
            ✅ 無需調整售價的產品
        </p>
    </div>
    
    <p>如有任何問題，請隨時聯繫。</p>
    
    <p>謝謝！</p>
</div>
            """
            return email_content
        
        # 生成低毛利產品表格
        table_html = self._generate_email_table(low_margin_products)
        
        # 生成Gmail格式的Email內容
        today = datetime.now().strftime("%Y年%m月%d日")
        email_content = f"""
<div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 800px; margin: 0 auto;">
    <h2 style="color: #dc3545; border-bottom: 2px solid #dc3545; padding-bottom: 10px;">
        ⚠️ 毛利分析報告 - 低毛利產品警示
    </h2>
    
    <p>各位好，</p>
    
    <p>以下是 <strong>{today}</strong> 毛利分析報告：</p>
    
    <div style="background-color: #fff3cd; border: 1px solid #ffeaa7; border-radius: 5px; padding: 15px; margin: 20px 0;">
        <h3 style="color: #856404; margin: 0;">⚠️ 低毛利產品警示</h3>
        <p style="margin: 10px 0 0 0; color: #856404;">
            <strong>目標毛利率：{self.target_margin * 100:.1f}%</strong><br>
            發現 <strong>{len(low_margin_products)}</strong> 個產品毛利率過低，需要關注
        </p>
    </div>
    
    {table_html}
    
    <div style="background-color: #e7f3ff; border: 1px solid #b3d9ff; border-radius: 5px; padding: 15px; margin: 20px 0;">
        <h3 style="color: #004085; margin: 0;">📊 建議行動</h3>
        <ol style="margin: 10px 0 0 0; color: #004085;">
            <li>檢視低毛利產品的成本結構</li>
            <li>考慮調整售價至建議售價</li>
            <li>評估是否有成本優化空間</li>
            <li>重新評估產品定位和市場策略</li>
        </ol>
    </div>
    
    <p>如有任何問題，請隨時聯繫。</p>
    
    <p>謝謝！</p>
</div>
        """
        
        return email_content

def main():
    """主程式"""
    print("=== 毛利分析器 ===\n")
    
    # 設定目標毛利率
    target_margin = 0.40  # 40%
    
    # 初始化分析器
    analyzer = GrossMarginAnalyzer(target_margin=target_margin)
    
    # 載入產品數據
    input_file = "../clean/products_clean.xlsx"
    if not analyzer.load_products_data(input_file):
        return
    
    # 生成毛利報表
    print("\n正在生成毛利報表...")
    df_analysis = analyzer.generate_margin_report(output_file="gross_margin_report.xlsx")
    
    if df_analysis is None:
        print("生成報表失敗")
        return
    
    # 生成Email草稿
    print("\n正在生成Email草稿...")
    
    # 純文本版本
    email_draft = analyzer.generate_low_margin_email(df_analysis)
    with open("gross_margin_email_draft.txt", "w", encoding="utf-8") as f:
        f.write(email_draft)
    
    # Gmail HTML版本
    gmail_email = analyzer.generate_gmail_email(df_analysis)
    with open("gross_margin_gmail_email.html", "w", encoding="utf-8") as f:
        f.write(gmail_email)
    
    print("\n=== 完成！ ===")
    print("✅ 毛利報表：gross_margin_report.xlsx")
    print("✅ Email草稿：gross_margin_email_draft.txt")
    print("✅ Gmail Email：gross_margin_gmail_email.html")
    
    print("\n=== 使用說明 ===")
    print("1. 毛利報表包含所有產品的毛利率分析和建議售價")
    print("2. 低毛利產品會特別標記")
    print("3. Email草稿可直接複製到Gmail使用")
    print("4. HTML版本提供更好的視覺效果")

if __name__ == "__main__":
    main()
