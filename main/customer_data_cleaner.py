import pandas as pd
import re
from datetime import datetime
import numpy as np

def clean_customer_data(input_file, output_file):
    """
    清洗客戶資料，處理各種資料品質問題
    """
    # 讀取資料
    df = pd.read_excel(input_file)
    print(f"原始資料筆數: {len(df)}")
    
    # 1. 去前後空白與 Title Case
    df['Customer ID'] = df['Customer ID'].str.strip()  # 先處理 Customer ID 的空白
    df['Name'] = df['Name'].str.strip().str.title()
    df['Email'] = df['Email'].str.strip()
    df['City'] = df['City'].str.strip()
    
    print(f"處理空白後 Customer ID: {df['Customer ID'].tolist()}")
    
    # 2. Email 合法才保留，不合法清除 Email 資料
    def is_valid_email(email):
        if pd.isna(email):
            return False
        email = str(email).strip()
        # 簡單的 email 驗證規則
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, email))
    
    # 標記無效的 email
    invalid_emails = ~df['Email'].apply(is_valid_email)
    df.loc[invalid_emails, 'Email'] = np.nan
    print(f"清除無效 Email 筆數: {invalid_emails.sum()}")
    
    # 3. 電話格式標準化
    def standardize_phone(phone):
        if pd.isna(phone):
            return np.nan
        
        phone = str(phone).strip()
        
        # 移除所有空白和特殊符號
        phone_clean = re.sub(r'[\s\-\(\)]', '', phone)
        
        # 處理台灣手機號碼 (09xxxxxxxx)
        if phone_clean.startswith('8869') and len(phone_clean) == 12:
            return '09' + phone_clean[4:]
        elif phone_clean.startswith('09') and len(phone_clean) == 10:
            return phone_clean
        # 處理市話 (02xxxxxxxx)
        elif phone_clean.startswith('02') and len(phone_clean) == 10:
            return f"({phone_clean[:2]}){phone_clean[2:6]}-{phone_clean[6:]}"
        else:
            # 如果不符合標準格式，設為 NaN
            return np.nan
    
    df['Phone'] = df['Phone'].apply(standardize_phone)
    
    # 4. 加入 valid 欄位標記
    df['Valid'] = True
    
    # 標記無效的資料
    df.loc[df['Email'].isna(), 'Valid'] = False
    df.loc[df['Phone'].isna(), 'Valid'] = False
    
    # 5. 日期轉 YYYY-MM-DD
    def standardize_date(date_str):
        if pd.isna(date_str):
            return np.nan
        
        date_str = str(date_str).strip()
        
        # 處理 "8月1日2025年" 格式
        if '月' in date_str and '日' in date_str and '年' in date_str:
            # 提取年月日
            year_match = re.search(r'(\d{4})年', date_str)
            month_match = re.search(r'(\d+)月', date_str)
            day_match = re.search(r'(\d+)日', date_str)
            
            if year_match and month_match and day_match:
                year = year_match.group(1)
                month = month_match.group(1).zfill(2)
                day = day_match.group(1).zfill(2)
                return f"{year}-{month}-{day}"
        
        # 處理 "2025/8/9" 格式
        elif '/' in date_str:
            try:
                date_obj = datetime.strptime(date_str, '%Y/%m/%d')
                return date_obj.strftime('%Y-%m-%d')
            except:
                pass
        
        # 處理 "09-08-2025" 格式
        elif '-' in date_str and len(date_str.split('-')) == 3:
            try:
                date_obj = datetime.strptime(date_str, '%d-%m-%Y')
                return date_obj.strftime('%Y-%m-%d')
            except:
                pass
        
        # 處理 "2025-08-05" 格式
        elif '-' in date_str and len(date_str.split('-')) == 3:
            try:
                date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                return date_obj.strftime('%Y-%m-%d')
            except:
                pass
        
        return np.nan
    
    df['Join Date'] = df['Join Date'].apply(standardize_date)
    
    # 6. 城市做字典映射
    city_mapping = {
        'Taipei': 'Taipei',
        'taipei': 'Taipei',
        'TPE': 'Taipei',
        '台北': 'Taipei',
        'Taoyuan': 'Taoyuan',
        'taoyuan': 'Taoyuan',
        '桃園': 'Taoyuan'
    }
    
    df['City'] = df['City'].map(city_mapping).fillna('Taipei')  # 預設為 Taipei
    
    # 7. 轉金額為數值（移除逗號/NT$，破折號當成空值）
    def clean_amount(amount):
        if pd.isna(amount):
            return np.nan
        
        amount_str = str(amount).strip()
        
        # 處理破折號
        if amount_str == '—' or amount_str == '-':
            return np.nan
        
        # 移除 NT$ 和逗號
        amount_str = re.sub(r'NT\$', '', amount_str)
        amount_str = re.sub(r',', '', amount_str)
        
        try:
            return float(amount_str)
        except:
            return np.nan
    
    df['Spend (NT$)'] = df['Spend (NT$)'].apply(clean_amount)
    
    # 8. 去除重複（以 Customer ID 為主要鍵）
    # 先處理 Customer ID 的大小寫問題，統一為大寫
    df['Customer ID'] = df['Customer ID'].str.upper()
    
    print(f"處理前 Customer ID 統計:")
    print(df['Customer ID'].value_counts())
    
    # 對於重複的 Customer ID，優先保留有 Email 的記錄，如果都有 Email 則保留第一筆
    df_sorted = df.sort_values(['Customer ID', 'Email'], na_position='last')
    
    # 去除重複，保留第一個出現的記錄
    df_cleaned = df_sorted.drop_duplicates(subset=['Customer ID'], keep='first')
    print(f"去除重複後筆數: {len(df_cleaned)}")
    
    print(f"處理後 Customer ID 統計:")
    print(df_cleaned['Customer ID'].value_counts())
    
    # 重新整理欄位順序
    columns_order = ['Customer ID', 'Name', 'Email', 'Phone', 'Join Date', 'City', 'Spend (NT$)', 'Valid']
    df_cleaned = df_cleaned[columns_order]
    
    # 儲存清洗後的資料
    df_cleaned.to_excel(output_file, index=False)
    print(f"清洗完成！資料已儲存至: {output_file}")
    
    # 顯示清洗結果摘要
    print("\n=== 清洗結果摘要 ===")
    print(f"有效 Email: {df_cleaned['Email'].notna().sum()}")
    print(f"有效電話: {df_cleaned['Phone'].notna().sum()}")
    print(f"有效日期: {df_cleaned['Join Date'].notna().sum()}")
    print(f"有效金額: {df_cleaned['Spend (NT$)'].notna().sum()}")
    print(f"有效記錄: {df_cleaned['Valid'].sum()}")
    
    return df_cleaned

if __name__ == "__main__":
    input_file = "dirty/1.customers_dirty.xlsx"
    output_file = "clean/customers_clean.xlsx"
    
    try:
        cleaned_df = clean_customer_data(input_file, output_file)
        print("\n=== 清洗後的資料 ===")
        print(cleaned_df.to_string())
    except Exception as e:
        print(f"錯誤: {e}")
