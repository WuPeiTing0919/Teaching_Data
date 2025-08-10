import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import re
import warnings
warnings.filterwarnings('ignore')

def clean_emp_id(emp_id):
    """正規化員工ID：去空白，統一 E-xx 形式"""
    if pd.isna(emp_id):
        return emp_id
    
    # 轉為字串並去空白
    emp_id = str(emp_id).strip()
    
    # 提取數字部分
    match = re.search(r'(\d+)', emp_id)
    if match:
        num = match.group(1)
        # 統一格式為 E-xx
        return f"E-{int(num):02d}"
    
    return emp_id

def clean_name(name):
    """標準化姓名：去空白 + Title Case"""
    if pd.isna(name):
        return name
    
    # 轉為字串並去空白
    name = str(name).strip()
    
    # 轉為 Title Case（首字母大寫，其餘小寫）
    return name.title()

def clean_date(date_str):
    """轉換日期為 YYYY-MM-DD 格式"""
    if pd.isna(date_str):
        return date_str
    
    date_str = str(date_str).strip()
    
    # 處理中文日期格式
    if '月' in date_str and '年' in date_str:
        # 例如：8月1日2025年
        match = re.search(r'(\d+)月(\d+)日(\d+)年', date_str)
        if match:
            month, day, year = match.groups()
            return f"{year}-{int(month):02d}-{int(day):02d}"
    
    # 處理其他格式
    try:
        # 嘗試多種日期格式
        for fmt in ['%Y/%m/%d', '%d-%m-%Y', '%Y.%m.%d', '%Y-%m-%d']:
            try:
                parsed_date = datetime.strptime(date_str, fmt)
                return parsed_date.strftime('%Y-%m-%d')
            except ValueError:
                continue
        
        # 如果都失敗，返回原值
        return date_str
    except:
        return date_str

def clean_time(time_str):
    """統一時間格式為 HH:MM"""
    if pd.isna(time_str) or str(time_str).strip() in ['', '—', 'nan']:
        return time_str
    
    time_str = str(time_str).strip()
    
    # 處理全形冒號
    time_str = time_str.replace('：', ':')
    
    # 處理 0905 格式
    if re.match(r'^\d{4}$', time_str):
        return f"{time_str[:2]}:{time_str[2:]}"
    
    # 處理 AM/PM 格式
    if 'AM' in time_str.upper() or 'PM' in time_str.upper():
        try:
            # 移除 AM/PM 並解析
            time_part = re.sub(r'[AP]M', '', time_str, flags=re.IGNORECASE).strip()
            parsed_time = datetime.strptime(time_part, '%H:%M')
            
            # 如果是 PM 且不是 12 點，加 12 小時
            if 'PM' in time_str.upper() and parsed_time.hour != 12:
                parsed_time = parsed_time.replace(hour=parsed_time.hour + 12)
            # 如果是 AM 且是 12 點，改為 0 點
            elif 'AM' in time_str.upper() and parsed_time.hour == 12:
                parsed_time = parsed_time.replace(hour=0)
            
            return parsed_time.strftime('%H:%M')
        except:
            return time_str
    
    # 處理標準格式 HH:MM
    if re.match(r'^\d{1,2}:\d{2}$', time_str):
        # 確保小時是兩位數
        parts = time_str.split(':')
        return f"{int(parts[0]):02d}:{parts[1]}"
    
    return time_str

def clean_status(status):
    """標準化狀態為 Late/On time"""
    if pd.isna(status):
        return status
    
    status = str(status).strip().lower()
    
    # 處理中文「遲到」
    if status in ['遲到', 'late']:
        return 'Late'
    elif status in ['on time', 'ontime']:
        return 'On time'
    else:
        return 'On time'  # 預設為準時

def calculate_work_hours(check_in, check_out):
    """計算工作時數（能算的才算）"""
    if pd.isna(check_in) or pd.isna(check_out):
        return np.nan
    
    try:
        # 轉換為時間物件
        if isinstance(check_in, str):
            check_in = datetime.strptime(check_in, '%H:%M').time()
        if isinstance(check_out, str):
            check_out = datetime.strptime(check_out, '%H:%M').time()
        
        # 計算時數差
        start = datetime.combine(datetime.today(), check_in)
        end = datetime.combine(datetime.today(), check_out)
        
        # 如果下班時間小於上班時間，表示跨日
        if end < start:
            end += timedelta(days=1)
        
        duration = end - start
        hours = duration.total_seconds() / 3600
        
        return round(hours, 2)
    except:
        return np.nan

def clean_attendance_data(input_file, output_file):
    """主要清洗函數"""
    print("開始清洗出勤資料...")
    
    # 讀取資料
    try:
        df = pd.read_excel(input_file)
        print(f"原始資料形狀: {df.shape}")
        print(f"原始欄位: {list(df.columns)}")
    except Exception as e:
        print(f"讀取檔案失敗: {e}")
        return
    
    # 顯示原始資料的前幾筆
    print("\n原始資料前5筆:")
    print(df.head())
    
    # 複製資料進行清洗
    df_clean = df.copy()
    
    # 1. 清洗員工ID
    print("\n1. 清洗員工ID...")
    df_clean['emp_id'] = df_clean['emp_id'].apply(clean_emp_id)
    
    # 2. 清洗姓名
    print("2. 清洗姓名...")
    df_clean['name'] = df_clean['name'].apply(clean_name)
    
    # 3. 清洗日期
    print("3. 清洗日期...")
    df_clean['date'] = df_clean['date'].apply(clean_date)
    
    # 4. 清洗時間
    print("4. 清洗時間...")
    df_clean['check_in'] = df_clean['check_in'].apply(clean_time)
    df_clean['check_out'] = df_clean['check_out'].apply(clean_time)
    
    # 5. 清洗狀態
    print("5. 清洗狀態...")
    df_clean['status'] = df_clean['status'].apply(clean_status)
    
    # 6. 新增工作時數欄位
    print("6. 計算工作時數...")
    df_clean['work_hours'] = df_clean.apply(
        lambda row: calculate_work_hours(row['check_in'], row['check_out']), axis=1
    )
    
    # 7. 過濾無效資料（日期為空或無效的記錄）
    print("7. 過濾無效資料...")
    initial_count = len(df_clean)
    df_clean = df_clean.dropna(subset=['date'])
    df_clean = df_clean[df_clean['date'] != '']  # 移除空字串
    print(f"過濾無效資料前: {initial_count} 筆，過濾後: {len(df_clean)} 筆")
    
    # 8. 去除重複（以 emp_id+date 為準）
    print("8. 去除重複記錄...")
    initial_count = len(df_clean)
    df_clean = df_clean.drop_duplicates(subset=['emp_id', 'date'], keep='first')
    final_count = len(df_clean)
    print(f"去除重複前: {initial_count} 筆，去除重複後: {final_count} 筆")
    
    # 顯示清洗後的資料
    print("\n清洗後資料:")
    print(df_clean)
    
    # 9. 創建遲到次數排行榜
    print("\n9. 創建遲到次數排行榜...")
    late_ranking = df_clean[df_clean['status'] == 'Late'].groupby(['emp_id', 'name']).size().reset_index(name='late_count')
    late_ranking = late_ranking.sort_values('late_count', ascending=False)
    
    print("遲到次數排行榜:")
    print(late_ranking)
    
    # 10. 儲存結果
    print(f"\n10. 儲存清洗後資料到 {output_file}...")
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # 主要清洗後的資料
        df_clean.to_excel(writer, sheet_name='清洗後資料', index=False)
        
        # 遲到次數排行榜
        late_ranking.to_excel(writer, sheet_name='遲到次數排行榜', index=False)
    
    print("資料清洗完成！")
    print(f"輸出檔案: {output_file}")
    print(f"包含兩個工作表: '清洗後資料' 和 '遲到次數排行榜'")
    
    return df_clean, late_ranking

if __name__ == "__main__":
    # 設定檔案路徑
    input_file = "dirty/4.attendance_dirty.xlsx"
    output_file = "clean/attendance_clean.xlsx"
    
    # 執行清洗
    clean_data, ranking = clean_attendance_data(input_file, output_file)
