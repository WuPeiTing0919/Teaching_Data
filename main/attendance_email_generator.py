import pandas as pd
from datetime import datetime
import os

def read_attendance_data():
    """讀取考勤數據"""
    try:
        # 讀取 Excel 文件
        file_path = "../clean/attendance_clean.xlsx"
        df = pd.read_excel(file_path)
        print("成功讀取考勤數據")
        print(f"數據形狀: {df.shape}")
        print(f"列名: {list(df.columns)}")
        print("\n前5行數據:")
        print(df.head())
        return df
    except Exception as e:
        print(f"讀取文件時發生錯誤: {e}")
        return None

def find_late_attendees(df):
    """找出遲到次數 ≥ 1 的人"""
    try:
        # 檢查列名
        print(f"\n可用的列名: {list(df.columns)}")
        
        # 檢查 status 列的值
        if 'status' in df.columns:
            print(f"\nStatus 列的唯一值: {df['status'].unique()}")
            
            # 找出狀態為 'Late' 的記錄
            late_attendees = df[df['status'] == 'Late'].copy()
            late_column = 'status'
            
            print(f"\n遲到人員記錄數: {len(late_attendees)}")
            
            return late_attendees, late_column
        
        # 如果沒有 status 列，嘗試其他方法
        late_columns = [col for col in df.columns if any(keyword in col.lower() for keyword in ['late', '遲到', '次數', 'count'])]
        print(f"可能的遲到相關列: {late_columns}")
        
        if not late_columns:
            # 如果沒有找到遲到列，嘗試其他可能的列名
            print("未找到遲到相關列，嘗試分析所有數值列...")
            numeric_columns = df.select_dtypes(include=['number']).columns.tolist()
            print(f"數值列: {numeric_columns}")
            
            # 假設第一列是姓名，其他列是各種考勤指標
            if len(numeric_columns) > 0:
                # 使用第一個數值列作為遲到次數
                late_column = numeric_columns[0]
                print(f"使用列 '{late_column}' 作為遲到次數")
            else:
                print("未找到數值列")
                return None, None
        else:
            late_column = late_columns[0]
        
        # 找出遲到次數 ≥ 1 的記錄
        late_attendees = df[df[late_column] >= 1].copy()
        
        print(f"\n遲到次數 ≥ 1 的記錄數: {len(late_attendees)}")
        
        return late_attendees, late_column
        
    except Exception as e:
        print(f"分析遲到數據時發生錯誤: {e}")
        return None, None

def generate_email_draft(late_attendees, late_column):
    """生成 Email 草稿"""
    if late_attendees is None or len(late_attendees) == 0:
        email_content = """
主旨：今日考勤報告 - 無遲到記錄

內容：
各位好，

今日考勤報告：所有員工均準時到班，無遲到記錄。

如有任何問題，請隨時聯繫。

謝謝！
        """
        return email_content
    
    # 生成表格 HTML
    table_html = late_attendees.to_html(index=False, classes='table table-striped')
    
    # 生成 Email 內容
    today = datetime.now().strftime("%Y年%m月%d日")
    email_content = f"""
主旨：今日考勤報告 - 遲到人員通知

內容：
各位好，

以下是 {today} 遲到次數 ≥ 1 的人員名單：

{table_html}

請相關人員注意準時到班，如有特殊情況請提前說明。

如有任何問題，請隨時聯繫。

謝謝！
    """
    
    return email_content

def main():
    """主函數"""
    print("=== 考勤 Email 生成器 ===\n")
    
    # 讀取數據
    df = read_attendance_data()
    if df is None:
        return
    
    # 找出遲到人員
    late_attendees, late_column = find_late_attendees(df)
    
    # 生成 Email 草稿
    email_draft = generate_email_draft(late_attendees, late_column)
    
    print("\n=== Email 草稿 ===")
    print(email_draft)
    
    # 保存到文件
    output_file = "attendance_email_draft.txt"
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(email_draft)
    
    print(f"\nEmail 草稿已保存到: {output_file}")

if __name__ == "__main__":
    main()
