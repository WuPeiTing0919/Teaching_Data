import pandas as pd
from datetime import datetime

def generate_text_email_draft(late_attendees):
    """生成純文本格式的 Email 草稿"""
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
    
    # 生成純文本表格
    table_text = "員工編號\t姓名\t\t日期\t\t\t簽到時間\t簽退時間\t狀態\t工作時數\n"
    table_text += "-" * 80 + "\n"
    
    for _, row in late_attendees.iterrows():
        table_text += f"{row['emp_id']}\t\t{row['name']:<8}\t{row['date']}\t{row['check_in']}\t\t{row['check_out']}\t\t{row['status']}\t\t{row['work_hours']}\n"
    
    # 生成 Email 內容
    today = datetime.now().strftime("%Y年%m月%d日")
    email_content = f"""
主旨：今日考勤報告 - 遲到人員通知

內容：
各位好，

以下是 {today} 遲到人員名單：

{table_text}

請相關人員注意準時到班，如有特殊情況請提前說明。

如有任何問題，請隨時聯繫。

謝謝！
    """
    
    return email_content

def main():
    """主函數"""
    # 讀取數據
    df = pd.read_excel("../clean/attendance_clean.xlsx")
    
    # 找出遲到人員
    late_attendees = df[df['status'] == 'Late'].copy()
    
    # 生成 Email 草稿
    email_draft = generate_text_email_draft(late_attendees)
    
    # 保存到文件
    output_file = "attendance_email_draft_clean.txt"
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(email_draft)
    
    print("=== 考勤 Email 草稿（純文本格式）===")
    print(email_draft)
    print(f"\nEmail 草稿已保存到: {output_file}")

if __name__ == "__main__":
    main()
