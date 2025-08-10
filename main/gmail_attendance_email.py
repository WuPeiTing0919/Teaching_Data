import pandas as pd
from datetime import datetime

def generate_gmail_email_draft(late_attendees):
    """生成適合 Gmail 的 HTML Email 草稿"""
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
    
    # 生成 Gmail 友好的 HTML 表格
    table_html = """
    <table style="width: 100%; border-collapse: collapse; margin: 20px 0; font-family: Arial, sans-serif;">
        <thead>
            <tr style="background-color: #f8f9fa; border: 1px solid #dee2e6;">
                <th style="padding: 12px; text-align: left; border: 1px solid #dee2e6; font-weight: bold; color: #495057;">員工編號</th>
                <th style="padding: 12px; text-align: left; border: 1px solid #dee2e6; font-weight: bold; color: #495057;">姓名</th>
                <th style="padding: 12px; text-align: left; border: 1px solid #dee2e6; font-weight: bold; color: #495057;">日期</th>
                <th style="padding: 12px; text-align: left; border: 1px solid #dee2e6; font-weight: bold; color: #495057;">簽到時間</th>
                <th style="padding: 12px; text-align: left; border: 1px solid #dee2e6; font-weight: bold; color: #495057;">簽退時間</th>
                <th style="padding: 12px; text-align: left; border: 1px solid #dee2e6; font-weight: bold; color: #495057;">狀態</th>
                <th style="padding: 12px; text-align: left; border: 1px solid #dee2e6; font-weight: bold; color: #495057;">工作時數</th>
            </tr>
        </thead>
        <tbody>
    """
    
    for _, row in late_attendees.iterrows():
        # 為遲到狀態添加特殊樣式
        status_style = "color: #dc3545; font-weight: bold;" if row['status'] == 'Late' else ""
        
        table_html += f"""
            <tr style="border: 1px solid #dee2e6;">
                <td style="padding: 12px; border: 1px solid #dee2e6;">{row['emp_id']}</td>
                <td style="padding: 12px; border: 1px solid #dee2e6; font-weight: bold;">{row['name']}</td>
                <td style="padding: 12px; border: 1px solid #dee2e6;">{row['date']}</td>
                <td style="padding: 12px; border: 1px solid #dee2e6;">{row['check_in']}</td>
                <td style="padding: 12px; border: 1px solid #dee2e6;">{row['check_out']}</td>
                <td style="padding: 12px; border: 1px solid #dee2e6; {status_style}">{row['status']}</td>
                <td style="padding: 12px; border: 1px solid #dee2e6;">{row['work_hours']}</td>
            </tr>
        """
    
    table_html += """
        </tbody>
    </table>
    """
    
    # 生成完整的 Gmail Email 內容
    today = datetime.now().strftime("%Y年%m月%d日")
    email_content = f"""
主旨：今日考勤報告 - 遲到人員通知

內容：
<div style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
    <p>各位好，</p>
    
    <p>以下是 <strong>{today}</strong> 遲到人員名單：</p>
    
    {table_html}
    
    <p style="margin-top: 20px; padding: 15px; background-color: #fff3cd; border: 1px solid #ffeaa7; border-radius: 5px;">
        <strong>⚠️ 注意：</strong> 請相關人員注意準時到班，如有特殊情況請提前說明。
    </p>
    
    <p style="margin-top: 20px;">如有任何問題，請隨時聯繫。</p>
    
    <p>謝謝！</p>
</div>
    """
    
    return email_content

def generate_gmail_plain_text(late_attendees):
    """生成 Gmail 純文本版本（備用）"""
    if late_attendees is None or len(late_attendees) == 0:
        return "今日考勤報告：所有員工均準時到班，無遲到記錄。"
    
    # 生成純文本表格
    table_text = "員工編號\t姓名\t\t日期\t\t\t簽到時間\t簽退時間\t狀態\t工作時數\n"
    table_text += "-" * 80 + "\n"
    
    for _, row in late_attendees.iterrows():
        table_text += f"{row['emp_id']}\t\t{row['name']:<8}\t{row['date']}\t{row['check_in']}\t\t{row['check_out']}\t\t{row['status']}\t\t{row['work_hours']}\n"
    
    return table_text

def main():
    """主函數"""
    print("=== Gmail 考勤 Email 生成器 ===\n")
    
    # 讀取數據
    df = pd.read_excel("../clean/attendance_clean.xlsx")
    
    # 找出遲到人員
    late_attendees = df[df['status'] == 'Late'].copy()
    
    print(f"找到 {len(late_attendees)} 名遲到人員")
    
    # 生成 Gmail HTML Email 草稿
    gmail_html = generate_gmail_email_draft(late_attendees)
    
    # 生成純文本版本（備用）
    gmail_text = generate_gmail_plain_text(late_attendees)
    
    # 保存 HTML 版本
    html_file = "gmail_attendance_email.html"
    with open(html_file, "w", encoding="utf-8") as f:
        f.write(gmail_html)
    
    # 保存純文本版本
    text_file = "gmail_attendance_email.txt"
    with open(text_file, "w", encoding="utf-8") as f:
        f.write(gmail_text)
    
    print("\n=== Gmail Email 草稿已生成 ===")
    print(f"HTML 版本：{html_file}")
    print(f"純文本版本：{text_file}")
    print("\n=== 使用說明 ===")
    print("1. 打開 Gmail")
    print("2. 點擊「撰寫」")
    print("3. 複製 HTML 版本的內容到郵件正文")
    print("4. 或者複製純文本版本（Gmail 會自動轉換格式）")
    print("\n=== HTML 版本預覽 ===")
    print(gmail_html)

if __name__ == "__main__":
    main()
