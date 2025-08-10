# 這是一個簡單的點餐程式
# 讓使用者輸入想吃的食物名稱
# 然後輸出「您點了 XX，請稍候」的訊息

# 定義主函數
def main():
    # 提示使用者輸入食物名稱
    food = input("請輸入您想點的食物名稱: ")
    
    # 輸出點餐確認訊息
    print(f"您點了 {food}，請稍候")
    
    # 返回結果，方便展示
    return f"您點了 {food}，請稍候"

# 當程式被直接執行時，呼叫 main 函數
if __name__ == "__main__":
    result = main()
    
    # 這裡可以添加 Markdown 格式的輸出，但在終端中不會有效果
    # 以下僅作為示範
    print("\n結果的 Markdown 格式：")
    print(f"**{result}**")