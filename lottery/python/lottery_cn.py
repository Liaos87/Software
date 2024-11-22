import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import random
import time

class LotteryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("抽奖程序")
        self.root.geometry("480x320")  # 设置窗口大小
        self.usernames = []
        self.lucky_user = None

        # 文件选择按钮
        self.btn_open_file = tk.Button(root, text="打开Excel文件", font=("Arial", 14), command=self.open_file)
        self.btn_open_file.pack(pady=20)

        # 抽奖按钮
        self.btn_start = tk.Button(root, text="开始抽奖", font=("Arial", 14), command=self.start_lottery, state=tk.DISABLED)
        self.btn_start.pack(pady=20)

        # 显示结果标签
        self.label_result = tk.Label(root, text="", font=("Arial", 24), fg="red")
        self.label_result.pack(pady=40)

    def open_file(self):
        # 打开文件对话框选择Excel文件
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return

        try:
            # 读取Excel文件并提取用户名字段
            df = pd.read_excel(file_path)
            if "name" in df.columns or "用户名" in df.columns:
                column_name = "name" if "name" in df.columns else "用户名"
                self.usernames = df[column_name].dropna().tolist()
                if self.usernames:
                    messagebox.showinfo("成功", f"成功加载 {len(self.usernames)} 个用户！")
                    self.btn_start.config(state=tk.NORMAL)
                else:
                    messagebox.showwarning("警告", "未找到有效的用户名信息！")
            else:
                messagebox.showerror("错误", "Excel 文件中未包含 'name' 或 '用户名' 字段！")
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败：{str(e)}")

    def start_lottery(self):
        if not self.usernames:
            messagebox.showwarning("警告", "没有加载用户信息！")
            return

        def update_result():
            for _ in range(30):  # 模拟动画效果，随机展示30次
                self.lucky_user = random.choice(self.usernames)
                self.label_result.config(text=self.lucky_user)
                self.root.update()
                time.sleep(0.1)  # 动画间隔时间
            messagebox.showinfo("恭喜", f"幸运用户是：{self.lucky_user}")

        self.label_result.config(text="抽奖中...")
        self.root.after(100, update_result)  # 延迟更新结果

# 打包命令 pyinstaller --onefile --noconsole lottery_cn.py

if __name__ == "__main__":
    root = tk.Tk()
    app = LotteryApp(root)
    root.mainloop()
