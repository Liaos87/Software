import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import random
import time

class LotteryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Lottery Program")
        self.root.geometry("480x320")  # Set window size
        self.usernames = []
        self.lucky_user = None

        # File selection button
        self.btn_open_file = tk.Button(root, text="Open Excel File", font=("Arial", 14), command=self.open_file)
        self.btn_open_file.pack(pady=20)

        # Start lottery button
        self.btn_start = tk.Button(root, text="Start Lottery", font=("Arial", 14), command=self.start_lottery, state=tk.DISABLED)
        self.btn_start.pack(pady=20)

        # Display result label
        self.label_result = tk.Label(root, text="", font=("Arial", 24), fg="red")
        self.label_result.pack(pady=40)

    def open_file(self):
        # Open file dialog to select Excel file
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return

        try:
            # Read the Excel file and extract the username column
            df = pd.read_excel(file_path)
            if "name" in df.columns or "用户名" in df.columns:
                column_name = "name" if "name" in df.columns else "用户名"
                self.usernames = df[column_name].dropna().tolist()
                if self.usernames:
                    messagebox.showinfo("Success", f"Successfully loaded {len(self.usernames)} users!")
                    self.btn_start.config(state=tk.NORMAL)
                else:
                    messagebox.showwarning("Warning", "No valid usernames found!")
            else:
                messagebox.showerror("Error", "The Excel file does not contain a 'name' or '用户名' column!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read the file: {str(e)}")

    def start_lottery(self):
        if not self.usernames:
            messagebox.showwarning("Warning", "No user information loaded!")
            return

        def update_result():
            for _ in range(30):  # Simulate animation, randomly display 30 times
                self.lucky_user = random.choice(self.usernames)
                self.label_result.config(text=self.lucky_user)
                self.root.update()
                time.sleep(0.1)  # Animation interval
            messagebox.showinfo("Congratulations", f"The lucky user is: {self.lucky_user}")

        self.label_result.config(text="Drawing...")
        self.root.after(100, update_result)  # Delay to update the result

if __name__ == "__main__":
    root = tk.Tk()
    app = LotteryApp(root)
    root.mainloop()
